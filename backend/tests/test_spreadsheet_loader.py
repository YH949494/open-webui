"""
Tests for the spreadsheet RAG guard in the document Loader.

Large spreadsheets used to be fully extracted (every cell turned into text),
producing payloads that exceed the content-filter maximum size. The Loader now
routes spreadsheet files to a metadata-only loader by default.

These tests import the real Loader and are skipped automatically when the
backend's heavy document-loading dependencies are not installed (e.g. in a
minimal local checkout). They run for real in CI where requirements are present.

Run from repo root:
    cd backend && python3 -m pytest tests/test_spreadsheet_loader.py -v
"""

import os
import tempfile

import pytest

# Skip the whole module if the heavy loader dependencies aren't available.
pytest.importorskip('langchain_community')
pytest.importorskip('ftfy')

main = pytest.importorskip('open_webui.retrieval.loaders.main')

Loader = main.Loader
SpreadsheetMetadataOnlyLoader = main.SpreadsheetMetadataOnlyLoader
is_spreadsheet_file = main.is_spreadsheet_file


def test_is_spreadsheet_file_by_extension():
    assert is_spreadsheet_file('data.xlsx', None)
    assert is_spreadsheet_file('data.xls', None)
    assert is_spreadsheet_file('data.csv', None)
    assert is_spreadsheet_file('data.ods', None)
    assert not is_spreadsheet_file('report.pdf', None)
    assert not is_spreadsheet_file('notes.txt', None)


def test_is_spreadsheet_file_by_mime_type():
    assert is_spreadsheet_file('blob', 'text/csv')
    assert is_spreadsheet_file(
        'blob',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    # MIME with charset parameter still matches
    assert is_spreadsheet_file('blob', 'text/csv; charset=utf-8')
    assert not is_spreadsheet_file('blob', 'application/pdf')


def test_xlsx_routes_to_metadata_only_loader():
    loader = Loader(engine='')
    got = loader._get_loader(
        'big.xlsx',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        '/tmp/big.xlsx',
    )
    assert isinstance(got, SpreadsheetMetadataOnlyLoader)


def test_csv_routes_to_metadata_only_loader():
    loader = Loader(engine='')
    got = loader._get_loader('big.csv', 'text/csv', '/tmp/big.csv')
    assert isinstance(got, SpreadsheetMetadataOnlyLoader)


def test_spreadsheet_guard_runs_before_extraction_engines():
    # Even with an external engine configured, spreadsheets must be guarded first.
    loader = Loader(
        engine='tika',
        TIKA_SERVER_URL='http://tika:9998',
    )
    got = loader._get_loader('big.xlsx', None, '/tmp/big.xlsx')
    assert isinstance(got, SpreadsheetMetadataOnlyLoader)


def test_pdf_still_routes_to_normal_loader():
    loader = Loader(engine='')
    got = loader._get_loader('doc.pdf', 'application/pdf', '/tmp/doc.pdf')
    assert not isinstance(got, SpreadsheetMetadataOnlyLoader)


def test_txt_still_routes_to_normal_loader():
    loader = Loader(engine='')
    got = loader._get_loader('notes.txt', 'text/plain', '/tmp/notes.txt')
    assert not isinstance(got, SpreadsheetMetadataOnlyLoader)


def test_metadata_only_content_excludes_cell_data():
    # Build a small real .xlsx with recognizable cell content.
    pd = pytest.importorskip('pandas')
    pytest.importorskip('openpyxl')

    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, 'sample.xlsx')
        df = pd.DataFrame({'secret_col': ['UNIQUE_CELL_VALUE_12345']})
        with pd.ExcelWriter(path) as writer:
            df.to_excel(writer, sheet_name='MySheet', index=False)

        docs = SpreadsheetMetadataOnlyLoader(
            file_path=path,
            filename='sample.xlsx',
            file_content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        ).load()

    assert len(docs) == 1
    content = docs[0].page_content
    # Metadata-only: cell values must NOT be present...
    assert 'UNIQUE_CELL_VALUE_12345' not in content
    # ...but sheet names (cheap to read) and the guidance message should be.
    assert 'MySheet' in content
    assert 'Spreadsheet file detected' in content
    assert 'Code Interpreter' in content


def test_metadata_only_loader_never_fails_on_bad_file():
    # A missing/corrupt file must still yield a metadata-only Document, not raise.
    docs = SpreadsheetMetadataOnlyLoader(
        file_path='/nonexistent/path/missing.xlsx',
        filename='missing.xlsx',
        file_content_type='application/vnd.ms-excel',
    ).load()
    assert len(docs) == 1
    assert 'Spreadsheet file detected' in docs[0].page_content
