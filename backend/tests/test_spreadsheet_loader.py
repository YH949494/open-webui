"""
Tests for spreadsheet loading in the document Loader.

Spreadsheet files (xlsx, xls, csv, ods) are now processed through the normal
extraction pipeline so their content is available for RAG retrieval. Previously
they were routed to a metadata-only loader, but that prevented the model from
seeing actual cell data.

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
ExcelLoader = main.ExcelLoader
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


def test_xlsx_routes_to_excel_loader():
    # xlsx should now use ExcelLoader (or UnstructuredExcelLoader), not metadata-only
    loader = Loader(engine='')
    got = loader._get_loader(
        'big.xlsx',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        '/tmp/big.xlsx',
    )
    assert not isinstance(got, SpreadsheetMetadataOnlyLoader)


def test_csv_routes_to_normal_loader():
    # CSV should use the standard CSVLoader, not metadata-only
    loader = Loader(engine='')
    got = loader._get_loader('big.csv', 'text/csv', '/tmp/big.csv')
    assert not isinstance(got, SpreadsheetMetadataOnlyLoader)


def test_xlsx_with_tika_engine():
    # With tika engine configured, xlsx should use TikaLoader (not metadata-only)
    loader = Loader(
        engine='tika',
        TIKA_SERVER_URL='http://tika:9998',
    )
    got = loader._get_loader('big.xlsx', None, '/tmp/big.xlsx')
    assert not isinstance(got, SpreadsheetMetadataOnlyLoader)


def test_pdf_still_routes_to_normal_loader():
    loader = Loader(engine='')
    got = loader._get_loader('doc.pdf', 'application/pdf', '/tmp/doc.pdf')
    assert not isinstance(got, SpreadsheetMetadataOnlyLoader)


def test_txt_still_routes_to_normal_loader():
    loader = Loader(engine='')
    got = loader._get_loader('notes.txt', 'text/plain', '/tmp/notes.txt')
    assert not isinstance(got, SpreadsheetMetadataOnlyLoader)


def test_excel_loader_includes_cell_data():
    # ExcelLoader should extract actual cell data, not just metadata.
    pd = pytest.importorskip('pandas')
    pytest.importorskip('openpyxl')

    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, 'sample.xlsx')
        df = pd.DataFrame({'secret_col': ['UNIQUE_CELL_VALUE_12345']})
        with pd.ExcelWriter(path) as writer:
            df.to_excel(writer, sheet_name='MySheet', index=False)

        docs = ExcelLoader(file_path=path).load()

    assert len(docs) >= 1
    content = '\n'.join(doc.page_content for doc in docs)
    assert 'UNIQUE_CELL_VALUE_12345' in content


def test_metadata_only_loader_never_fails_on_bad_file():
    # SpreadsheetMetadataOnlyLoader class still exists and works for direct use.
    docs = SpreadsheetMetadataOnlyLoader(
        file_path='/nonexistent/path/missing.xlsx',
        filename='missing.xlsx',
        file_content_type='application/vnd.ms-excel',
    ).load()
    assert len(docs) == 1
    assert 'Spreadsheet file detected' in docs[0].page_content
