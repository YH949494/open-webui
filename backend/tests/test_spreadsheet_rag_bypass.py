from pathlib import Path

from open_webui.utils.file_types import get_file_extension, is_spreadsheet_file


REPO_ROOT = Path(__file__).resolve().parents[2]


def test_spreadsheet_extension_detection():
    assert get_file_extension('report.XLSX') == 'xlsx'
    assert get_file_extension('report.csv') == 'csv'
    assert is_spreadsheet_file('report.xlsx')
    assert is_spreadsheet_file('report.xls')
    assert is_spreadsheet_file('report.csv')
    assert not is_spreadsheet_file('report.pdf')
    assert not is_spreadsheet_file('report.docx')
    assert not is_spreadsheet_file('report.txt')
    assert not is_spreadsheet_file('report.md')


def test_spreadsheet_upload_does_not_bypass_allowed_extensions():
    files_router = REPO_ROOT / 'backend' / 'open_webui' / 'routers' / 'files.py'
    source = files_router.read_text()

    allowed_extension_check = source[
        source.index('if process and request.app.state.config.ALLOWED_FILE_EXTENSIONS:')
        : source.index('# replace filename with uuid')
    ]

    assert 'file_extension not in request.app.state.config.ALLOWED_FILE_EXTENSIONS' in allowed_extension_check
    assert 'skip_spreadsheet_rag' not in allowed_extension_check


def test_spreadsheet_bypass_guards_are_present():
    changed_files = [
        REPO_ROOT / 'backend' / 'open_webui' / 'routers' / 'files.py',
        REPO_ROOT / 'backend' / 'open_webui' / 'routers' / 'retrieval.py',
        REPO_ROOT / 'backend' / 'open_webui' / 'retrieval' / 'utils.py',
        REPO_ROOT / 'backend' / 'open_webui' / 'utils' / 'middleware.py',
    ]

    for path in changed_files:
        source = path.read_text()
        assert 'Skipping RAG processing for spreadsheet file' in source

    files_source = changed_files[0].read_text()
    assert "'raw_data_file': True" in files_source
    assert "'process_skipped': True" in files_source
