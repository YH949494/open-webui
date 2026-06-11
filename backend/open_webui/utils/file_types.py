import os


SPREADSHEET_FILE_EXTENSIONS = {'csv', 'xls', 'xlsx'}


def get_file_extension(filename: str | None) -> str:
    if not filename:
        return ''
    extension = os.path.splitext(filename)[1]
    return extension[1:].lower() if extension else ''


def is_spreadsheet_file(filename: str | None) -> bool:
    return get_file_extension(filename) in SPREADSHEET_FILE_EXTENSIONS
