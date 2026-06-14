"""Tests for the Excel analyzer tool and the uploaded-file path resolver.

Importing ``open_webui.utils.files`` / ``open_webui.storage.provider`` pulls in
the full application stack (boto3, azure, langchain, redis, ...), so — mirroring
``test_spreadsheet_rag_bypass`` — the resolver is verified via source-contract
assertions, while the tool's behaviour is exercised directly with the storage
provider stubbed through ``sys.modules`` (no heavy imports required).
"""

import json
import sys
import types
from pathlib import Path
from unittest.mock import AsyncMock, MagicMock, patch

import pytest

pytest.importorskip('openpyxl')

from open_webui.tools.contrib.excel_analyzer import Tools

REPO_ROOT = Path(__file__).resolve().parents[2]


def _make_xlsx(path: Path) -> None:
    import openpyxl

    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = 'People'
    ws1.append(['name', 'age'])
    ws1.append(['alice', 30])
    ws1.append(['bob', 25])

    ws2 = wb.create_sheet('Orders')
    ws2.append(['id', 'total'])
    for i in range(5):
        ws2.append([i, i * 10])

    wb.save(path)


def _stub_storage(get_file_return):
    """Inject a fake ``open_webui.storage.provider`` exposing ``Storage.get_file``.

    The tool imports it lazily, so a stub in ``sys.modules`` is picked up without
    importing the real (boto3/azure/gcs) module. Returns (patcher, mock_get_file).
    """
    stub = types.ModuleType('open_webui.storage.provider')
    mock_get_file = MagicMock(return_value=get_file_return)
    stub.Storage = types.SimpleNamespace(get_file=mock_get_file)
    return patch.dict(sys.modules, {'open_webui.storage.provider': stub}), mock_get_file


# --- _analyze: real spreadsheet, real openpyxl ----------------------------


def test_analyze_returns_sheet_names_and_row_counts(tmp_path):
    xlsx = tmp_path / 'User Intelligence Master.xlsx'
    _make_xlsx(xlsx)

    result = Tools._analyze(str(xlsx), 'User Intelligence Master.xlsx')

    assert result['filename'] == 'User Intelligence Master.xlsx'
    sheets = {s['name']: s['row_count'] for s in result['sheets']}
    assert sheets == {'People': 3, 'Orders': 6}


def test_analyze_does_not_leak_path_by_default(tmp_path):
    xlsx = tmp_path / 'data.xlsx'
    _make_xlsx(xlsx)
    assert 'resolved_path' not in Tools._analyze(str(xlsx), 'data.xlsx')


# --- spreadsheet detection (extension + MIME) -----------------------------


def test_is_spreadsheet_by_extension():
    assert Tools._is_spreadsheet({'name': 'a.xlsx'})
    assert Tools._is_spreadsheet({'name': 'a.csv'})
    assert not Tools._is_spreadsheet({'name': 'a.pdf'})


def test_is_spreadsheet_by_mime_type():
    # No spreadsheet extension, but the nested record carries the xlsx MIME type.
    item = {
        'name': 'export',
        'file': {'meta': {'content_type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}},
    }
    assert Tools._is_spreadsheet(item)


# --- (1)(2)(5) tool receives __files__ and resolves via Storage.get_file ---


@pytest.mark.asyncio
async def test_tool_resolves_local_path_via_storage_get_file(tmp_path):
    xlsx = tmp_path / 'uim.xlsx'
    _make_xlsx(xlsx)

    tool = Tools()
    storage_uri = '/data/uploads/abc_User Intelligence Master.xlsx'  # value stored in file.path
    __files__ = [
        {
            'type': 'file',
            'id': 'file-1',
            'name': 'User Intelligence Master.xlsx',
            'file': {'filename': 'User Intelligence Master.xlsx', 'path': storage_uri, 'meta': {'name': 'x'}},
        }
    ]

    patcher, mock_get_file = _stub_storage(str(xlsx))
    with patcher:
        out = json.loads(await tool.inspect_uploaded_spreadsheet(__files__=__files__, __user__={'id': 'u1'}))

    # Resolution went through Storage.get_file with the nested file.path (never open() directly).
    mock_get_file.assert_called_once_with(storage_uri)
    sheets = {s['name']: s['row_count'] for s in out['sheets']}
    assert sheets == {'People': 3, 'Orders': 6}


# --- (8) missing / unresolvable files give a clear diagnostic --------------


@pytest.mark.asyncio
async def test_missing_files_gives_clear_diagnostic():
    tool = Tools()
    out = json.loads(await tool.inspect_uploaded_spreadsheet(__files__=None, __user__={'id': 'u1'}))
    assert out['error'] == 'file not found; the file path was not accessible to the tool'
    assert out['attached_file_count'] == 0
    assert 'no files were passed to the tool' in out['reason']


@pytest.mark.asyncio
async def test_unresolvable_path_gives_full_diagnostic():
    tool = Tools()
    __files__ = [
        {'type': 'file', 'id': 'file-1', 'name': 'broken.xlsx', 'file': {'path': 's3://bucket/x', 'meta': {'size': 10}}}
    ]

    # Storage returns a non-existent path; id-fallback also yields nothing.
    patcher, _ = _stub_storage('/nonexistent/broken.xlsx')
    with patcher, patch.object(tool, '_resolve_by_id', new=AsyncMock(return_value=None)):
        out = json.loads(await tool.inspect_uploaded_spreadsheet(__files__=__files__, __user__={'id': 'u1'}))

    assert out['error'] == 'file not found; the file path was not accessible to the tool'
    assert out['attached_file_count'] == 1
    assert out['filenames'] == ['broken.xlsx']
    assert out['file_ids'] == ['file-1']
    assert 'top_level_keys' in out and 'nested_file_keys' in out
    assert out['path_exists_after_storage_get_file'] is False
    # No absolute paths leaked into the output.
    assert 'resolved_path' not in out


# --- spreadsheet metadata-only upload still works -------------------------


@pytest.mark.asyncio
async def test_metadata_only_spreadsheet_upload_still_works(tmp_path):
    """A RAG-skipped spreadsheet has no data['content']; the tool must still
    read it from disk via the stored path."""
    xlsx = tmp_path / 'meta_only.xlsx'
    _make_xlsx(xlsx)

    tool = Tools()
    __files__ = [
        {
            'type': 'file',
            'id': 'file-1',
            'name': 'meta_only.xlsx',
            'file': {
                'filename': 'meta_only.xlsx',
                'path': '/data/uploads/abc_meta_only.xlsx',
                'data': {'status': 'completed', 'raw_data_file': True, 'process_skipped': True},  # note: no 'content'
                'meta': {'raw_data_file': True, 'process_skipped': True, 'content_type': 'application/vnd.ms-excel'},
            },
        }
    ]

    patcher, _ = _stub_storage(str(xlsx))
    with patcher:
        out = json.loads(await tool.inspect_uploaded_spreadsheet(__files__=__files__, __user__={'id': 'u1'}))

    sheets = {s['name']: s['row_count'] for s in out['sheets']}
    assert sheets == {'People': 3, 'Orders': 6}


@pytest.mark.asyncio
async def test_id_fallback_used_when_nested_path_missing(tmp_path):
    xlsx = tmp_path / 'fallback.xlsx'
    _make_xlsx(xlsx)

    tool = Tools()
    __files__ = [{'type': 'file', 'id': 'file-1', 'name': 'fallback.xlsx', 'file': {'meta': {}}}]  # no path

    with patch.object(tool, '_resolve_by_id', new=AsyncMock(return_value=str(xlsx))):
        out = json.loads(await tool.inspect_uploaded_spreadsheet(__files__=__files__, __user__={'id': 'u1'}))

    sheets = {s['name']: s['row_count'] for s in out['sheets']}
    assert sheets == {'People': 3, 'Orders': 6}


# --- resolver source contract (avoids heavy import) -----------------------


def test_resolver_source_contract():
    src = (REPO_ROOT / 'backend' / 'open_webui' / 'utils' / 'files.py').read_text()
    assert 'async def resolve_uploaded_file_path' in src

    body = src[src.index('async def resolve_uploaded_file_path') : src.index('async def get_image_base64_from_file_id')]
    assert 'Storage.get_file' in body  # resolves storage URI -> local path
    assert 'has_access_to_file' in body  # access control
    assert 'resolved file id=' in body  # server-side log
    assert 'has no stored path' in body  # diagnostic when path missing


def test_tool_signature_declares_reserved_params():
    import inspect

    params = inspect.signature(Tools.inspect_uploaded_spreadsheet).parameters
    # (1) __files__ must be declared so Open WebUI injects the attachments.
    assert '__files__' in params
    assert '__user__' in params
