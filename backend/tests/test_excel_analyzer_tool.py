"""Tests for the Excel analyzer tool and the uploaded-file path resolver.

Importing ``open_webui.utils.files`` pulls in the full application stack
(langchain, redis, authlib, ...), so — mirroring ``test_spreadsheet_rag_bypass``
— the resolver is verified via source-contract assertions, while the tool's
behaviour is exercised directly with a stubbed resolver injected through
``sys.modules`` (no heavy imports required).
"""

import json
import sys
import types
from pathlib import Path
from unittest.mock import AsyncMock, patch

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

    result = Tools._analyze(str(xlsx), 'data.xlsx')
    assert 'resolved_path' not in result


# --- tool behaviour with a stubbed resolver -------------------------------


def _stub_resolver(return_value):
    """Inject a fake ``open_webui.utils.files`` exposing the resolver.

    The tool imports ``resolve_uploaded_file_path`` lazily, so a stub module in
    ``sys.modules`` is picked up without importing the real (heavy) module.
    """
    stub = types.ModuleType('open_webui.utils.files')
    stub.resolve_uploaded_file_path = AsyncMock(return_value=return_value)
    return patch.dict(sys.modules, {'open_webui.utils.files': stub})


@pytest.mark.asyncio
async def test_tool_inspects_attached_spreadsheet(tmp_path):
    xlsx = tmp_path / 'uim.xlsx'
    _make_xlsx(xlsx)

    tool = Tools()
    user = types.SimpleNamespace(id='u1', role='user')

    __files__ = [
        {
            'type': 'file',
            'id': 'file-1',
            'name': 'User Intelligence Master.xlsx',
            'file': {'filename': 'User Intelligence Master.xlsx', 'meta': {'name': 'x'}},
        }
    ]

    with patch.object(tool, '_resolve_user', new=AsyncMock(return_value=user)), _stub_resolver(str(xlsx)):
        out = json.loads(await tool.inspect_uploaded_spreadsheet(__files__=__files__, __user__={'id': 'u1'}))

    sheets = {s['name']: s['row_count'] for s in out['files'][0]['sheets']}
    assert sheets == {'People': 3, 'Orders': 6}


@pytest.mark.asyncio
async def test_tool_returns_diagnostic_when_path_missing():
    tool = Tools()
    user = types.SimpleNamespace(id='u1', role='user')

    __files__ = [{'type': 'file', 'id': 'file-1', 'name': 'broken.xlsx', 'file': {'meta': {'size': 10}}}]

    with patch.object(tool, '_resolve_user', new=AsyncMock(return_value=user)), _stub_resolver(None):
        out = json.loads(await tool.inspect_uploaded_spreadsheet(__files__=__files__, __user__={'id': 'u1'}))

    diag = out['files'][0]
    assert diag['error'] == 'file not found; the file path was not accessible to the tool'
    assert diag['file_id'] == 'file-1'
    assert diag['filename'] == 'broken.xlsx'
    assert 'available_meta_keys' in diag


@pytest.mark.asyncio
async def test_tool_reports_when_no_files_attached():
    tool = Tools()
    user = types.SimpleNamespace(id='u1', role='user')

    with patch.object(tool, '_resolve_user', new=AsyncMock(return_value=user)):
        out = json.loads(await tool.inspect_uploaded_spreadsheet(__files__=[], __user__={'id': 'u1'}))

    assert 'No uploaded file' in out['error']


@pytest.mark.asyncio
async def test_tool_requires_user_context():
    tool = Tools()
    with patch.object(tool, '_resolve_user', new=AsyncMock(return_value=None)):
        out = json.loads(await tool.inspect_uploaded_spreadsheet(__files__=[{'id': 'x'}], __user__=None))
    assert 'User context not available' in out['error']


@pytest.mark.asyncio
async def test_tool_passes_attachment_path_to_resolver(tmp_path):
    """attachment_path from item['file']['path'] is forwarded to the resolver."""
    xlsx = tmp_path / 'report.xlsx'
    _make_xlsx(xlsx)

    tool = Tools()
    user = types.SimpleNamespace(id='u1', role='user')

    storage_path = '/uploads/uuid_report.xlsx'
    __files__ = [
        {
            'type': 'file',
            'id': 'file-2',
            'name': 'report.xlsx',
            'file': {
                'filename': 'report.xlsx',
                'path': storage_path,
                'meta': {'name': 'report.xlsx'},
            },
        }
    ]

    stub = types.ModuleType('open_webui.utils.files')
    resolver_mock = AsyncMock(return_value=str(xlsx))
    stub.resolve_uploaded_file_path = resolver_mock

    with (
        patch.object(tool, '_resolve_user', new=AsyncMock(return_value=user)),
        patch.dict(sys.modules, {'open_webui.utils.files': stub}),
    ):
        await tool.inspect_uploaded_spreadsheet(__files__=__files__, __user__={'id': 'u1'})

    resolver_mock.assert_awaited_once()
    _, kwargs = resolver_mock.call_args
    assert kwargs.get('attachment_path') == storage_path


@pytest.mark.asyncio
async def test_tool_works_without_attachment_path():
    """When item['file'] has no path, attachment_path is None (DB fallback path)."""
    tool = Tools()
    user = types.SimpleNamespace(id='u1', role='user')

    __files__ = [
        {
            'type': 'file',
            'id': 'file-3',
            'name': 'data.xlsx',
            'file': {'filename': 'data.xlsx', 'meta': {}},
        }
    ]

    stub = types.ModuleType('open_webui.utils.files')
    resolver_mock = AsyncMock(return_value=None)
    stub.resolve_uploaded_file_path = resolver_mock

    with (
        patch.object(tool, '_resolve_user', new=AsyncMock(return_value=user)),
        patch.dict(sys.modules, {'open_webui.utils.files': stub}),
    ):
        await tool.inspect_uploaded_spreadsheet(__files__=__files__, __user__={'id': 'u1'})

    resolver_mock.assert_awaited_once()
    _, kwargs = resolver_mock.call_args
    assert kwargs.get('attachment_path') is None


# --- resolver source contract (avoids heavy import) -----------------------


def test_resolver_source_contract():
    src = (REPO_ROOT / 'backend' / 'open_webui' / 'utils' / 'files.py').read_text()
    assert 'async def resolve_uploaded_file_path' in src

    body = src[src.index('async def resolve_uploaded_file_path') : src.index('async def get_image_base64_from_file_id')]

    # Accepts an optional attachment_path to use before hitting the DB.
    assert 'attachment_path' in body
    # Resolves the storage URI to a local path via the storage provider
    # (keeps local + cloud/Fly volume compatible).
    assert 'Storage.get_file' in body
    # Enforces access control before reading another user's file (DB fallback).
    assert 'has_access_to_file' in body
    # Logs the resolved path (server-side debug visibility).
    assert 'resolved file id=' in body
    # Returns None (diagnostic-friendly) when the file has no stored path.
    assert 'has no stored path' in body
