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
    sheets = {s['sheet_name']: s['rows'] for s in result['sheets']}
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

    sheets = {s['sheet_name']: s['rows'] for s in out['files'][0]['sheets']}
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


# --- row-level preview: merged cells, repeated headers, blanks, totals ----


def _make_finance_xlsx(path: Path) -> None:
    """Build a small "monthly expenses" sheet mirroring a real-world export:

    - the ``Month`` cell is merged across each month's detail rows (a common
      pattern from spreadsheets exported/copy-pasted out of reporting tools)
    - a blank spacer row
    - the header row repeated mid-sheet (e.g. from concatenating two exports)
    - a trailing "Grand Total" summary row
    """
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Expenses'

    headers = ['Month', 'Description', 'Total (RM)']
    ws.append(headers)

    ws.append(['Jan', 'Rent', 1000])
    ws.append([None, 'Utilities', 200])
    ws.append([None, 'Supplies', 50])
    ws.append(['Feb', 'Rent', 1000])
    ws.append([None, 'Utilities', 220])
    ws.merge_cells(start_row=2, end_row=4, start_column=1, end_column=1)
    ws.merge_cells(start_row=5, end_row=6, start_column=1, end_column=1)

    ws.append([None, None, None])  # blank row
    ws.append(headers)  # repeated header row
    ws.append(['Mar', 'Rent', 1000])
    ws.append([None, 'Utilities', 210])
    ws.append(['Grand Total', None, 2680])

    wb.save(path)


def test_analyze_xlsx_preview_preserves_merged_month_relationships(tmp_path):
    """The root-cause regression: previews must not come back empty, and the
    Month/Description/Total (RM) relationship must survive merged Month cells."""
    xlsx = tmp_path / 'expenses.xlsx'
    _make_finance_xlsx(xlsx)

    result = Tools._analyze(str(xlsx), 'expenses.xlsx')
    sheet = result['sheets'][0]

    assert sheet['rows'] == 11  # header + 10 sheet rows, including blank/repeat/total
    assert sheet['column_names'] == ['Month', 'Description', 'Total (RM)']

    preview = sheet['preview']
    assert preview != []
    assert not sheet['preview_truncated']

    # Blank row and the repeated header row are dropped entirely.
    assert sheet['blank_rows_skipped'] == 1
    assert sheet['repeated_header_rows_skipped'] == 1
    assert sheet['grand_total_rows'] == 1
    assert sheet['data_row_count'] == len(preview) == 8

    # Merged "Month" cells are forward-filled onto every detail row they cover.
    jan_rows = [r for r in preview if r['Month'] == 'Jan']
    assert {r['Description'] for r in jan_rows} == {'Rent', 'Utilities', 'Supplies'}
    assert [r['Total (RM)'] for r in jan_rows] == [1000, 200, 50]

    feb_rows = [r for r in preview if r['Month'] == 'Feb']
    assert {r['Description'] for r in feb_rows} == {'Rent', 'Utilities'}

    # The Grand Total row stays in the preview (flagged), not silently dropped.
    total_row = next(r for r in preview if r.get('is_summary_row'))
    assert total_row['Month'] == 'Grand Total'
    assert total_row['Total (RM)'] == 2680

    # Value counts reflect the real per-column distribution, excluding the
    # Grand Total row so it doesn't skew the counts.
    month_counts = {v['value']: v['count'] for v in sheet['value_counts']['Month']}
    assert month_counts == {'Jan': 3, 'Feb': 2, 'Mar': 1}
    description_counts = {v['value']: v['count'] for v in sheet['value_counts']['Description']}
    assert description_counts['Rent'] == 3
    assert description_counts['Utilities'] == 3


def test_analyze_xlsx_lightweight_mode_omits_preview(tmp_path):
    """lightweight_mode is an explicit opt-out, not the default failure mode."""
    xlsx = tmp_path / 'expenses.xlsx'
    _make_finance_xlsx(xlsx)

    result = Tools._analyze(str(xlsx), 'expenses.xlsx', lightweight_mode=True)
    sheet = result['sheets'][0]

    assert 'preview' not in sheet
    assert 'value_counts' not in sheet
    assert sheet['rows'] == 11
    assert sheet['column_names'] == ['Month', 'Description', 'Total (RM)']


def test_analyze_xlsx_large_sheet_gets_bounded_preview(tmp_path):
    """Sheets above the full-preview threshold get a bounded sample, not a full dump."""
    import openpyxl

    xlsx = tmp_path / 'large.xlsx'
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['id', 'value'])
    for i in range(30):
        ws.append([i, i * 2])
    wb.save(xlsx)

    result = Tools._analyze(
        str(xlsx),
        'large.xlsx',
        full_preview_row_threshold=10,
        bounded_preview_rows=5,
    )
    sheet = result['sheets'][0]

    assert sheet['data_row_count'] == 30
    assert sheet['preview_truncated'] is True
    assert len(sheet['preview']) == 5
    assert sheet['preview_omitted_rows'] == 25
    assert sheet['preview'][0] == {'id': 0, 'value': 0}


def test_analyze_xlsx_large_sheet_does_not_materialize_every_cell(tmp_path, monkeypatch):
    """A large sheet's row-level scan must stay bounded by bounded_preview_rows,
    not read every row before the preview is truncated (memory/latency guard)."""
    import openpyxl

    xlsx = tmp_path / 'huge.xlsx'
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['id', 'value'])
    for i in range(5000):
        ws.append([i, i * 2])
    wb.save(xlsx)

    seen_rows = []
    wb2 = openpyxl.load_workbook(str(xlsx), read_only=True, data_only=True)
    real_iter_rows = wb2.worksheets[0].__class__.iter_rows

    def counting_iter_rows(self, *args, **kwargs):
        for row in real_iter_rows(self, *args, **kwargs):
            seen_rows.append(row)
            yield row

    monkeypatch.setattr('openpyxl.worksheet._read_only.ReadOnlyWorksheet.iter_rows', counting_iter_rows)
    wb2.close()

    result = Tools._analyze(
        str(xlsx),
        'huge.xlsx',
        full_preview_row_threshold=200,
        bounded_preview_rows=50,
    )
    sheet = result['sheets'][0]

    assert sheet['data_row_count'] == 5000
    assert sheet['preview_truncated'] is True
    assert len(sheet['preview']) == 50
    assert sheet['preview_omitted_rows'] == 4950
    # The header pass (1 row) + the bounded data pass (<=51 rows) — not 5001 rows.
    assert len(seen_rows) <= 52


def test_analyze_xlsx_date_cells_are_json_serializable(tmp_path):
    """A date/datetime column must not crash JSON serialization of the analysis."""
    import datetime as dt

    import openpyxl

    xlsx = tmp_path / 'dated.xlsx'
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['Date', 'Amount'])
    ws.append([dt.date(2026, 1, 15), 100])
    ws.append([dt.datetime(2026, 1, 16, 9, 30), 200])
    wb.save(xlsx)

    result = Tools._analyze(str(xlsx), 'dated.xlsx')
    # Must round-trip through json.dumps without raising, exactly like _guard_output does.
    raw = Tools._guard_output({'files': [result]})
    reparsed = json.loads(raw)

    sheet = reparsed['files'][0]['sheets'][0]
    assert sheet['preview'][0]['Date'] == '2026-01-15T00:00:00'
    assert sheet['preview'][1]['Date'] == '2026-01-16T09:30:00'
    assert sheet['preview'][0]['Amount'] == 100


def test_analyze_csv_large_file_reports_true_total_not_sample_size(tmp_path):
    """A CSV above the threshold must report its true row count/omission, not the sample size."""
    csv_path = tmp_path / 'big.csv'
    lines = ['Month,Description,Total (RM)']
    lines += [f'Jan,Item {i},{i}' for i in range(500)]
    csv_path.write_text('\n'.join(lines) + '\n')

    result = Tools._analyze(
        str(csv_path),
        'big.csv',
        full_preview_row_threshold=100,
        bounded_preview_rows=20,
    )
    sheet = result['sheets'][0]

    assert sheet['data_row_count'] == 500
    assert sheet['preview_truncated'] is True
    assert len(sheet['preview']) == 20
    assert sheet['preview_omitted_rows'] == 480
    assert sheet.get('value_counts_from_sample') is True


def test_analyze_csv_preview_preserves_row_relationships(tmp_path):
    """Same root-cause coverage for CSV: blank line, repeated header, Grand Total."""
    csv_path = tmp_path / 'expenses.csv'
    csv_path.write_text(
        'Month,Description,Total (RM)\n'
        'Jan,Rent,1000\n'
        'Jan,Utilities,200\n'
        'Feb,Rent,1000\n'
        '\n'
        'Month,Description,Total (RM)\n'
        'Mar,Rent,1000\n'
        'Grand Total,,2200\n'
    )

    result = Tools._analyze(str(csv_path), 'expenses.csv')
    sheet = result['sheets'][0]

    preview = sheet['preview']
    assert preview != []
    assert sheet['repeated_header_rows_skipped'] == 1
    assert sheet['grand_total_rows'] == 1

    jan_rows = [r for r in preview if r['Month'] == 'Jan']
    assert {r['Description'] for r in jan_rows} == {'Rent', 'Utilities'}

    total_row = next(r for r in preview if r.get('is_summary_row'))
    assert total_row['Total (RM)'] == 2200


@pytest.mark.asyncio
async def test_tool_end_to_end_preview_reaches_output(tmp_path):
    """End-to-end through inspect_uploaded_spreadsheet: the preview a caller
    (and, via build_spreadsheet_analysis_sources, the LLM) actually receives
    is non-empty and keeps Month/Description/Total (RM) together."""
    xlsx = tmp_path / 'expenses.xlsx'
    _make_finance_xlsx(xlsx)

    tool = Tools()
    user = types.SimpleNamespace(id='u1', role='user')
    __files__ = [
        {
            'type': 'file',
            'id': 'file-1',
            'name': 'expenses.xlsx',
            'file': {'filename': 'expenses.xlsx', 'meta': {'name': 'x'}},
        }
    ]

    with patch.object(tool, '_resolve_user', new=AsyncMock(return_value=user)), _stub_resolver(str(xlsx)):
        out = json.loads(await tool.inspect_uploaded_spreadsheet(__files__=__files__, __user__={'id': 'u1'}))

    sheet = out['files'][0]['sheets'][0]
    assert sheet['preview'] != []
    assert any(row.get('Month') == 'Jan' and row.get('Description') == 'Rent' for row in sheet['preview'])


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
