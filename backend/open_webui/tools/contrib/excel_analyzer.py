"""
title: Excel Analyzer
author: open-webui
version: 1.1.0
required_open_webui_version: 0.5.0
requirements: openpyxl
description: >
  Inspect spreadsheet files (.xlsx/.xlsm/.xls/.csv) that a user has attached to
  the chat. The tool resolves the *actual server-side path* of the uploaded
  file and reads it directly with openpyxl/pandas instead of relying on RAG
  extracted text. This is the companion tool to the "skip RAG for spreadsheets"
  upload path: spreadsheets are stored as raw data files (never chunked into the
  vector DB and never injected into the prompt), so a tool must open the file on
  disk to analyze it. In addition to sheet/column metadata, it returns a
  row-level preview (every row for small sheets, a bounded sample for large
  ones) and per-column top values, so row relationships (e.g. a Month column
  next to a Description and a Total) survive into the model's context.

INSTALLATION
------------
This file is the source for a *custom* Open WebUI tool. Built-in tools are
registered in ``open_webui/tools/builtin.py``; this module is intentionally NOT
auto-loaded. To use it:

  Workspace -> Tools -> (+) -> paste the contents of the ``Tools`` class below
  (or this whole file) -> Save.

Then attach a spreadsheet to a chat and ask, e.g.:

  "Use the Excel analyzer tool to inspect the uploaded file.
   Return only sheet names and row counts."

HOW FILE ACCESS WORKS
---------------------
Open WebUI passes attached files to a tool through the reserved ``__files__``
parameter (``metadata['files']``). Each entry looks like::

    {"type": "file", "id": "<file_id>", "name": "Foo.xlsx",
     "file": { ...full File record incl. "path"... }}

The stored ``file.path`` is a *storage URI* (``s3://``, ``gs://``, an Azure
URL, or — for local/Fly persistent-volume storage — an absolute path). It is
NOT necessarily openable with ``open()``. ``resolve_uploaded_file_path`` turns a
``file_id`` into a readable local path: it returns the absolute path as-is for
local storage, and downloads cloud objects into ``UPLOAD_DIR`` for cloud
providers. The same code path works on a Fly.io persistent volume because
``UPLOAD_DIR`` is the mounted volume there.
"""

from __future__ import annotations

import datetime as dt
import json
import logging
import os
import zipfile
from collections import Counter
from decimal import Decimal
from xml.etree import ElementTree as ET

from pydantic import BaseModel, Field

log = logging.getLogger(__name__)

# Extensions this tool knows how to open directly.
_SPREADSHEET_EXTENSIONS = {'.xlsx', '.xlsm', '.xltx', '.xltm', '.xls', '.csv'}
_MAX_COL_NAMES = 20
_MAX_INSIGHTS = 10
_MAX_SAMPLE_ROWS = 5000
_OUTPUT_SIZE_LIMIT = 12000
_MIN_PREVIEW_ROWS = 5  # floor used when the size guard has to shrink previews

_XML_NS = {
    'm': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}


class Tools:
    class Valves(BaseModel):
        max_files: int = Field(
            default=10,
            description='Maximum number of attached spreadsheet files to inspect in one call.',
        )
        expose_paths_in_output: bool = Field(
            default=False,
            description=(
                'If true, include the resolved absolute server path in the tool output. '
                'Leave false in production so sensitive paths stay in the server logs only.'
            ),
        )
        lightweight_mode: bool = Field(
            default=False,
            description=(
                'When true, return only compact metadata: sheet names, row/column counts, and up to '
                '20 column names — no row-level preview or value counts. When false (default), also '
                'include row-level records (full for small sheets, bounded for large ones) and '
                'per-column top values so row relationships (e.g. Month/Description/Total) survive.'
            ),
        )
        full_preview_row_threshold: int = Field(
            default=200,
            description=(
                'Sheets with this many data rows or fewer get a full row-by-row preview. Larger sheets '
                'get a bounded preview instead (see bounded_preview_rows).'
            ),
        )
        bounded_preview_rows: int = Field(
            default=50,
            description='Number of data rows to include in the preview for sheets above full_preview_row_threshold.',
        )
        max_value_counts_per_column: int = Field(
            default=5,
            description='Maximum number of top values reported per column in value_counts.',
        )

    def __init__(self):
        self.valves = self.Valves()
        # We read the raw uploaded file ourselves; we do NOT want Open WebUI to
        # strip the files from the request or treat this as a citation handler.
        self.file_handler = False
        self.citation = False

    async def inspect_uploaded_spreadsheet(
        self,
        file_id: str = '',
        include_insights: bool = False,
        __files__: list | None = None,
        __user__: dict | None = None,
    ) -> str:
        """
        Inspect an uploaded spreadsheet (.xlsx/.xlsm/.xls/.csv) and return sheet
        names, row/column counts, up to 20 column headers, a row-level preview
        (full for small sheets, bounded for large ones), and per-column top
        values per sheet. Optionally include up to 10 structural insights.
        Output is always compact JSON capped at 12 000 characters.

        :param file_id: Optional id of a specific attached file. If omitted, the most recent spreadsheet is used.
        :param include_insights: If true, include up to 10 structural findings per sheet.
        :return: Compact JSON with sheet metadata and a row-level preview.
        """
        user = await self._resolve_user(__user__)
        if user is None:
            return json.dumps({'error': 'User context not available; cannot access uploaded files.'})

        candidates = self._collect_candidates(__files__, file_id)
        if not candidates:
            return json.dumps(
                {
                    'error': 'No uploaded file was available to the tool.',
                    'hint': 'Attach a spreadsheet to the chat before invoking this tool.',
                    'received_files': self._describe_files(__files__),
                }
            )

        from open_webui.utils.files import resolve_uploaded_file_path

        results = []
        for item in candidates[: max(1, self.valves.max_files)]:
            fid = item.get('id')
            fname = item.get('name') or item.get('filename') or (item.get('file') or {}).get('filename')
            # Pass the storage path from the attachment dict so the resolver can
            # use it directly without a DB round-trip.
            attachment_path = (item.get('file') or {}).get('path') or None

            local_path = await resolve_uploaded_file_path(fid, user=user, attachment_path=attachment_path)
            if not local_path:
                results.append(self._missing_path_diagnostic(item, fid, fname))
                continue

            log.info(f'inspect_uploaded_spreadsheet: analyzing file_id={fid} at {local_path}')
            analysis = self._analyze(
                local_path,
                fname,
                include_insights=include_insights,
                lightweight_mode=self.valves.lightweight_mode,
                full_preview_row_threshold=self.valves.full_preview_row_threshold,
                bounded_preview_rows=self.valves.bounded_preview_rows,
                max_value_counts_per_column=self.valves.max_value_counts_per_column,
            )
            if self.valves.expose_paths_in_output:
                analysis['resolved_path'] = local_path
            results.append(analysis)

        return self._guard_output({'files': results})

    # --- helpers -----------------------------------------------------------

    async def _resolve_user(self, __user__: dict | None):
        """Turn the ``__user__`` dict Open WebUI injects into a full user object."""
        if not __user__ or not __user__.get('id'):
            return None
        try:
            from open_webui.models.users import Users

            return await Users.get_user_by_id(__user__['id'])
        except Exception as e:
            log.warning(f'inspect_uploaded_spreadsheet: could not resolve user: {e}')
            return None

    def _collect_candidates(self, __files__: list | None, file_id: str) -> list:
        """Pick which attached file(s) to inspect."""
        files = [f for f in (__files__ or []) if isinstance(f, dict)]

        if file_id:
            for f in files:
                if f.get('id') == file_id:
                    return [f]
            return [{'id': file_id}]

        spreadsheets = [f for f in files if self._is_spreadsheet(f)]
        return spreadsheets or files

    @staticmethod
    def _is_spreadsheet(item: dict) -> bool:
        name = item.get('name') or item.get('filename') or (item.get('file') or {}).get('filename') or ''
        return os.path.splitext(name.lower())[1] in _SPREADSHEET_EXTENSIONS

    @staticmethod
    def _describe_files(__files__: list | None) -> list:
        described = []
        for f in __files__ or []:
            if isinstance(f, dict):
                described.append(
                    {
                        'id': f.get('id'),
                        'name': f.get('name') or f.get('filename'),
                        'type': f.get('type'),
                        'keys': sorted(f.keys()),
                    }
                )
        return described

    @staticmethod
    def _missing_path_diagnostic(item: dict, fid, fname) -> dict:
        file_meta = (item.get('file') or {}).get('meta') or {}
        return {
            'error': 'file not found; the file path was not accessible to the tool',
            'file_id': fid,
            'filename': fname,
            'available_item_keys': sorted(item.keys()),
            'available_meta_keys': sorted(file_meta.keys()),
        }

    @staticmethod
    def _shrink_previews_pass(payload: dict) -> bool:
        """Halve (down to _MIN_PREVIEW_ROWS) each sheet's preview. Returns True if anything shrank."""
        shrunk = False
        for f in payload.get('files', []):
            for s in f.get('sheets', []):
                preview = s.get('preview')
                if preview and len(preview) > _MIN_PREVIEW_ROWS:
                    kept = max(_MIN_PREVIEW_ROWS, len(preview) // 2)
                    if kept < len(preview):
                        s['preview'] = preview[:kept]
                        s['preview_truncated'] = True
                        s['preview_omitted_rows'] = s.get('preview_omitted_rows', 0) + (len(preview) - kept)
                        shrunk = True
        return shrunk

    @staticmethod
    def _drop_sheet_field_pass(payload: dict, field: str) -> bool:
        for f in payload.get('files', []):
            for s in f.get('sheets', []):
                s.pop(field, None)
        return True

    @staticmethod
    def _cap_sheets_pass(payload: dict) -> bool:
        for f in payload.get('files', []):
            if len(f.get('sheets', [])) > 10:
                f['sheets'] = f['sheets'][:10]
                f['sheets_truncated'] = True
        return True

    @staticmethod
    def _guard_output(payload: dict) -> str:
        """Ensure output stays under _OUTPUT_SIZE_LIMIT characters with valid JSON.

        Degrades the payload in passes, checking the size after each: shrink
        previews first (the row-level preview is the primary payload, so it's
        degraded gracefully before anything is dropped outright), then drop
        progressively less essential fields, then cap the sheet count.
        """
        raw = json.dumps(payload, ensure_ascii=False, separators=(',', ':'))
        if len(raw) <= _OUTPUT_SIZE_LIMIT:
            return raw

        passes = [
            Tools._shrink_previews_pass,
            lambda p: Tools._drop_sheet_field_pass(p, 'insights'),
            lambda p: Tools._drop_sheet_field_pass(p, 'value_counts'),
            lambda p: Tools._drop_sheet_field_pass(p, 'preview'),
            lambda p: Tools._drop_sheet_field_pass(p, 'column_names'),
            Tools._cap_sheets_pass,
        ]
        for run_pass in passes:
            if run_pass(payload):
                payload['truncated'] = True
                raw = json.dumps(payload, ensure_ascii=False, separators=(',', ':'))
                if len(raw) <= _OUTPUT_SIZE_LIMIT:
                    return raw
        return raw

    @staticmethod
    def _analyze(
        local_path: str,
        fname: str | None,
        include_insights: bool = False,
        lightweight_mode: bool = False,
        full_preview_row_threshold: int = 200,
        bounded_preview_rows: int = 50,
        max_value_counts_per_column: int = 5,
    ) -> dict:
        """Dispatch to the correct reader based on file extension."""
        ext = os.path.splitext((fname or local_path).lower())[1]
        display_name = fname or os.path.basename(local_path)
        kwargs = dict(
            include_insights=include_insights,
            lightweight_mode=lightweight_mode,
            full_preview_row_threshold=full_preview_row_threshold,
            bounded_preview_rows=bounded_preview_rows,
            max_value_counts_per_column=max_value_counts_per_column,
        )

        try:
            if ext == '.csv':
                return Tools._analyze_csv(local_path, display_name, **kwargs)
            if ext == '.xls':
                return Tools._analyze_xls(local_path, display_name, **kwargs)
            # Default: xlsx / xlsm / xltx / xltm via openpyxl streaming read.
            return Tools._analyze_xlsx(local_path, display_name, **kwargs)
        except Exception as e:
            log.exception(f'inspect_uploaded_spreadsheet: failed to read {local_path}: {e}')
            return {'filename': display_name, 'error': f'Failed to read spreadsheet: {e}'}

    @staticmethod
    def _merged_ranges_xlsx(local_path: str, sheet_index: int) -> list[tuple[int, int, int, int]]:
        """Return (min_row, max_row, min_col, max_col) for each merged range on a sheet.

        Read via the raw sheet XML (not ``openpyxl``'s ``merged_cells``) because
        ``read_only`` worksheets don't expose that attribute, and switching to a
        non-read-only load would force a full in-memory parse of every cell.
        """
        try:
            from openpyxl.utils.cell import range_boundaries

            with zipfile.ZipFile(local_path) as z:
                wb_xml = ET.fromstring(z.read('xl/workbook.xml'))
                sheet_els = wb_xml.findall('.//m:sheets/m:sheet', _XML_NS)
                if sheet_index >= len(sheet_els):
                    return []
                rid = sheet_els[sheet_index].get(f'{{{_XML_NS["r"]}}}id')
                rels_xml = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
                target = next((rel.get('Target') for rel in rels_xml if rel.get('Id') == rid), None)
                if not target:
                    return []
                path = target.lstrip('/') if target.startswith('/') else 'xl/' + target
                sheet_xml = ET.fromstring(z.read(path))
                ranges = []
                for mc in sheet_xml.findall('.//m:mergeCells/m:mergeCell', _XML_NS):
                    ref = mc.get('ref')
                    if not ref:
                        continue
                    min_col, min_row, max_col, max_row = range_boundaries(ref)
                    ranges.append((min_row, max_row, min_col, max_col))
                return ranges
        except Exception as e:
            log.warning(f'inspect_uploaded_spreadsheet: could not read merged cell ranges: {e}')
            return []

    @staticmethod
    def _analyze_xlsx(
        local_path: str,
        display_name: str,
        include_insights: bool,
        lightweight_mode: bool,
        full_preview_row_threshold: int,
        bounded_preview_rows: int,
        max_value_counts_per_column: int,
    ) -> dict:
        """Read xlsx data via openpyxl's streaming reader.

        Header + counts always come from a single streaming pass. When
        ``lightweight_mode`` is off (the default), a second pass builds a
        row-level preview and per-column value counts, forward-filling merged
        cells (e.g. a "Month" column merged across several detail rows) so the
        row relationships survive even though openpyxl's read-only reader only
        reports a value on a merged range's anchor cell.
        """
        import openpyxl

        wb = openpyxl.load_workbook(local_path, read_only=True, data_only=True)
        try:
            sheets = []
            for sheet_index, ws in enumerate(wb.worksheets):
                # max_row/max_column from the workbook manifest — no full read.
                row_count = int(ws.max_row) if ws.max_row is not None else 0
                col_count = int(ws.max_column) if ws.max_column is not None else 0

                col_names: list[str] = []
                for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
                    col_names = [str(c) if c is not None else '' for c in row[:_MAX_COL_NAMES]]
                    break

                sheet_info: dict = {
                    'sheet_name': ws.title,
                    'rows': row_count,
                    'columns': col_count,
                    'column_names': col_names,
                }
                if include_insights:
                    sheet_info['insights'] = Tools._structural_insights(row_count, col_count, col_names)

                if not lightweight_mode and row_count > 1 and col_count > 0:
                    sheet_info.update(
                        Tools._build_row_preview_xlsx(
                            ws,
                            local_path,
                            sheet_index,
                            row_count,
                            col_count,
                            full_preview_row_threshold,
                            bounded_preview_rows,
                            max_value_counts_per_column,
                        )
                    )
                sheets.append(sheet_info)
            return {'filename': display_name, 'sheets': sheets}
        finally:
            wb.close()

    @staticmethod
    def _build_row_preview_xlsx(
        ws,
        local_path: str,
        sheet_index: int,
        row_count: int,
        col_count: int,
        full_preview_row_threshold: int,
        bounded_preview_rows: int,
        max_value_counts_per_column: int,
    ) -> dict:
        """Build the row-level preview/value_counts fields for one xlsx sheet.

        For sheets whose data rows exceed ``full_preview_row_threshold``, only
        rows up to a bounded scan window are ever read into ``matrix`` — a
        multi-million-row workbook must not materialize every cell just to
        return a ``bounded_preview_rows``-sized sample.
        """
        total_data_rows = max(0, row_count - 1)
        sample_only = total_data_rows > full_preview_row_threshold
        scan_max_row = min(row_count, 1 + bounded_preview_rows) if sample_only else row_count

        matrix: dict[tuple[int, int], object] = {}
        for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=scan_max_row, values_only=False), start=1):
            for c_idx, cell in enumerate(row, start=1):
                matrix[(r_idx, c_idx)] = getattr(cell, 'value', None)

        # Forward-fill merged ranges (e.g. a "Month" cell merged across several
        # detail rows) using the anchor cell's already-read value.
        for min_row, max_row, min_col, max_col in Tools._merged_ranges_xlsx(local_path, sheet_index):
            for r in range(min_row, min(max_row, scan_max_row) + 1):
                anchor = matrix.get((min_row, min_col))
                for c in range(min_col, max_col + 1):
                    if matrix.get((r, c)) is None:
                        matrix[(r, c)] = anchor

        full_col_names = [matrix.get((1, c)) for c in range(1, col_count + 1)]
        full_col_names = [str(c) if c is not None else '' for c in full_col_names]

        row_values = []
        for r in range(2, scan_max_row + 1):
            row_values.append([matrix.get((r, c)) for c in range(1, col_count + 1)])

        return Tools._summarize_rows(
            full_col_names,
            row_values,
            max_value_counts_per_column,
            sample_only=sample_only,
            known_total_data_rows=total_data_rows,
        )

    @staticmethod
    def _analyze_xls(
        local_path: str,
        display_name: str,
        include_insights: bool,
        lightweight_mode: bool,
        full_preview_row_threshold: int,
        bounded_preview_rows: int,
        max_value_counts_per_column: int,
    ) -> dict:
        """Read legacy .xls metadata (and, unless lightweight, row data) via xlrd."""
        try:
            import xlrd

            wb = xlrd.open_workbook(local_path)
            sheets = []
            for sh in wb.sheets():
                nrows = int(sh.nrows)
                ncols = int(sh.ncols)
                col_names = [str(sh.cell_value(0, c)) for c in range(min(ncols, _MAX_COL_NAMES))] if nrows > 0 else []
                sheet_info: dict = {
                    'sheet_name': sh.name,
                    'rows': nrows,
                    'columns': ncols,
                    'column_names': col_names,
                }
                if include_insights:
                    sheet_info['insights'] = Tools._structural_insights(nrows, ncols, col_names)
                if not lightweight_mode and nrows > 1 and ncols > 0:
                    # xlrd loads the whole workbook into memory regardless, but
                    # still bound how many rows we copy out into our own preview.
                    total_data_rows = nrows - 1
                    sample_only = total_data_rows > full_preview_row_threshold
                    scan_max_row = min(nrows, 1 + bounded_preview_rows) if sample_only else nrows
                    full_col_names = [str(sh.cell_value(0, c)) for c in range(ncols)]
                    row_values = [[sh.cell_value(r, c) for c in range(ncols)] for r in range(1, scan_max_row)]
                    sheet_info.update(
                        Tools._summarize_rows(
                            full_col_names,
                            row_values,
                            max_value_counts_per_column,
                            sample_only=sample_only,
                            known_total_data_rows=total_data_rows,
                        )
                    )
                sheets.append(sheet_info)
            return {'filename': display_name, 'sheets': sheets}
        except ImportError:
            # xlrd not installed — fall back to pandas header-only read.
            import pandas as pd

            header_frames = pd.read_excel(local_path, sheet_name=None, header=0, nrows=0)
            sheets = []
            for name, df in header_frames.items():
                col_names = [str(c) for c in df.columns[:_MAX_COL_NAMES]]
                sheet_info = {
                    'sheet_name': name,
                    'rows': None,
                    'columns': len(df.columns),
                    'column_names': col_names,
                    'note': 'row count unavailable without xlrd',
                }
                if include_insights:
                    sheet_info['insights'] = Tools._structural_insights(None, len(df.columns), col_names)
                sheets.append(sheet_info)
            return {'filename': display_name, 'sheets': sheets}

    @staticmethod
    def _analyze_csv(
        local_path: str,
        display_name: str,
        include_insights: bool,
        lightweight_mode: bool,
        full_preview_row_threshold: int,
        bounded_preview_rows: int,
        max_value_counts_per_column: int,
    ) -> dict:
        """Count CSV rows by line scan; read only the header (or, unless lightweight, the data too) with pandas."""
        import pandas as pd

        row_count = 0
        try:
            with open(local_path, 'rb') as fh:
                row_count = max(0, sum(1 for _ in fh) - 1)  # subtract header line
        except Exception:
            pass

        header_df = pd.read_csv(local_path, nrows=0)
        col_count = len(header_df.columns)
        col_names = [str(c) for c in header_df.columns[:_MAX_COL_NAMES]]

        sheet_info: dict = {
            'sheet_name': 'Sheet1',
            'rows': row_count,
            'columns': col_count,
            'column_names': col_names,
        }
        if include_insights:
            sheet_info['insights'] = Tools._structural_insights(row_count, col_count, col_names)

        if not lightweight_mode and row_count > 0 and col_count > 0:
            full_col_names = [str(c) for c in header_df.columns]
            # For large files, only physically read a bounded sample — this keeps
            # memory bounded for large files while still using the accurate
            # full-file row_count computed above for the small-file decision.
            read_all = row_count <= full_preview_row_threshold
            data_df = pd.read_csv(local_path, nrows=None if read_all else bounded_preview_rows)
            # Build per-column (not via .values, which unifies dtypes across the
            # whole frame and would stringify numeric columns next to text ones).
            # A repeated header row mixed into the data also forces pandas to
            # read an otherwise-numeric column as strings, so coerce back.
            columns = [
                [Tools._coerce_numeric(v) for v in data_df[col].where(pd.notna(data_df[col]), None).tolist()]
                for col in data_df.columns
            ]
            row_values = [list(row) for row in zip(*columns)] if columns else []
            sheet_info.update(
                Tools._summarize_rows(
                    full_col_names,
                    row_values,
                    max_value_counts_per_column,
                    sample_only=not read_all,
                    known_total_data_rows=row_count,
                )
            )
        return {'filename': display_name, 'sheets': [sheet_info]}

    @staticmethod
    def _coerce_numeric(v):
        """Turn a cleanly-numeric string back into an int/float.

        A single non-numeric row mixed into an otherwise numeric CSV column
        (e.g. a repeated header row) makes pandas read the whole column as
        strings; this restores the original numeric type for the cells that
        actually are numbers, without touching real text values.
        """
        if not isinstance(v, str):
            return v
        s = v.strip()
        if s == '':
            return None
        try:
            return int(s)
        except ValueError:
            pass
        try:
            return float(s)
        except ValueError:
            return v

    @staticmethod
    def _json_safe_value(v):
        """Coerce a raw cell value into something json.dumps can handle.

        openpyxl/xlrd return native ``datetime.date``/``datetime.datetime``/
        ``datetime.time`` objects for date-formatted cells, which json.dumps
        rejects outright; without this, a spreadsheet with a date column would
        make the whole analysis fail instead of returning a preview.
        """
        if isinstance(v, (dt.datetime, dt.date, dt.time)):
            return v.isoformat()
        if isinstance(v, Decimal):
            return float(v)
        return v

    @staticmethod
    def _is_blank_row(values: list) -> bool:
        return all(v is None or (isinstance(v, str) and v.strip() == '') for v in values)

    @staticmethod
    def _is_repeated_header(values: list, col_names: list[str]) -> bool:
        str_values = [str(v) if v is not None else '' for v in values]
        return str_values == col_names

    @staticmethod
    def _is_summary_row(values: list) -> bool:
        first = next((v for v in values if v is not None and (not isinstance(v, str) or v.strip() != '')), None)
        return isinstance(first, str) and 'total' in first.lower()

    @staticmethod
    def _classify_rows(col_names: list[str], row_values: list[list]) -> tuple[list[tuple[list, bool]], dict]:
        """Split raw rows into qualifying (row_values, is_summary) pairs plus skip counters.

        Blank rows and rows that just repeat the header are dropped entirely.
        Rows that look like a "Grand Total" summary line are kept (flagged
        ``is_summary_row``) so they stay visible in the preview, but excluded
        from ``value_counts`` so they don't skew per-column frequency stats.
        """
        col_count = len(col_names)
        qualifying: list[tuple[list, bool]] = []
        counters = {'blank_rows_skipped': 0, 'repeated_header_rows_skipped': 0, 'grand_total_rows': 0}

        for values in row_values:
            values = [Tools._json_safe_value(v) for v in values[:col_count]] + [None] * max(0, col_count - len(values))
            if Tools._is_blank_row(values):
                counters['blank_rows_skipped'] += 1
                continue
            if Tools._is_repeated_header(values, col_names):
                counters['repeated_header_rows_skipped'] += 1
                continue
            is_summary = Tools._is_summary_row(values)
            if is_summary:
                counters['grand_total_rows'] += 1
            qualifying.append((values, is_summary))

        return qualifying, counters

    @staticmethod
    def _build_value_counts(
        col_names: list[str], qualifying: list[tuple[list, bool]], max_value_counts_per_column: int
    ) -> dict[str, list[dict]]:
        value_counts: dict[str, list[dict]] = {}
        for i, name in enumerate(col_names):
            counter = Counter(
                values[i]
                for values, is_summary in qualifying
                if not is_summary and not (values[i] is None or (isinstance(values[i], str) and not values[i].strip()))
            )
            if counter:
                value_counts[name or f'col_{i + 1}'] = [
                    {'value': v, 'count': c} for v, c in counter.most_common(max_value_counts_per_column)
                ]
        return value_counts

    @staticmethod
    def _summarize_rows(
        col_names: list[str],
        row_values: list[list],
        max_value_counts_per_column: int,
        sample_only: bool = False,
        known_total_data_rows: int | None = None,
    ) -> dict:
        """Classify raw data rows and build the preview + value_counts fields.

        ``row_values`` is exactly what should appear in the preview: callers
        that only physically read a bounded sample (``sample_only=True``) pass
        just that sample, not the whole sheet, so building it never requires
        materializing more rows than the preview will actually show. In that
        case ``known_total_data_rows`` (the true row count from a cheap
        metadata-only read) is used for ``data_row_count``/``preview_omitted_rows``
        instead of the sample size, so truncation is reported accurately rather
        than looking like a small file that happened to fit.
        """
        col_count = len(col_names)
        qualifying, counters = Tools._classify_rows(col_names, row_values)

        def _row_dict(values, is_summary):
            record = {(col_names[i] or f'col_{i + 1}'): values[i] for i in range(col_count)}
            if is_summary:
                record['is_summary_row'] = True
            return record

        preview = [_row_dict(values, is_summary) for values, is_summary in qualifying]
        value_counts = Tools._build_value_counts(col_names, qualifying, max_value_counts_per_column)

        if sample_only:
            data_row_count = known_total_data_rows if known_total_data_rows is not None else len(qualifying)
            preview_omitted_rows = max(0, data_row_count - len(preview))
        else:
            data_row_count = len(qualifying)
            preview_omitted_rows = 0

        result = {
            'data_row_count': data_row_count,
            **counters,
            'preview': preview,
            'preview_truncated': sample_only,
            'preview_omitted_rows': preview_omitted_rows,
            'value_counts': value_counts,
        }
        if sample_only:
            result['value_counts_from_sample'] = True
        return result

    @staticmethod
    def _structural_insights(rows, columns, col_names: list[str]) -> list[str]:
        """Return up to _MAX_INSIGHTS compact structural observations."""
        findings: list[str] = []
        if rows is not None:
            findings.append(f'{rows} data rows, {columns} columns')
            if rows > _MAX_SAMPLE_ROWS:
                findings.append(f'Large sheet (>{_MAX_SAMPLE_ROWS} rows); sample for analysis')
            elif rows == 0:
                findings.append('Sheet appears empty')
        else:
            findings.append(f'{columns} columns (row count not available)')
        if col_names:
            extra = columns - len(col_names) if columns > len(col_names) else 0
            summary = ', '.join(col_names)
            if extra:
                summary += f' … +{extra} more'
            findings.append(f'Columns: {summary}')
        return findings[:_MAX_INSIGHTS]
