"""
title: Excel Analyzer
author: open-webui
version: 1.0.0
required_open_webui_version: 0.5.0
requirements: openpyxl
description: >
  Inspect spreadsheet files (.xlsx/.xlsm/.xls/.csv) that a user has attached to
  the chat. The tool resolves the *actual server-side path* of the uploaded
  file and reads it directly with openpyxl/pandas instead of relying on RAG
  extracted text. This is the companion tool to the "skip RAG for spreadsheets"
  upload path: spreadsheets are stored as raw data files (never chunked into the
  vector DB and never injected into the prompt), so a tool must open the file on
  disk to analyze it.

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

import json
import logging
import os

from pydantic import BaseModel, Field

log = logging.getLogger(__name__)

# Extensions this tool knows how to open directly.
_SPREADSHEET_EXTENSIONS = {'.xlsx', '.xlsm', '.xltx', '.xltm', '.xls', '.csv'}


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

    def __init__(self):
        self.valves = self.Valves()
        # We read the raw uploaded file ourselves; we do NOT want Open WebUI to
        # strip the files from the request or treat this as a citation handler.
        self.file_handler = False
        self.citation = False

    async def inspect_uploaded_spreadsheet(
        self,
        file_id: str = '',
        __files__: list | None = None,
        __user__: dict | None = None,
    ) -> str:
        """
        Inspect an uploaded spreadsheet (.xlsx/.xlsm/.xls/.csv) and return its
        sheet names and the row count of each sheet. Operates on files the user
        has attached to the current chat — call it whenever the user asks to
        analyze, inspect, or summarize an uploaded spreadsheet/Excel file.

        :param file_id: Optional id of a specific attached file. If omitted, the most recent spreadsheet is used.
        :return: JSON with each spreadsheet's sheet names and row counts, or a diagnostic if no file resolved.
        """
        # Resolve the user object once (access control needs id + role).
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
            # use it directly without a DB round-trip.  For legacy/pre-migration
            # file records where the DB path column is NULL this is the only
            # source of truth available.
            attachment_path = (item.get('file') or {}).get('path') or None

            local_path = await resolve_uploaded_file_path(fid, user=user, attachment_path=attachment_path)
            if not local_path:
                # Diagnostic: surface what we *did* receive so the failure is debuggable.
                results.append(self._missing_path_diagnostic(item, fid, fname))
                continue

            log.info(f'inspect_uploaded_spreadsheet: analyzing file_id={fid} at {local_path}')
            analysis = self._analyze(local_path, fname)
            if self.valves.expose_paths_in_output:
                analysis['resolved_path'] = local_path
            results.append(analysis)

        return json.dumps({'files': results}, ensure_ascii=False)

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
        """Pick which attached file(s) to inspect.

        Prefers an explicitly requested ``file_id``; otherwise returns all
        attached spreadsheet files (most-recent first, matching how Open WebUI
        orders attachments).
        """
        files = [f for f in (__files__ or []) if isinstance(f, dict)]

        if file_id:
            for f in files:
                if f.get('id') == file_id:
                    return [f]
            # Requested id wasn't in __files__ — still try to resolve it directly.
            return [{'id': file_id}]

        spreadsheets = [f for f in files if self._is_spreadsheet(f)]
        # Fall back to all files if extension detection found nothing (e.g. a
        # spreadsheet uploaded with an unusual/missing content type).
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
        """Build a debuggable diagnostic when a file path cannot be resolved.

        Per requirements this includes the file_id, filename, and the available
        metadata keys — but NOT any resolved absolute path.
        """
        file_meta = (item.get('file') or {}).get('meta') or {}
        return {
            'error': 'file not found; the file path was not accessible to the tool',
            'file_id': fid,
            'filename': fname,
            'available_item_keys': sorted(item.keys()),
            'available_meta_keys': sorted(file_meta.keys()),
        }

    @staticmethod
    def _analyze(local_path: str, fname: str | None) -> dict:
        """Read sheet names and row counts using openpyxl (xlsx) / pandas (xls/csv)."""
        ext = os.path.splitext((fname or local_path).lower())[1]
        display_name = fname or os.path.basename(local_path)

        try:
            if ext == '.csv':
                import pandas as pd

                df = pd.read_csv(local_path)
                return {
                    'filename': display_name,
                    'sheets': [{'name': 'Sheet1', 'row_count': int(len(df))}],
                }

            if ext == '.xls':
                # Legacy binary format — openpyxl cannot read it; use pandas+xlrd.
                import pandas as pd

                sheets = pd.read_excel(local_path, sheet_name=None, header=None)
                return {
                    'filename': display_name,
                    'sheets': [{'name': name, 'row_count': int(len(df))} for name, df in sheets.items()],
                }

            # Default: modern Excel via openpyxl in read-only (streaming) mode.
            import openpyxl

            wb = openpyxl.load_workbook(local_path, read_only=True, data_only=True)
            try:
                sheets = []
                for ws in wb.worksheets:
                    # max_row is reliable for real .xlsx files; guard against None.
                    row_count = ws.max_row if ws.max_row is not None else 0
                    sheets.append({'name': ws.title, 'row_count': int(row_count)})
                return {'filename': display_name, 'sheets': sheets}
            finally:
                wb.close()
        except Exception as e:
            log.exception(f'inspect_uploaded_spreadsheet: failed to read {local_path}: {e}')
            return {'filename': display_name, 'error': f'Failed to read spreadsheet: {e}'}
