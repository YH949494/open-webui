"""
title: Excel Analyzer
author: open-webui
version: 1.1.0
required_open_webui_version: 0.5.0
requirements: openpyxl, pandas
description: >
  Inspect spreadsheet files (.xlsx/.xlsm/.xls/.csv) a user has attached to the
  chat. The tool reads the attachment metadata from the reserved ``__files__``
  parameter, finds the first spreadsheet (by extension or MIME type), takes the
  storage path from the nested file record, resolves it to a readable local path
  via ``Storage.get_file`` (never ``open(file.path)`` directly), and then reads
  sheet names + row counts with openpyxl/pandas.

  Spreadsheets skip RAG (their text is never extracted or injected into the
  prompt), so this tool deliberately does NOT read ``file.data["content"]`` — it
  opens the stored file on disk instead.

INSTALLATION
------------
This is the source for a *custom* Open WebUI tool (built-in tools live in
``open_webui/tools/builtin.py``; this module is intentionally NOT auto-loaded).
To use it: Workspace -> Tools -> (+) -> paste this ``Tools`` class -> Save, then
attach a spreadsheet and ask:

  "Use the Excel analyzer tool to inspect the uploaded file.
   Return only sheet names and row counts."

WHY ``__files__`` MATTERS
-------------------------
Open WebUI only injects reserved parameters (``__files__``, ``__user__``,
``__id__``) into a tool when the function *declares them in its signature*
(``open_webui/utils/tools.py`` filters extra params to the signature). A tool
without ``__files__`` receives nothing about the upload. Each ``__files__``
entry looks like::

    {"type": "file", "id": "<file_id>", "name": "Foo.xlsx",
     "file": { ... full File record incl. "path" ... }}

``file.path`` is a *storage URI* (``s3://``/``gs://``/Azure URL, or — for local
/ Fly persistent-volume storage — an absolute path). ``Storage.get_file``
returns local paths as-is and downloads cloud objects into ``UPLOAD_DIR``, so
the result is always a path that exists on the local filesystem.
"""

from __future__ import annotations

import asyncio
import json
import logging
import os

from pydantic import BaseModel, Field

log = logging.getLogger(__name__)

# Spreadsheet detection by extension ...
_SPREADSHEET_EXTENSIONS = {'.xlsx', '.xlsm', '.xltx', '.xltm', '.xls', '.csv'}

# ... and by MIME type (content_type stored on the nested file record).
_SPREADSHEET_MIME_TYPES = {
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',  # xlsx
    'application/vnd.openxmlformats-officedocument.spreadsheetml.template',  # xltx
    'application/vnd.ms-excel',  # xls
    'application/vnd.ms-excel.sheet.macroenabled.12',  # xlsm
    'application/vnd.ms-excel.template.macroenabled.12',  # xltm
    'text/csv',
    'application/csv',
}


class Tools:
    class Valves(BaseModel):
        expose_paths_in_output: bool = Field(
            default=False,
            description=(
                'If true, include the resolved absolute server path in the tool output. '
                'Leave false in production so sensitive paths stay in the server logs only.'
            ),
        )

    def __init__(self):
        self.valves = self.Valves()
        # We read the raw uploaded file ourselves; do NOT let Open WebUI strip
        # the files from the request or treat this as a citation handler.
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
        sheet names and the row count of each sheet. Operates on files attached
        to the current chat — call it whenever the user asks to analyze, inspect
        or summarize an uploaded spreadsheet/Excel file.

        :param file_id: Optional id of a specific attached file. If omitted, the first spreadsheet is used.
        :return: JSON with the spreadsheet's sheet names and row counts, or a diagnostic if no file resolved.
        """
        files = [f for f in (__files__ or []) if isinstance(f, dict)]

        # (8) No attachments reached the tool at all.
        if not files:
            return json.dumps(
                {
                    'error': 'file not found; the file path was not accessible to the tool',
                    'reason': 'no files were passed to the tool (__files__ is empty)',
                    'attached_file_count': 0,
                    'filenames': [],
                    'file_ids': [],
                }
            )

        # (3) Find the first spreadsheet by extension or MIME type.
        target = self._select_spreadsheet(files, file_id)
        if target is None:
            return json.dumps(self._diagnostic(files, files[0], local_path=None, attempted=False))

        # (4) Take the storage path from the nested file record.
        file_record = target.get('file') or {}
        storage_path = file_record.get('path')
        display_name = target.get('name') or target.get('filename') or file_record.get('filename') or 'spreadsheet'

        # (5) Resolve via Storage.get_file — never open(file.path) directly.
        local_path = None
        if storage_path:
            local_path = await self._storage_get_file(storage_path)

        # Fallback: nested record had no usable path — resolve by file_id through
        # the access-checked helper (which itself uses Storage.get_file).
        if not (local_path and os.path.isfile(local_path)):
            fid = target.get('id')
            if fid:
                local_path = await self._resolve_by_id(fid, __user__)

        path_exists = bool(local_path and os.path.isfile(local_path))
        if not path_exists:
            return json.dumps(self._diagnostic(files, target, local_path=local_path, attempted=True))

        # Server-side debug log only — the absolute path is never returned.
        log.info(f"inspect_uploaded_spreadsheet: resolved '{display_name}' -> {local_path}")

        # (6) + (7) Read sheet names and row counts from the real local file.
        analysis = self._analyze(local_path, display_name)
        if self.valves.expose_paths_in_output:
            analysis['resolved_path'] = local_path
        return json.dumps(analysis, ensure_ascii=False)

    # --- resolution helpers ------------------------------------------------

    @staticmethod
    async def _storage_get_file(storage_path: str):
        """Resolve a storage URI/path to a local path via the storage provider."""
        try:
            from open_webui.storage.provider import Storage

            return await asyncio.to_thread(Storage.get_file, storage_path)
        except Exception as e:
            log.warning(f'inspect_uploaded_spreadsheet: Storage.get_file failed for {storage_path!r}: {e}')
            return None

    @staticmethod
    async def _resolve_by_id(file_id: str, __user__: dict | None):
        """Access-checked fallback resolver (also uses Storage.get_file)."""
        try:
            from open_webui.models.users import Users
            from open_webui.utils.files import resolve_uploaded_file_path

            user = None
            if __user__ and __user__.get('id'):
                user = await Users.get_user_by_id(__user__['id'])
            return await resolve_uploaded_file_path(file_id, user=user)
        except Exception as e:
            log.warning(f'inspect_uploaded_spreadsheet: id-based resolve failed for {file_id}: {e}')
            return None

    # --- selection / diagnostics ------------------------------------------

    def _select_spreadsheet(self, files: list, file_id: str):
        if file_id:
            for f in files:
                if f.get('id') == file_id:
                    return f
            return None
        for f in files:
            if self._is_spreadsheet(f):
                return f
        return None

    @classmethod
    def _is_spreadsheet(cls, item: dict) -> bool:
        file_record = item.get('file') or {}
        name = item.get('name') or item.get('filename') or file_record.get('filename') or ''
        if os.path.splitext(name.lower())[1] in _SPREADSHEET_EXTENSIONS:
            return True
        content_type = item.get('content_type') or (file_record.get('meta') or {}).get('content_type') or ''
        return content_type.split(';')[0].strip().lower() in _SPREADSHEET_MIME_TYPES

    @staticmethod
    def _diagnostic(files: list, target: dict, local_path, attempted: bool) -> dict:
        """Build the debuggable diagnostic required when resolution fails.

        Includes attached file count, filenames, file ids, the chosen item's
        top-level keys, the nested file record keys, and whether the path existed
        after Storage.get_file. Absolute paths are NOT included (logged only).
        """
        file_record = target.get('file') or {}
        diag = {
            'error': 'file not found; the file path was not accessible to the tool',
            'attached_file_count': len(files),
            'filenames': [f.get('name') or f.get('filename') or (f.get('file') or {}).get('filename') for f in files],
            'file_ids': [f.get('id') for f in files],
            'top_level_keys': sorted(target.keys()),
            'nested_file_keys': sorted(file_record.keys()),
            'path_exists_after_storage_get_file': bool(local_path and os.path.isfile(local_path)),
        }
        if not attempted:
            diag['reason'] = 'no spreadsheet attachment found (by extension or MIME type)'
        # Server-side only: include the resolved path in the log, never the output.
        log.warning(
            f'inspect_uploaded_spreadsheet: resolution failed; attempted_local_path={local_path!r} diagnostic={diag}'
        )
        return diag

    # --- reading -----------------------------------------------------------

    @staticmethod
    def _analyze(local_path: str, fname: str | None) -> dict:
        """Read sheet names and row counts using openpyxl (xlsx) / pandas (xls/csv).

        Never reads RAG-extracted ``file.data["content"]`` — always the file on disk.
        """
        ext = os.path.splitext((fname or local_path).lower())[1]
        display_name = fname or os.path.basename(local_path)

        try:
            if ext == '.csv':
                import pandas as pd

                df = pd.read_csv(local_path)
                return {'filename': display_name, 'sheets': [{'name': 'Sheet1', 'row_count': int(len(df))}]}

            if ext == '.xls':
                # Legacy binary format — openpyxl cannot read it; use pandas + xlrd.
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
                    row_count = ws.max_row if ws.max_row is not None else 0
                    sheets.append({'name': ws.title, 'row_count': int(row_count)})
                return {'filename': display_name, 'sheets': sheets}
            finally:
                wb.close()
        except Exception as e:
            log.exception(f'inspect_uploaded_spreadsheet: failed to read {local_path}: {e}')
            return {'filename': display_name, 'error': f'Failed to read spreadsheet: {e}'}
