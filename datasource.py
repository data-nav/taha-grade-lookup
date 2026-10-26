"""
Data sources. The engine only needs .values(sheet_name) -> 2D list of cell
values. This isolation is the whole point: Phase 1 reads the live Google Sheet,
and swapping to uploaded exports or a DB later means only this file changes.
"""

import threading
import time

import openpyxl


class LocalXlsxSource:
    """Reads sheet values out of a local .xlsx (used for testing and for the
    'uploaded export' mode). The workbook is parsed once at startup; restart
    the app after replacing the file."""

    def __init__(self, path):
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        self._cache = {ws.title: [list(row) for row in ws.iter_rows(values_only=True)]
                       for ws in wb.worksheets}
        wb.close()

    def values(self, sheet_name):
        return self._cache[sheet_name]


class GoogleSheetSource:
    """Phase-1 production source: reads via the Google Sheets API.

    All data sheets are fetched in ONE batchGet call and kept for `ttl` seconds,
    so a lookup costs at most one API request instead of one per program.
    Marks edited in the Sheet show up after at most `ttl` seconds (0 = always
    fetch fresh).

    Requires a service account with read access to the spreadsheet. Install:
        pip install google-api-python-client google-auth
    """

    def __init__(self, spreadsheet_id, sa_json_path, sheet_names, ttl=30):
        from google.oauth2.service_account import Credentials
        from googleapiclient.discovery import build
        creds = Credentials.from_service_account_file(
            sa_json_path, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
        self._svc = build("sheets", "v4", credentials=creds, cache_discovery=False)
        self._id = spreadsheet_id
        self._names = list(sheet_names)
        self._ttl = ttl
        self._cache = {}
        self._fetched_at = 0.0
        self._lock = threading.Lock()

    def _refresh(self):
        resp = self._svc.spreadsheets().values().batchGet(
            spreadsheetId=self._id,
            ranges=["'%s'" % n.replace("'", "''") for n in self._names],
            valueRenderOption="UNFORMATTED_VALUE",
            # Dates come back as serial numbers; render.py turns them into real dates
            dateTimeRenderOption="SERIAL_NUMBER",
        ).execute()
        self._cache = {n: vr.get("values", [])
                       for n, vr in zip(self._names, resp.get("valueRanges", []))}
        self._fetched_at = time.monotonic()

    def values(self, sheet_name):
        with self._lock:
            if not self._cache or time.monotonic() - self._fetched_at >= self._ttl:
                self._refresh()
            return self._cache[sheet_name]
