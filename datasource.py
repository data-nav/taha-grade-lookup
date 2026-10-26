"""
Data sources. The engine needs a 2D list of cell values per data sheet, with
the two header rows first. A source provides either

    values(sheet_name)              the whole sheet, or
    rows_for(student_id, sheets)    {sheet_name: [header, subheader, matching row]}
                                    (just the header rows when the ID is not in that sheet)

The live Google Sheet source uses the second form so that every lookup reads
fresh from the Sheet without pulling every sheet down.
"""

import openpyxl

import engine as E


class LocalXlsxSource:
    """Reads sheet values out of a local .xlsx (offline testing). The workbook
    is parsed once at startup; restart the app after replacing the file."""

    def __init__(self, path):
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        self._cache = {ws.title: [list(row) for row in ws.iter_rows(values_only=True)]
                       for ws in wb.worksheets}
        wb.close()

    def values(self, sheet_name):
        return self._cache[sheet_name]


class GoogleSheetSource:
    """Production source: reads the live Google Sheet through the Sheets API.

    Nothing is cached. Each lookup makes three small requests that cover all
    programs at once: the header rows, the Student ID columns, and the matching
    row of each sheet that has the ID. Marks typed into the Sheet are used by
    the very next lookup, exactly like the Apps Script.

    Requires a service account with read access to the spreadsheet.
    """

    def __init__(self, spreadsheet_id, sa_json_path):
        from google.oauth2.service_account import Credentials
        from googleapiclient.discovery import build
        creds = Credentials.from_service_account_file(
            sa_json_path, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
        self._svc = build("sheets", "v4", credentials=creds, cache_discovery=False)
        self._id = spreadsheet_id

    def _batch_get(self, ranges):
        """-> one 2D list per range, in order (empty list for an empty range)."""
        from googleapiclient.errors import HttpError
        if not ranges:
            return []
        try:
            resp = self._svc.spreadsheets().values().batchGet(
                spreadsheetId=self._id, ranges=ranges,
                valueRenderOption="UNFORMATTED_VALUE",
                dateTimeRenderOption="SERIAL_NUMBER",     # dates as serial numbers
            ).execute()
        except HttpError as e:
            if e.resp.status == 400:          # a sheet name that does not exist
                raise ValueError("A data sheet was not found: " + ", ".join(ranges))
            raise
        return [vr.get("values", []) for vr in resp.get("valueRanges", [])]

    def values(self, sheet_name):
        return self._batch_get([_q(sheet_name)])[0]

    def rows_for(self, student_id, sheets):
        target = E.id_str_(student_id)
        sheets = list(sheets)

        # 1. header rows of every sheet -> where the Student ID column is
        heads = self._batch_get(["%s!%d:%d" % (_q(s), E.HEADER_ROW, E.SUBHEADER_ROW) for s in sheets])
        out = {s: h for s, h in zip(sheets, heads)}
        id_col = {}
        for s, h in out.items():
            header = h[0] if h else []
            for i, v in enumerate(header):
                if E.norm_(v) == E.norm_(E.FIXED_HEADERS["id"]):
                    id_col[s] = i
                    break
            # A sheet without the column stays as its header rows only; the
            # engine then reports the missing column, as Code.gs does.

        # 2. the Student ID columns -> which row (if any) holds the ID
        with_col = [s for s in sheets if s in id_col]
        cols = self._batch_get(["%s!%s%d:%s" % (_q(s), _col(id_col[s]), E.FIRST_DATA_ROW, _col(id_col[s]))
                                for s in with_col])
        hit = {}
        for s, col in zip(with_col, cols):
            for i, r in enumerate(col):
                if r and E.id_str_(r[0]) == target:
                    hit[s] = E.FIRST_DATA_ROW + i
                    break

        # 3. those rows
        found = list(hit)
        rows = self._batch_get(["%s!%d:%d" % (_q(s), hit[s], hit[s]) for s in found])
        for s, r in zip(found, rows):
            out[s] = out[s] + [r[0] if r else []]
        return out


def _q(sheet_name):
    return "'%s'" % sheet_name.replace("'", "''")


def _col(index0):
    from openpyxl.utils import get_column_letter
    return get_column_letter(index0 + 1)
