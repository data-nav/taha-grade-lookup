"""
Marksheet & Transcript Viewer - EC2 backend (Phase 1).

Replaces the Apps Script web app. doGet -> GET /, getStudentDocuments ->
POST /api/lookup. Same response shape the frontend already expects
(ambiguous | studentId/studentName/program/filename/marksheetBase64/transcriptBase64),
so Index.html barely changes.

Run locally:
    export MT_TEMPLATE=templates.xlsx
    export MT_LOCAL_XLSX=webhost.xlsx          # test mode (reads data from a file)
    uvicorn app:app --host 0.0.0.0 --port 8000

Production (live Google Sheet):
    export MT_TEMPLATE=templates.xlsx
    export MT_SHEET_ID=<spreadsheet id>
    export MT_SA_JSON=/opt/mt/service-account.json
    uvicorn app:app --host 127.0.0.1 --port 8000   # behind nginx
"""

import base64
import os
import tempfile
import threading

from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse
from pydantic import BaseModel

import engine
import render
from datasource import LocalXlsxSource, GoogleSheetSource

TEMPLATE = os.environ.get("MT_TEMPLATE", "templates.xlsx")
HERE = os.path.dirname(os.path.abspath(__file__))

# Rendering touches a shared LibreOffice profile, so serialize it. This is the
# EC2 equivalent of the Apps Script LockService, but far lighter: each request
# already fills its own private workbook copy, so only the soffice step is guarded.
_render_lock = threading.Lock()


def _make_source():
    if os.environ.get("MT_SHEET_ID"):
        return GoogleSheetSource(os.environ["MT_SHEET_ID"], os.environ["MT_SA_JSON"],
                                 [p["dataSheet"] for p in engine.PROGRAMS],
                                 ttl=float(os.environ.get("MT_CACHE_SECONDS", "30")))
    if os.environ.get("MT_LOCAL_XLSX"):
        return LocalXlsxSource(os.environ["MT_LOCAL_XLSX"])
    raise RuntimeError("Set MT_SHEET_ID (+MT_SA_JSON) or MT_LOCAL_XLSX.")


SOURCE = _make_source()

app = FastAPI(title="Marksheet & Transcript")


class LookupIn(BaseModel):
    id: str
    programKey: str | None = None


@app.get("/", response_class=HTMLResponse)
def index():
    with open(os.path.join(HERE, "static", "index.html"), encoding="utf-8") as f:
        return f.read()


@app.post("/api/lookup")
def lookup(body: LookupIn):
    # Each request works in its own temp folder, deleted once the PDFs are read
    with tempfile.TemporaryDirectory(prefix="mt_") as work:
        try:
            with _render_lock:
                res = render.generate(TEMPLATE, SOURCE, body.id, body.programKey, out_dir=work)
        except ValueError as e:               # expected, user-facing (not found, etc.)
            raise HTTPException(status_code=404, detail=str(e))
        except Exception as e:                # unexpected
            raise HTTPException(status_code=500, detail=str(e))

        if res.get("ambiguous"):
            return JSONResponse(res)

        def b64(path):
            with open(path, "rb") as f:
                return base64.b64encode(f.read()).decode("ascii")

        return JSONResponse({
            "studentId": res["studentId"],
            "studentName": res["studentName"],
            "program": res["program"],
            "filename": res["filename"],
            "marksheetBase64": b64(res["marksheetPdf"]),
            "transcriptBase64": b64(res["transcriptPdf"]),
        })


@app.get("/healthz")
def healthz():
    return {"ok": True}
