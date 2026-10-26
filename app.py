"""
Marksheet & Transcript Viewer - server version of the Apps Script web app.

doGet -> GET /, getStudentDocuments -> POST /api/lookup. Same response shape
the frontend already expects (ambiguous | studentId/studentName/program/
filename/marksheetBase64/transcriptBase64).

Production (connected to the live Google Sheet; marks are written into its
template sheets and Google makes the PDFs, exactly like the Apps Script):
    export MT_SHEET_ID=<spreadsheet id>
    export MT_SA_JSON=/opt/mt/service-account.json
    uvicorn app:app --host 127.0.0.1 --port 8000   # behind nginx

Offline test mode (no Google; reads an .xlsx export, renders with LibreOffice):
    export MT_LOCAL_XLSX=webhost.xlsx
    export MT_TEMPLATE=templates.xlsx              # from build_templates.py
    uvicorn app:app --host 0.0.0.0 --port 8000
"""

import base64
import os
import tempfile
import threading

from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse
from pydantic import BaseModel

from datasource import LocalXlsxSource, GoogleSheetSource

HERE = os.path.dirname(os.path.abspath(__file__))


def _make_backend():
    """Returns (data source, generate(id, programKey) -> result with PDF bytes)."""
    if os.environ.get("MT_SHEET_ID"):
        from gsheets import GoogleSheetRenderer
        sheet_id, sa_json = os.environ["MT_SHEET_ID"], os.environ["MT_SA_JSON"]
        source = GoogleSheetSource(sheet_id, sa_json)
        renderer = GoogleSheetRenderer(sheet_id, sa_json)
        return lambda sid, key: renderer.generate(source, sid, key)

    if os.environ.get("MT_LOCAL_XLSX"):
        import render
        source = LocalXlsxSource(os.environ["MT_LOCAL_XLSX"])
        template = os.environ.get("MT_TEMPLATE", "templates.xlsx")
        # LibreOffice uses one shared profile, so render one request at a time
        lock = threading.Lock()

        def generate(sid, key):
            # Each request works in its own temp folder, deleted once the PDFs are read
            with tempfile.TemporaryDirectory(prefix="mt_") as work, lock:
                res = render.generate(template, source, sid, key, out_dir=work)
                for k in ("marksheetPdf", "transcriptPdf"):
                    if k in res:
                        with open(res[k], "rb") as f:
                            res[k] = f.read()
                return res
        return generate

    raise RuntimeError("Set MT_SHEET_ID (+MT_SA_JSON) or MT_LOCAL_XLSX.")


GENERATE = _make_backend()

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
    try:
        res = GENERATE(body.id, body.programKey)
    except ValueError as e:               # expected, user-facing (not found, etc.)
        raise HTTPException(status_code=404, detail=str(e))
    except Exception as e:                # unexpected
        raise HTTPException(status_code=500, detail=str(e))

    if res.get("ambiguous"):
        return JSONResponse(res)

    return JSONResponse({
        "studentId": res["studentId"],
        "studentName": res["studentName"],
        "program": res["program"],
        "filename": res["filename"],
        "marksheetBase64": base64.b64encode(res["marksheetPdf"]).decode("ascii"),
        "transcriptBase64": base64.b64encode(res["transcriptPdf"]).decode("ascii"),
    })


@app.get("/healthz")
def healthz():
    return {"ok": True}
