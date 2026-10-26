# Marksheet & Transcript Viewer - EC2 port (Phase 1)

A drop-in replacement for the Google Apps Script web app. Same look, same
behaviour (enter a Student ID, get the marksheet + transcript preview and one
combined PDF). The difference is where it runs: compute, rendering and hosting
move off Apps Script.

## What changed vs Apps Script

- Marks logic: ported line-for-line from Code.gs (`engine.py`). Same PROGRAMS
  config, grade scales, aliases, credential rules, cell mappings.
- PDF rendering: Google's spreadsheet-to-PDF export is gone. The same .xlsx
  template sheets are filled with openpyxl and rendered by LibreOffice headless
  (`render.py`). Output is near-identical.
- Hosting: a small FastAPI app (`app.py`) serves the page and a `/api/lookup`
  endpoint. `doGet` -> `GET /`, `getStudentDocuments` -> `POST /api/lookup`.
- Frontend: `static/index.html` is the original, with one function changed
  (`google.script.run` -> `fetch`). Response shape is identical, so preview,
  merge and download are untouched.

## Files

    engine.py          marks logic (pure, no web/Google dependency)
    render.py          fill templates + LibreOffice PDF render + merge
    datasource.py      where the data comes from (local file, or live Sheet)
    app.py             FastAPI web app
    build_templates.py one-off: make the slim templates.xlsx
    static/index.html  frontend
    requirements.txt

## Local run (file mode, no Google)

    pip install -r requirements.txt
    sudo apt-get install -y libreoffice-calc        # or the full libreoffice
    python3 build_templates.py "Web Host.xlsx" templates.xlsx
    export MT_TEMPLATE=templates.xlsx
    export MT_LOCAL_XLSX="Web Host.xlsx"
    uvicorn app:app --host 0.0.0.0 --port 8000

Open http://localhost:8000 and enter a Student ID.

## Production (live Google Sheet) on EC2

The data still comes from the current Sheet, read through the Sheets API.

1. Service account: in Google Cloud, create a service account, enable the
   Sheets API, download its JSON key. Share the spreadsheet with the service
   account's email (Viewer is enough).

2. On the instance (Ubuntu):

        sudo apt-get update
        sudo apt-get install -y python3-venv libreoffice-calc fonts-liberation
        python3 -m venv /opt/mt/venv && /opt/mt/venv/bin/pip install -r requirements.txt
        # copy the full export once, build templates:
        /opt/mt/venv/bin/python build_templates.py "Web Host.xlsx" /opt/mt/templates.xlsx

3. Environment:

        MT_TEMPLATE=/opt/mt/templates.xlsx
        MT_SHEET_ID=<the spreadsheet id>
        MT_SA_JSON=/opt/mt/service-account.json

4. Run under systemd and put nginx in front. Ready-made files are in `deploy/`:

        sudo useradd --system --home /opt/mt marksheet
        sudo cp deploy/marksheet.env.example /opt/mt/marksheet.env   # then fill it in
        sudo cp deploy/marksheet.service /etc/systemd/system/
        sudo systemctl enable --now marksheet
        sudo cp deploy/nginx.conf /etc/nginx/sites-available/marksheet  # set your domain
        sudo ln -s ../sites-available/marksheet /etc/nginx/sites-enabled/
        sudo certbot --nginx -d <your domain>

5. Point your domain's A record at the instance (use an Elastic IP so it
   survives restarts). Security group: open 80 and 443 only, plus 22 for you.

## Notes / known tweaks

- The fonts package matters. Install `fonts-liberation` (and the family the
  templates use) so LibreOffice renders text the same as the samples.
- First render after a restart pays LibreOffice's cold start once (~30s). Every
  render after is ~2-3s. A warm-up call to `/healthz` isn't enough; hit
  `/api/lookup` once after deploy to prime it, or send one dummy conversion at
  startup.
- Page setup: each template sheet keeps its own orientation (marksheets
  landscape, transcripts portrait); only Letter size, fit-to-page and the
  print range are forced.
- Pillow is required. Without it openpyxl silently drops the logo and
  signature images from the templates.
- Data freshness: the live-Sheet source fetches all data sheets in one API call
  and reuses them for `MT_CACHE_SECONDS` (default 30). Set it to 0 to always
  read fresh (each lookup then waits on one Sheets API call).
- Concurrency: rendering is serialized by a lock (the LibreOffice profile is
  shared). Fine for staff lookups. If you ever need parallelism, give each
  worker its own profile dir.

## Access caveat

Standing up the EC2 instance needs AWS permissions to launch EC2 + configure
IAM/security groups. The SES/S3/CloudFront keys in the Colab notebook almost
certainly don't carry that. Confirm you can provision an instance (or use
Lightsail / a container host) before wiring the domain.

## Phase 2 (off Google entirely)

Only `datasource.py` changes. Swap `GoogleSheetSource` for a DB-backed or
uploaded-export source with the same `.values(sheet_name)` method. Nothing in
`engine.py`, `render.py` or `app.py` needs to change.
