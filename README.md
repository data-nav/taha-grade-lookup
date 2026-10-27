# Marksheet & Transcript Viewer (server version)

The same app as the Google Apps Script web app: enter a Student ID, get the
marksheet + transcript preview and one combined PDF. It runs on your own server
(e.g. AWS EC2) instead of inside Apps Script.

## How it works

It stays connected to the Google Sheet, just like the Apps Script:

1. Finds the student in the program data sheets (IBA, HT, CSW, MOA, AC2 (New),
   AC2 (Old), PSW, ECA).
2. Writes their marks into **that program's** Marksheet and Transcript sheets
   in the Google Sheet (e.g. AC2 (Old) -> "AC2 (Old) Marksheet" and
   "AC2 (Old) Transcript").
3. Asks Google to turn those two sheets into PDFs, and shows them.

Because the template sheets are read from the Google Sheet on every lookup,
any change you make to a template in Google Sheets is used immediately, and
the PDFs look exactly like the Apps Script ones (Google makes them).

The marks logic in `engine.py` is a line-for-line port of Code.gs; it was
checked against the original Code.gs on every student in the workbook
(9,507 students, 0 differences).

## Adding a new program

Same as with the Apps Script:

1. In the Google Sheet, add the program's data sheet plus its
   "<Name> Marksheet" and "<Name> Transcript" sheets.
2. Add one entry for it to `PROGRAMS` in `engine.py` (copy the closest
   existing program and change the sheet names and cell positions), then
   restart the app (`sudo systemctl restart marksheet`).

## Files

    engine.py          marks logic and program settings (port of Code.gs)
    gsheets.py         fills the Google Sheet templates and exports the PDFs
    datasource.py      reads the data sheets (Google Sheet, or a local .xlsx)
    app.py             the web app
    static/index.html  the page (Index.html with google.script.run -> fetch)
    deploy/            server setup files (systemd, nginx, settings)
    apps-script/       the original Apps Script, for reference
    render.py, build_templates.py   offline test mode only (see below)

## Hosting it

You need a Google **service account** either way (a "robot" Google account the
app signs in as). Steps are in `docs/system-design.html`, Part A. In short:
Google Cloud Console -> new project -> enable Google Sheets API and Google
Drive API -> create a service account -> create a JSON key -> share the
spreadsheet with the service account's email as **Editor** -> copy the
spreadsheet ID from the Sheet's web address.

### Easiest on AWS: App Runner (runs it from GitHub, no server to manage)

1. Sign in to the AWS console, search **App Runner** -> **Create service**.
2. Source: **Source code repository** -> **Add new** connection -> sign in to
   GitHub and allow access to `taha-grade-lookup`. Branch: `ec2-port`.
   Deployment trigger: **Automatic** (every push to GitHub redeploys).
3. Build settings: **Configure all settings here**. Runtime **Python 3.11**.
   Build command: `pip install -r requirements.txt`
   Start command: `uvicorn app:app --host 0.0.0.0 --port 8080`   Port: `8080`
4. Service name `marksheet`. Smallest size (0.25 vCPU, 0.5 GB) is enough.
   Environment variables: `MT_SHEET_ID` = the spreadsheet ID, and
   `MT_SA_JSON_CONTENT` = the whole text of the service account's .json key.
   Health check: path `/healthz`.
5. **Create & deploy**. A few minutes later it prints a web address
   (`https://xxxx.awsapprunner.com`). Add your own domain under **Custom domains**.

### Also easy: Render (same idea, outside AWS)

Sign up at render.com with GitHub -> **New -> Blueprint** -> pick this repo
(it reads `render.yaml`) -> enter `MT_SHEET_ID` and `MT_SA_JSON_CONTENT` ->
**Apply**. About $7/month.

### Alternative: your own server (AWS EC2 etc.)

Ubuntu server, download this code, save the key to
`/opt/mt/service-account.json`, then run the one-command setup:

    sudo bash deploy/install.sh

It asks for the spreadsheet ID and (optionally) your domain, installs
everything, starts the app and sets up HTTPS. Run it again after `git pull`
to update; your settings are kept.

The full click-by-click guide, with diagrams of how it works, is in
`docs/system-design.html` (also shared as a PDF).

## Notes

- Data freshness: every lookup reads the Google Sheet fresh (three small
  requests: the header rows, the Student ID columns, and the student's row),
  so a mark typed into the Sheet is used by the very next lookup. Nothing is
  cached and no copy of the data is kept on the server.
- One lookup at a time: the template sheets are shared, so lookups take turns
  (like the Apps Script lock). If the Apps Script web app is still in use at the
  same time, the two can't see each other's lock, so retire the old web app
  link once the server version is live.
- Access: like the current deployment (`ANYONE_ANONYMOUS`), anyone with the
  link can look up any ID. Put the server behind a login if that's not wanted.

## Offline test mode (no Google)

For testing without Google credentials: reads an .xlsx export and renders the
PDFs with LibreOffice. The look is close to, but not exactly, Google's.

    sudo apt-get install -y libreoffice-calc fonts-liberation
    python3 build_templates.py "Web Host.xlsx" templates.xlsx
    MT_LOCAL_XLSX="Web Host.xlsx" MT_TEMPLATE=templates.xlsx uvicorn app:app --port 8000
