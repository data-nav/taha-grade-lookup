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

## Setting it up on a server

You need:
- A Linux server (e.g. AWS EC2, Ubuntu) and a domain name if you want one.
- A Google Cloud **service account** (a "robot" Google account for the app):
  1. In Google Cloud Console, create a project, enable the **Google Sheets API**
     and **Google Drive API**.
  2. Create a service account and download its JSON key file.
  3. Share the spreadsheet with the service account's email as **Editor**
     (it writes into the template sheets, as the Apps Script did).

On the server, download this code, save the key to
`/opt/mt/service-account.json`, then run the one-command setup:

    sudo bash deploy/install.sh

It asks for the spreadsheet ID and (optionally) your domain, installs
everything, starts the app and sets up HTTPS. Run it again after `git pull`
to update; your settings are kept.

The full click-by-click guide, with diagrams of how it works, is in
`docs/system-design.html` (also shared as a PDF).

Point your domain's A record at the server (on EC2 use an Elastic IP).
Security group: open ports 80 and 443, plus 22 for yourself.

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
