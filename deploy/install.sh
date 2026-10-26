#!/usr/bin/env bash
# One-time server setup for the Marksheet & Transcript app (Ubuntu 22.04 / 24.04).
# Run it from the downloaded code:   sudo bash deploy/install.sh
# Running it again later installs the newest code and keeps your settings.
set -euo pipefail

if [ "$(id -u)" != 0 ]; then
  echo "Please run with sudo:  sudo bash deploy/install.sh"; exit 1
fi

SRC="$(cd "$(dirname "$0")/.." && pwd)"
BASE=/opt/mt
KEY="$BASE/service-account.json"
ENV_FILE="$BASE/marksheet.env"

if [ ! -s "$KEY" ]; then
  echo "Missing $KEY"
  echo "Save the Google service account key there first (Part C of the guide), then run this again."
  exit 1
fi

if [ ! -f "$ENV_FILE" ]; then
  read -rp "Spreadsheet ID (the long code in the Google Sheet's web address): " SHEET_ID
  [ -n "$SHEET_ID" ] || { echo "The spreadsheet ID is required."; exit 1; }
fi
read -rp "Domain name for the app, e.g. marks.tahacollege.ca (leave blank to use the server's IP for now): " DOMAIN

echo "== Installing system packages"
apt-get update -q
apt-get install -y -q python3-venv nginx certbot python3-certbot-nginx curl

echo "== Installing the app into $BASE/app"
id marksheet >/dev/null 2>&1 || useradd --system --home "$BASE" --shell /usr/sbin/nologin marksheet
mkdir -p "$BASE/app"
cp -r "$SRC"/. "$BASE/app/"
python3 -m venv "$BASE/venv"
"$BASE/venv/bin/pip" install -q --upgrade pip
"$BASE/venv/bin/pip" install -q -r "$BASE/app/requirements.txt"

if [ ! -f "$ENV_FILE" ]; then
  sed "s|^MT_SHEET_ID=.*|MT_SHEET_ID=$SHEET_ID|" "$BASE/app/deploy/marksheet.env.example" > "$ENV_FILE"
fi
chown -R marksheet: "$BASE"
chmod 600 "$ENV_FILE" "$KEY"

echo "== Starting the app"
cp "$BASE/app/deploy/marksheet.service" /etc/systemd/system/marksheet.service
systemctl daemon-reload
systemctl enable marksheet
systemctl restart marksheet

echo "== Setting up nginx"
sed "s|server_name .*;|server_name ${DOMAIN:-_};|" "$BASE/app/deploy/nginx.conf" > /etc/nginx/sites-available/marksheet
ln -sf ../sites-available/marksheet /etc/nginx/sites-enabled/marksheet
rm -f /etc/nginx/sites-enabled/default
nginx -t
systemctl reload nginx

for _ in $(seq 1 30); do
  curl -sf http://127.0.0.1:8000/healthz >/dev/null && break
  sleep 1
done
if ! curl -sf http://127.0.0.1:8000/healthz >/dev/null; then
  echo "The app did not start. See the error with:  sudo journalctl -u marksheet -n 50"
  exit 1
fi

if [ -n "$DOMAIN" ]; then
  echo "== Getting the HTTPS certificate for $DOMAIN"
  certbot --nginx -d "$DOMAIN"
  echo "Done. Open https://$DOMAIN"
else
  echo "Done. Open http://$(curl -s https://checkip.amazonaws.com || echo '<server IP>')"
fi
