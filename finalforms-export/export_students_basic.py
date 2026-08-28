#!/usr/bin/env python3
# /// script
# requires-python = ">=3.11"
# dependencies = [
#     "requests>=2.32",
#     "google-api-python-client>=2.140",
#     "google-auth>=2.34",
# ]
# ///
"""Trigger and download the FinalForms "Basic Student CSV" export, then upload it to Google Drive.

Designed to run headless in GitHub Actions on a schedule (see .github/workflows/finalforms-export.yml)
and locally for debugging. The FinalForms flow is plain HTML over a session cookie, no JavaScript:

  1. POST /staff/login with email/password and the CSRF token scraped from the login page
  2. GET  /students/background_export?export=students_basic&... to queue an async export
  3. Poll /background_exports/staff/{staff_id}/modal until a new download link appears
  4. GET  /background_exports/{id}/download and upload the CSV to the Drive exports folder

Drive uploads authenticate as madisonultimate@gmail.com via an OAuth refresh token (service
accounts have no storage quota and cannot create files in a My Drive folder). Run
authorize_drive.py to generate or regenerate the token.

Required environment variables:
  FINALFORMS_EMAIL, FINALFORMS_PASSWORD    staff login credentials
  GOOGLE_OAUTH_REFRESH_TOKEN               from authorize_drive.py
  plus the OAuth client: either a client_secret.json next to this script (path overridable
  via GOOGLE_OAUTH_CLIENT_FILE) or GOOGLE_OAUTH_CLIENT_ID + GOOGLE_OAUTH_CLIENT_SECRET
  FINALFORMS_SPORT_ID                      numeric sport id from the roster page URL (/sports/<id>);
                                           changes every season, so it is deliberately not defaulted
  FINALFORMS_SCHOOL_YEAR                   e.g. 2026 for the 2026-27 school year
  DRIVE_FOLDER_ID                          season's FinalForms exports folder in Drive
Optional:
  FINALFORMS_BASE_URL                      defaults to the SPS FinalForms host (stable across seasons)

For local runs, a .env file next to this script is loaded first (KEY=VALUE lines; real
environment variables win over .env values). GitHub Actions supplies real env vars instead.
"""

import csv
import io
import json
import os
import re
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from zoneinfo import ZoneInfo

import requests
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

SCRIPT_DIR = Path(__file__).resolve().parent


def load_dotenv():
    env_path = SCRIPT_DIR / ".env"
    if not env_path.exists():
        return
    for line in env_path.read_text().splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key, value = key.strip(), value.strip().strip("'\"")
        if key and value and key not in os.environ:
            os.environ[key] = value


load_dotenv()

BASE_URL = os.environ.get("FINALFORMS_BASE_URL", "https://seattleschools-wa.finalforms.com")
# Per-season values with no defaults: fail fast at startup if they haven't been set for the
# current season (locally in .env, in CI as GitHub Actions repository variables).
SPORT_ID = os.environ.get("FINALFORMS_SPORT_ID")
SCHOOL_YEAR = os.environ.get("FINALFORMS_SCHOOL_YEAR")
DRIVE_FOLDER_ID = os.environ.get("DRIVE_FOLDER_ID")

USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/126.0 Safari/537.36"
)
POLL_INTERVAL_SECONDS = 15
POLL_TIMEOUT_SECONDS = 600

# Downstream XLOOKUP formulas in the coach sheet depend on column positions, so fail loudly
# if SPS/FinalForms ever changes the export layout rather than uploading a surprise.
EXPECTED_LEADING_COLUMNS = ["StudentID", "School", "Residential School", "First Name", "Last Name"]
EXPECTED_COLUMNS = {
    "Are All Forms Parent Signed": 15,   # P
    "Are All Forms Student Signed": 16,  # Q
    "Gender": 20,                        # U
    "Grade": 22,                         # W
    "Date of Birth": 23,                 # X
    "Physical Clearance": 27,            # AB
    "Parent 1 First Name": 38,           # AM
    "Parent 1 Email": 40,                # AO
    "Parent 2 Email": 46,                # AU
}


def die(message):
    print(f"ERROR: {message}", file=sys.stderr)
    sys.exit(1)


def parse_form_inputs(html, action_path):
    """Return a dict of input name -> value for the form posting to action_path."""
    form_match = re.search(
        r'<form[^>]*action="' + re.escape(action_path) + r'"[^>]*>(.*?)</form>', html, re.S
    )
    if not form_match:
        die(f"could not find form with action {action_path} on the login page")
    fields = {}
    for input_tag in re.findall(r"<input[^>]*>", form_match.group(1)):
        name = re.search(r'name="([^"]*)"', input_tag)
        value = re.search(r'value="([^"]*)"', input_tag)
        if name and name.group(1) != "commit":
            fields[name.group(1)] = value.group(1) if value else ""
    return fields


def login(session, email, password):
    """Log in as staff; return the numeric staff id from the post-login redirect."""
    login_page = session.get(f"{BASE_URL}/staff/login")
    login_page.raise_for_status()
    fields = parse_form_inputs(login_page.text, "/staff/login")
    fields["email"] = email
    fields["password"] = password
    response = session.post(f"{BASE_URL}/staff/login", data=fields, allow_redirects=True)
    response.raise_for_status()

    landing = session.get(f"{BASE_URL}/staff")
    landing.raise_for_status()
    staff_match = re.search(r"/staff/(\d+)", landing.url)
    if not staff_match:
        die(
            "login did not reach a staff page; check FINALFORMS_EMAIL/FINALFORMS_PASSWORD "
            f"(landed on {landing.url})"
        )
    staff_id = staff_match.group(1)
    print(f"Logged in; staff id {staff_id}")
    return staff_id


def completed_export_ids(session, staff_id):
    """Export ids that currently have a download link (only completed exports get one)."""
    modal = session.get(f"{BASE_URL}/background_exports/staff/{staff_id}/modal?layout=modal")
    modal.raise_for_status()
    return {int(m) for m in re.findall(r"/background_exports/(\d+)/download", modal.text)}


def trigger_export(session):
    params = {
        "export": "students_basic",
        "layout": "partial",
        "sports.id_eq": SPORT_ID,
        "statuses.enrollment_status_in": "active,external,homeschooled",
        "statuses.school_year_eq": SCHOOL_YEAR,
    }
    response = session.get(f"{BASE_URL}/students/background_export", params=params)
    response.raise_for_status()
    print(f"Export queued for sport {SPORT_ID}, school year {SCHOOL_YEAR}")


def wait_for_new_export(session, staff_id, ids_before):
    deadline = time.monotonic() + POLL_TIMEOUT_SECONDS
    while time.monotonic() < deadline:
        time.sleep(POLL_INTERVAL_SECONDS)
        new_ids = completed_export_ids(session, staff_id) - ids_before
        if new_ids:
            export_id = max(new_ids)
            print(f"Export {export_id} complete")
            return export_id
        print("Waiting for export to complete...")
    die(f"export did not complete within {POLL_TIMEOUT_SECONDS} seconds")


def download_export(session, export_id):
    response = session.get(f"{BASE_URL}/background_exports/{export_id}/download")
    response.raise_for_status()
    return response.content


def validate_csv(data):
    """Assert the export still has the layout downstream formulas expect; return row count."""
    text = data.decode("utf-8-sig")
    rows = list(csv.reader(io.StringIO(text)))
    if len(rows) < 2:
        die(f"export has {len(rows)} rows; expected a header plus at least one student")
    header = rows[0]
    if header[: len(EXPECTED_LEADING_COLUMNS)] != EXPECTED_LEADING_COLUMNS:
        die(f"unexpected leading columns: {header[:6]}")
    for name, index in EXPECTED_COLUMNS.items():
        actual = header[index] if index < len(header) else "<missing>"
        if actual != name:
            die(f"expected column {index} to be {name!r}, found {actual!r}; export layout changed")
    return len(rows) - 1


def oauth_client_config():
    """Client id/secret from GOOGLE_OAUTH_CLIENT_FILE (a downloaded client_secret*.json,
    default ./client_secret.json next to this script) or GOOGLE_OAUTH_CLIENT_ID/SECRET."""
    file_setting = os.environ.get("GOOGLE_OAUTH_CLIENT_FILE", "client_secret.json")
    client_path = Path(file_setting)
    if not client_path.is_absolute():
        client_path = SCRIPT_DIR / client_path
    if client_path.exists():
        installed = json.loads(client_path.read_text())["installed"]
        return installed["client_id"], installed["client_secret"]
    client_id = os.environ.get("GOOGLE_OAUTH_CLIENT_ID")
    client_secret = os.environ.get("GOOGLE_OAUTH_CLIENT_SECRET")
    if client_id and client_secret:
        return client_id, client_secret
    die(
        f"no OAuth client found: provide {client_path} (or GOOGLE_OAUTH_CLIENT_FILE), "
        "or set GOOGLE_OAUTH_CLIENT_ID and GOOGLE_OAUTH_CLIENT_SECRET"
    )


def drive_client():
    client_id, client_secret = oauth_client_config()
    creds = Credentials(
        token=None,
        refresh_token=os.environ["GOOGLE_OAUTH_REFRESH_TOKEN"],
        client_id=client_id,
        client_secret=client_secret,
        token_uri="https://oauth2.googleapis.com/token",
        scopes=["https://www.googleapis.com/auth/drive"],
    )
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def upload_to_drive(data, filename):
    drive = drive_client()
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype="text/csv")
    existing = (
        drive.files()
        .list(
            q=f"name = '{filename}' and '{DRIVE_FOLDER_ID}' in parents and trashed = false",
            fields="files(id)",
        )
        .execute()
        .get("files", [])
    )
    if existing:
        drive.files().update(fileId=existing[0]["id"], media_body=media).execute()
        print(f"Updated existing {filename} in Drive")
    else:
        drive.files().create(
            body={"name": filename, "parents": [DRIVE_FOLDER_ID]},
            media_body=media,
            fields="id",
        ).execute()
        print(f"Uploaded {filename} to Drive")


def main():
    email = os.environ.get("FINALFORMS_EMAIL") or die("FINALFORMS_EMAIL is not set")
    password = os.environ.get("FINALFORMS_PASSWORD") or die("FINALFORMS_PASSWORD is not set")
    oauth_client_config()  # fail fast if no client id/secret source is available
    if not os.environ.get("GOOGLE_OAUTH_REFRESH_TOKEN"):
        die("GOOGLE_OAUTH_REFRESH_TOKEN is not set; run authorize_drive.py to generate it")
    for name, value in [
        ("FINALFORMS_SPORT_ID", SPORT_ID),
        ("FINALFORMS_SCHOOL_YEAR", SCHOOL_YEAR),
        ("DRIVE_FOLDER_ID", DRIVE_FOLDER_ID),
    ]:
        if not value:
            die(f"{name} is not set; these are per-season values, update them for the current season")

    session = requests.Session()
    session.headers["User-Agent"] = USER_AGENT

    staff_id = login(session, email, password)
    ids_before = completed_export_ids(session, staff_id)
    trigger_export(session)
    export_id = wait_for_new_export(session, staff_id, ids_before)
    data = download_export(session, export_id)
    student_count = validate_csv(data)

    # Include a full timestamp, not just the date: workflow_dispatch (the on-demand refresh
    # button) can run this more than once a day, and a date-only filename means a same-day
    # rerun silently overwrites the earlier export instead of leaving a trail. The consumer
    # (madison-ultimate's getMostRecentFileInfoFromFolder) picks the newest file by parsing
    # this exact "<date> <ISO8601 UTC timestamp>Z" pattern out of the filename.
    now_pt = datetime.now(ZoneInfo("America/Los_Angeles"))
    date_part = now_pt.strftime("%Y_%m_%d")
    timestamp_part = now_pt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    filename = f"students_basic_{date_part} {timestamp_part}.csv"
    upload_to_drive(data, filename)
    print(f"Done: {student_count} students in {filename}")


if __name__ == "__main__":
    main()
