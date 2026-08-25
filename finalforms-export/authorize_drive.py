#!/usr/bin/env python3
# /// script
# requires-python = ">=3.11"
# dependencies = [
#     "google-auth-oauthlib>=1.2",
# ]
# ///
"""Generate (or regenerate) the Google Drive OAuth refresh token for export_students_basic.py.

Runs the standard installed-app consent flow in your browser as madisonultimate@gmail.com,
then writes GOOGLE_OAUTH_REFRESH_TOKEN into the .env file next to this script. With
--gh-secrets it also pushes the client id, client secret, and refresh token to GitHub
Actions secrets via the gh CLI.

Prerequisites (one-time, in the Google Cloud console for the madisonultimate project):
  1. Google Auth Platform / OAuth consent screen: configure as External and click
     "Publish app" (leaving it in Testing makes refresh tokens expire after 7 days;
     published-but-unverified only means an "unverified app" warning during consent)
  2. Credentials -> Create credentials -> OAuth client ID -> Desktop app; put the
     client id and secret in .env as GOOGLE_OAUTH_CLIENT_ID / GOOGLE_OAUTH_CLIENT_SECRET

Usage:
  uv run authorize_drive.py [--gh-secrets [--repo OWNER/REPO]]
"""

import argparse
import json
import os
import subprocess
import sys
from pathlib import Path

from google_auth_oauthlib.flow import InstalledAppFlow

SCRIPT_DIR = Path(__file__).resolve().parent
ENV_PATH = SCRIPT_DIR / ".env"
SCOPES = ["https://www.googleapis.com/auth/drive"]
DEFAULT_REPO = "BigLep/madison-ultimate-admin"


def load_client_config():
    """Client id/secret from GOOGLE_OAUTH_CLIENT_FILE (a downloaded client_secret*.json,
    default ./client_secret.json) or from GOOGLE_OAUTH_CLIENT_ID/SECRET env vars."""
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
    sys.exit(
        f"No OAuth client found: put the downloaded client secret JSON at {client_path} "
        "(or set GOOGLE_OAUTH_CLIENT_FILE), or set GOOGLE_OAUTH_CLIENT_ID and "
        "GOOGLE_OAUTH_CLIENT_SECRET in .env"
    )


def load_dotenv():
    if not ENV_PATH.exists():
        return
    for line in ENV_PATH.read_text().splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key, value = key.strip(), value.strip().strip("'\"")
        if key and value and key not in os.environ:
            os.environ[key] = value


def set_env_value(key, value):
    """Update or append key=value in .env, leaving all other lines untouched."""
    lines = ENV_PATH.read_text().splitlines() if ENV_PATH.exists() else []
    replaced = False
    for i, line in enumerate(lines):
        if line.split("=", 1)[0].strip() == key and not line.lstrip().startswith("#"):
            lines[i] = f"{key}={value}"
            replaced = True
            break
    if not replaced:
        lines.append(f"{key}={value}")
    ENV_PATH.write_text("\n".join(lines) + "\n")


def push_gh_secrets(repo, client_id, client_secret, refresh_token):
    for name, value in [
        ("GOOGLE_OAUTH_CLIENT_ID", client_id),
        ("GOOGLE_OAUTH_CLIENT_SECRET", client_secret),
        ("GOOGLE_OAUTH_REFRESH_TOKEN", refresh_token),
    ]:
        subprocess.run(
            ["gh", "secret", "set", name, "-R", repo],
            input=value.encode(),
            check=True,
        )
        print(f"Set GitHub secret {name}")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--gh-secrets",
        action="store_true",
        help="also push client id/secret and the new refresh token to GitHub Actions secrets",
    )
    parser.add_argument("--repo", default=DEFAULT_REPO, help=f"GitHub repo (default {DEFAULT_REPO})")
    args = parser.parse_args()

    load_dotenv()
    client_id, client_secret = load_client_config()

    flow = InstalledAppFlow.from_client_config(
        {
            "installed": {
                "client_id": client_id,
                "client_secret": client_secret,
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "redirect_uris": ["http://localhost"],
            }
        },
        scopes=SCOPES,
    )
    print("A browser window will open; sign in as madisonultimate@gmail.com.")
    print('If Google shows an "unverified app" warning, use Advanced -> Continue.')
    creds = flow.run_local_server(port=0, access_type="offline", prompt="consent")

    if not creds.refresh_token:
        sys.exit("No refresh token returned; re-run and make sure the consent prompt appears")

    set_env_value("GOOGLE_OAUTH_REFRESH_TOKEN", creds.refresh_token)
    print(f"Wrote GOOGLE_OAUTH_REFRESH_TOKEN to {ENV_PATH}")

    if args.gh_secrets:
        push_gh_secrets(args.repo, client_id, client_secret, creds.refresh_token)

    print("Done. Verify with: uv run export_students_basic.py")


if __name__ == "__main__":
    main()
