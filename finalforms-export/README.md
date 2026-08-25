# finalforms-export

Exports the SPS FinalForms "Basic Student CSV" for the season's ultimate sport and uploads it to Google Drive, where the player portal and coach sheet ingest it as "the newest CSV in the folder." Replaces the manual routine of logging into FinalForms, clicking Export → Basic Student CSV, downloading the file, and dropping it in Drive.

How it runs (locally, on a schedule, or anywhere else) is not this tool's concern; it is a single script that reads configuration from the environment and exits nonzero on any failure. This repo schedules it via GitHub Actions (see below), but nothing here depends on that.

## How it works

The whole FinalForms flow is plain HTML over a session cookie (no JavaScript), so this is a small `requests` script rather than a headless browser:

1. `POST /staff/login` with email/password and the CSRF token scraped from the login page
2. `GET /students/background_export?export=students_basic&sports.id_eq=<SPORT_ID>&statuses.enrollment_status_in=active,external,homeschooled&statuses.school_year_eq=<YEAR>` queues an async export (exactly what the Export → Basic Student CSV click does)
3. Poll `GET /background_exports/staff/<staff_id>/modal?layout=modal` until a new download link appears (only completed exports get one)
4. `GET /background_exports/<id>/download`, validate the column layout, and upload to the Drive exports folder as `students_basic_YYYY_MM_DD.csv` (Pacific date; re-running the same day overwrites that day's file)

The script validates the header row against the column positions the coach sheet's XLOOKUP formulas and the portal parser depend on (StudentID in A, signature flags in P/Q, DOB in X, physical clearance in AB, parent contact in AM-AO/AS-AU) and fails loudly if FinalForms ever changes the layout. Any failure (login rejected, missing config, export timeout, layout change, Drive upload error) is a nonzero exit; whatever runs the script decides how to surface that. The fallback is always the manual process this replaces.

## Drive authentication

Uploads run as `madisonultimate@gmail.com` via an OAuth refresh token, not the service account: Google service accounts have no storage quota and cannot create files in a My Drive folder, and we want real per-day files in the folder. One-time setup in the Google Cloud console (madisonultimate project):

1. Google Auth Platform / OAuth consent screen: configure as External and **Publish app**. Leaving it in Testing makes refresh tokens expire every 7 days; published-but-unverified just shows an "unverified app" warning during the one consent click, which is fine for our own app
2. Credentials → Create credentials → OAuth client ID → **Desktop app**; download its JSON and save it as `client_secret.json` in this directory (gitignored). Alternatively set `GOOGLE_OAUTH_CLIENT_ID`/`GOOGLE_OAUTH_CLIENT_SECRET` in `.env`; a `GOOGLE_OAUTH_CLIENT_FILE` path overrides the default file location

Then generate (or later regenerate) the refresh token:

    uv run authorize_drive.py               # browser consent as madisonultimate@gmail.com; writes .env
    uv run authorize_drive.py --gh-secrets  # same, and also pushes the three values to GitHub secrets

The token stays valid as long as it is used (a nightly run keeps it alive) and survives until revoked or a password change. If a run ever fails with `invalid_grant`, re-run `authorize_drive.py --gh-secrets`.

## Configuration

All via environment variables. Dependencies are declared inline in the scripts (PEP 723) and managed by [uv](https://docs.astral.sh/uv/): `uv run <script>` resolves everything automatically.

Credentials:

- `FINALFORMS_EMAIL` / `FINALFORMS_PASSWORD`: staff login
- Drive access as madisonultimate@gmail.com (see above): `client_secret.json` (or `GOOGLE_OAUTH_CLIENT_ID`/`GOOGLE_OAUTH_CLIENT_SECRET`) plus `GOOGLE_OAUTH_REFRESH_TOKEN`. CI uses the three env values, which `authorize_drive.py --gh-secrets` extracts and pushes for you

Per-season values, deliberately without code defaults so a stale season fails fast:

- `FINALFORMS_SPORT_ID`: the sport's numeric id, visible in the roster page URL (`/sports/<id>`); `2605` for Fall 2026
- `FINALFORMS_SCHOOL_YEAR`: `2026` for the 2026-27 school year
- `DRIVE_FOLDER_ID`: the season's FinalForms exports folder

## Running locally

Edit `.env` in this directory (gitignored; the script loads it automatically, and real environment variables win over `.env` values), then:

    uv run export_students_basic.py

## Scheduling via GitHub Actions

`.github/workflows/finalforms-export.yml` runs the script nightly (6am Pacific) and on manual dispatch. Its setup, all under repo Settings → Secrets and variables → Actions:

- Secrets: `FINALFORMS_EMAIL`, `FINALFORMS_PASSWORD`, plus the three `GOOGLE_OAUTH_*` values (`authorize_drive.py --gh-secrets` sets those)
- Variables: `FINALFORMS_SPORT_ID`, `FINALFORMS_SCHOOL_YEAR`, `DRIVE_FOLDER_ID`, e.g.:

      gh variable set FINALFORMS_SPORT_ID --body 2605 -R BigLep/madison-ultimate-admin
      gh variable set FINALFORMS_SCHOOL_YEAR --body 2026 -R BigLep/madison-ultimate-admin
      gh variable set DRIVE_FOLDER_ID --body 1WgD4hY0fIZlQEBt7ekOlHIECA-HgOMIZ -R BigLep/madison-ultimate-admin

GitHub emails the repo owner when a scheduled run fails. Two platform caveats: schedules only run from the default branch, and GitHub auto-disables scheduled workflows after ~60 days without repo activity, so re-enabling from the Actions tab is part of each season's bootup.

## Per-season checklist

1. Create the new season's exports folder in Drive
2. Update the per-season values wherever the script runs (local `.env`, GitHub repo variables): new sport id, school year, folder id
3. Re-enable the scheduled workflow if GitHub disabled it during the off-season
