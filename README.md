# watermark-python

Scheduled SharePoint watermark automation for new Word/Excel documents.

## What It Does

- Connects to a SharePoint site through Microsoft Graph (app-only auth).
- Scans one or more document libraries.
- Finds newly created `.docx/.docm/.xlsx/.xlsm` files since the last successful run.
- Applies a PNG watermark to each file.
- Uploads the updated file back to SharePoint.
- Saves run state to a local JSON file.

## Prerequisites

- Python 3.13+
- Azure App Registration with Graph application permissions to SharePoint files.
- Admin consent granted for the app permissions.

## Setup

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -e .
pip install pytest
```

Create `.env` from `.env.example` and fill in your values.

## Deploy On A Server

1. Place the repo in a fixed folder, for example: `C:\Apps\watermark-python`
2. Open PowerShell in that folder.
3. Create and install the runtime environment:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -e .
```

4. Create `C:\Apps\watermark-python\.env` from `.env.example` and set real values.
5. Run a verification test:

```powershell
.\.venv\Scripts\python.exe -m watermark_app --dry-run --log-level INFO
```

6. Run a real test:

```powershell
.\.venv\Scripts\python.exe -m watermark_app --log-level INFO
```

## Configuration

Environment variables:

- `AZURE_TENANT_ID`
- `AZURE_CLIENT_ID`
- `AZURE_CLIENT_SECRET` (required for client secret auth)
- `AZURE_CLIENT_CERT_PFX_PATH` (required for certificate auth)
- `AZURE_CLIENT_CERT_PFX_PASSWORD` (optional; required only if the PFX is password protected)
- `CLOUD_ENV` (optional; `commercial` default, or `gcch`)
- `SP_SITE_HOSTNAME` (example: `contoso.sharepoint.com`)
- `SP_SITE_PATH` (example: `/sites/Finance`)
- `SP_LIBRARY_NAMES` (optional, comma-separated library names; blank = all)
- `SP_LIBRARY_WATERMARKS` (optional, `Library=path;Other Library=path` mapping)
- `WATERMARK_IMAGE_PATH` (local PNG path)
- `STATE_FILE` (optional; default `.watermark_state.json`)

`WATERMARK_IMAGE_PATH` remains the default watermark for any library not listed in
`SP_LIBRARY_WATERMARKS`.

Example:

`SP_LIBRARY_WATERMARKS=Documents=C:\Projects\watermark python\classified_watermark.png;Legal Docs=C:\Projects\watermark python\legal_watermark.png`

For GCC High, set:

- `CLOUD_ENV=gcch`
- `SP_SITE_HOSTNAME` typically ends with `.sharepoint.us`

Authentication mode:

- Set `AZURE_CLIENT_SECRET` to use client secret authentication.
- Set `AZURE_CLIENT_CERT_PFX_PATH` (and optional `AZURE_CLIENT_CERT_PFX_PASSWORD`) to use
  certificate authentication from a PFX file.
- If both are set, certificate authentication is used.

## Run

```powershell
python -m watermark_app --log-level INFO
```

Dry run:

```powershell
python -m watermark_app --dry-run --log-level DEBUG
```

## Schedule (Windows Task Scheduler)

Use **Task Scheduler -> Create Task** (not Basic Task).

General:

- Run whether user is logged on or not.
- Use a service account that has access to local files and network.

Trigger:

- Set your desired schedule (daily/hourly/etc.).

Action:

- Program/script:
  `C:\Apps\watermark-python\.venv\Scripts\python.exe`
- Add arguments:
  `-m watermark_app --log-level INFO`
- Start in:
  `C:\Apps\watermark-python`

Optional log file output:

- Add arguments:
  `-m watermark_app --log-level INFO >> C:\Apps\watermark-python\logs\watermark.log 2>&1`

Dependency note:

- `Pillow` and other required packages are already included in project dependencies.
- Installing with `pip install -e .` installs everything needed; no separate manual install is required.

## Test

```powershell
pytest
```
