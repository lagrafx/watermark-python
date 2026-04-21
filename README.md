# watermark-python

Scheduled SharePoint watermark automation for new Office/PDF documents.

## Project Docs

- Operations runbook: `docs/OPERATIONS.md`
- Troubleshooting guide: `docs/TROUBLESHOOTING.md`
- Change history: `docs/CHANGELOG.md`

## What It Does

- Connects to a SharePoint site through Microsoft Graph (app-only auth).
- Scans one or more document libraries.
- Finds newly created `.docx/.docm/.xlsx/.xlsm/.pptx/.pptm/.pdf` files since the last successful run.
- Applies a PNG watermark to each file.
- Uploads the updated file back to SharePoint.
- Saves run state to a local JSON file.

## Prerequisites

- Python 3.10+
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
- `SP_LIBRARY_WATERMARKS` (required, `Library=path;Other Library=path` mapping)
- `STATE_FILE` (optional; default `.watermark_state.json`)

Example:

`SP_LIBRARY_WATERMARKS=Documents=C:\Projects\watermark python\classified_watermark.png;Legal Docs=C:\Projects\watermark python\legal_watermark.png`

Safety behavior:

- Every targeted library must have an explicit entry in `SP_LIBRARY_WATERMARKS`.
- The run fails fast if any targeted library is missing a mapping, to prevent accidental watermarking.

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

Note: `--dry-run` does not update the run-state file.

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

## Portable EXE Bundle (No Python Install On VM)

If the target VM cannot install Python or run pip against the internet, build a portable bundle on another machine:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-portable.ps1
```

This produces a timestamped folder and zip in `C:\Projects` containing:

- `watermark-app.exe`
- `_internal\` runtime payload
- `.env.example`
- watermark PNG assets
- `RUN_ME_FIRST.txt`

On VM:

1. Unzip to `C:\Apps\watermark-app`
2. Copy `.env.example` to `.env` and fill values
3. Run:
   `.\watermark-app.exe --dry-run --log-level INFO`

## Test

```powershell
pytest
```

Lint:

```powershell
ruff check src tests
```
