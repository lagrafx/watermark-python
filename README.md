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
pip install -e .
pip install pytest
```

Create `.env` from `.env.example` and fill in your values.

## Configuration

Environment variables:

- `AZURE_TENANT_ID`
- `AZURE_CLIENT_ID`
- `AZURE_CLIENT_SECRET`
- `SP_SITE_HOSTNAME` (example: `contoso.sharepoint.com`)
- `SP_SITE_PATH` (example: `/sites/Finance`)
- `SP_LIBRARY_NAMES` (optional, comma-separated library names; blank = all)
- `WATERMARK_IMAGE_PATH` (local PNG path)
- `STATE_FILE` (optional; default `.watermark_state.json`)

## Run

```powershell
python -m watermark_app --log-level INFO
```

Dry run:

```powershell
python -m watermark_app --dry-run --log-level DEBUG
```

## Schedule (Windows Task Scheduler)

Program/script:

`C:\Projects\watermark python\.venv\Scripts\python.exe`

Add arguments:

`-m watermark_app --log-level INFO`

Start in:

`C:\Projects\watermark python`

## Test

```powershell
pytest
```
