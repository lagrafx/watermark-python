# Operations Runbook

## Purpose
Run and operate the SharePoint watermark automation safely in production.

## Repos
- Main source repo: `https://github.com/lagrafx/watermark-python`
- Portable deployment repo: `https://github.com/lagrafx/watermark-python-portable`

## Supported File Types
- `.docx`
- `.docm`
- `.xlsx`
- `.xlsm`
- `.pptx`
- `.pptm`
- `.pdf`

## Authentication
- Microsoft Graph app-only authentication.
- Current production path supports certificate auth via:
  - `AZURE_CLIENT_CERT_PFX_PATH`
  - `AZURE_CLIENT_CERT_PFX_PASSWORD` (if protected PFX)
- `CLOUD_ENV=gcch` for GCC High.

## Required Environment Variables
- `AZURE_TENANT_ID`
- `AZURE_CLIENT_ID`
- `AZURE_CLIENT_SECRET` (if using secret auth)
- `AZURE_CLIENT_CERT_PFX_PATH` (if using cert auth)
- `AZURE_CLIENT_CERT_PFX_PASSWORD` (optional if cert has no password)
- `CLOUD_ENV` (`commercial` or `gcch`)
- `SP_SITE_HOSTNAME`
- `SP_SITE_PATH`
- `SP_LIBRARY_NAMES`
- `SP_LIBRARY_WATERMARKS`
- `STATE_FILE` (optional, default `.watermark_state.json`)

## Day-1 Deploy (Portable)
1. Download latest zip from portable repo.
2. Unzip to fixed path, e.g. `C:\Apps\watermark-app`.
3. Copy `.env.example` to `.env` and populate values.
4. Confirm watermark PNG and cert file paths exist and are readable.
5. Run:
   - `.\watermark-app.exe --dry-run --log-level INFO`
   - `.\watermark-app.exe --log-level INFO`

## Schedule (Task Scheduler)
Use `Create Task` (not Basic Task):
- Run whether user is logged on or not.
- Run under service/ops account that can access local paths.
- Action: run `watermark-app.exe` with `--log-level INFO`.
- Start in: folder containing the exe and `.env`.

## Safe Update Procedure
1. Disable/stop scheduled task.
2. Back up:
   - `.env`
   - `.watermark_state.json` (or configured state file)
3. Replace executable/runtime from latest portable zip.
4. Restore `.env` and state file.
5. Run one manual `--dry-run`.
6. Re-enable scheduled task.

## Recovery Context (When Chat Is Lost)
Capture these in tickets or handoff notes:
- Current main repo commit hash.
- Current portable repo zip name deployed.
- Current auth mode (secret/cert) and cloud (`commercial`/`gcch`).
- Last known good run timestamp.
- Current blocker/error message.

Also maintain and review:
- `docs/SESSION_HISTORY.md` for curated conversation context snapshots.
