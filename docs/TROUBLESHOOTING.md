# Troubleshooting

## Quick Triage Checklist
1. Confirm `.env` is present in the runtime folder.
2. Confirm `CLOUD_ENV` matches tenant (`gcch` vs `commercial`).
3. Confirm certificate file path is valid and readable.
4. Confirm `SP_LIBRARY_WATERMARKS` has entries for all targeted libraries.
5. Confirm targeted library names match Graph drive display names.
6. Confirm app has Graph permission and required site grants (`Sites.Selected` model).

## Common Errors

### `Missing SP_LIBRARY_WATERMARKS entries for targeted libraries`
Cause:
- A targeted library has no mapping in `SP_LIBRARY_WATERMARKS`.

Fix:
- Add mapping for each targeted library:
- Example:
  `SP_LIBRARY_WATERMARKS=Documents=C:\wm\classified.png;Archive=C:\wm\archived.png`

### `InvalidAuthenticationToken` / token expired during long run
Cause:
- Expired access token while processing many files.

Fix:
- Use a build that includes token refresh retry logic.
- Confirm latest portable zip is deployed.

### `Processed=0 skipped=N` right after dry-run
Cause:
- Older build updated run-state after dry-run.

Fix:
- Use latest build where dry-run does not update state.
- If needed for retest, delete state file once:
  - `Remove-Item .\.watermark_state.json -ErrorAction SilentlyContinue`

### PowerShell `NativeCommandError` noise in ISE
Cause:
- Host formatting of stderr output, not necessarily job failure.

Fix:
- Trust the app summary line:
  - `Run successful. Processed=X skipped=Y failed=Z`
- Prefer standard PowerShell console over ISE when possible.

### Files skipped unexpectedly
Possible causes:
- Unsupported file extension.
- File created before last successful run timestamp.
- Library not selected by filter.

Fix:
- Run with `--dry-run --log-level DEBUG`.
- Check state file timestamp and extension support.

## Supported Extensions (Current)
- `.docx`, `.docm`
- `.xlsx`, `.xlsm`
- `.pptx`, `.pptm`
- `.pdf`

## Log Collection
Capture and share:
- Full command used.
- Start and end summary lines.
- Any `ERROR` lines.
- Current commit/zip deployed.
