# Changelog

## 2026-08-25

### Changed
- Replaced the fragile Word DrawingML anchor approach with a first-page-only VML
  watermark pattern to avoid corrupted documents and repeated page watermarks.
- Office post-upload verification now tolerates SharePoint package rewrites only
  when the app-added watermark media survives in the returned Office package.

### Fixed
- Word repair cleanup now recognizes any prior `watermark-python|...` tag,
  including the rejected `behind-text-v2` build.
- Excel files no longer fail solely because SharePoint rewrites `.xlsx` package
  bytes after upload.

## 2026-08-24

### Added
- New `--repair-watermarks` mode to intentionally reprocess supported files that
  were already checkpointed in state.
- New `--file-extension` filter for one-file-per-type production testing.
- Tests for repair mode, file-extension filtering, Word behind-text placement,
  PowerPoint layer ordering, and semi-transparent Excel watermark assets.

### Changed
- Word watermarks now use behind-text anchored placement instead of normal inline
  header pictures.
- Word repair now removes tagged watermarks and only removes legacy header
  watermark images when they match both the prior approximate 6-inch size and
  configured watermark PNG bytes.
- PowerPoint watermarks are moved behind existing slide shapes.
- PDF watermarks are drawn behind existing page content.
- Watermark PNGs are made semi-transparent at runtime, so existing PNG assets can
  be reused.

## 2026-08-17

### Added
- New `--save-diagnostics [DIR]` troubleshooting flag.
- Per-file diagnostic captures now include:
  - `01_original_download_<file>`
  - `02_local_watermarked_<file>`
  - `03_sharepoint_after_upload_<file>`
- Tests covering diagnostic artifact output.

### Fixed
- Word watermarking now also writes to active first-page and even-page headers,
  not only the primary section header.

## 2026-08-10

### Added
- More run diagnostics at `INFO` level:
  - target site and state-file path
  - configured library filter and watermark mappings
  - discovered SharePoint library/drive names
  - selected drive ID and watermark path
  - whether each library uses stored delta state or initial/full baseline
  - per-library changed/processed/skipped/failed summary
  - state save path, processed ID count, delta-link count, and state-file size
- Graph delta pagination diagnostics:
  - page number
  - raw item count
  - file/folder/deleted/other counts per page
  - whether each page had a next link or final delta link
  - final page/raw/file/folder/deleted/other totals
- Periodic per-library progress logging every 250 changed file items.

### Fixed
- Runs now fail with an explicit error if `SP_LIBRARY_NAMES` matches no SharePoint
  libraries/drives instead of exiting successfully with no work.
- State output now always includes `processed_item_ids`, even when empty, making
  state-file inspection less ambiguous.

## 2026-05-14

### Added
- New CLI mode:
  - `--list-fields` prints SharePoint library metadata fields for targeted libraries.
- Field details in output:
  - `field` (internal name)
  - `displayName` (friendly label)
  - `readOnly`
  - `hidden`

### Updated
- Docs updated with field-discovery usage:
  - `README.md`
  - `docs/OPERATIONS.md`
- Portable build now auto-generates release mapping artifacts:
  - `PORTABLE_RELEASE_MANIFEST.json`
  - `PORTABLE_RELEASE_NOTES.md`

## 2026-04-21

### Added
- PowerPoint watermark support:
  - `.pptx`
  - `.pptm`
- PDF watermark support:
  - `.pdf`
- New dependencies:
  - `python-pptx`
  - `pypdf`
  - `reportlab`

### Fixed
- Graph token refresh retry when token expires during long runs.
- Dry-run behavior:
  - `--dry-run` no longer updates run-state.
- Improved visibility for skipped unsupported file types.

### Notes
- Portable deployment artifact updated in `watermark-python-portable`.
- Existing `.env` and state files can be retained across runtime updates.
