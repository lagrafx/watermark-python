# Changelog

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
