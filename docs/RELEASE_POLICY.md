# Release Policy

## Goal
Create an automatic rollback point for every push to `main`.

## How it works
- Workflow: `.github/workflows/release-on-main.yml`
- Trigger: every push to `main`
- Output:
  - creates a unique git tag
  - creates a GitHub Release from that tag

Tag format:
- `auto-vYYYYMMDD-HHMMSS-<shortsha>`

Example:
- `auto-v20260514-184512-a1b2c3d`

## Rollback options
1. In GitHub, open Releases and pick the desired release tag.
2. Deploy artifacts/code from that tag.
3. Or locally:
   - `git fetch --tags`
   - `git checkout <tag>`

## Notes
- Releases are created automatically and are marked non-latest to avoid changing manual release strategy.
- For formal/manual versioning (for example `v1.2.0`), continue creating manual tags/releases when desired.
