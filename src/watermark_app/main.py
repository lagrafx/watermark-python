"""CLI entry point for SharePoint watermark automation."""

from __future__ import annotations

import argparse
import logging
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from watermark_app.config import AppConfig
from watermark_app.graph import GraphClient, GraphClientError
from watermark_app.state import load_state, save_state, should_process
from watermark_app.watermarking import apply_watermark, is_supported_extension

LOG = logging.getLogger(__name__)


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Apply a PNG watermark to new Word/Excel files in SharePoint libraries."
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="List/process files locally without uploading changes to SharePoint.",
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        help="Console log verbosity.",
    )
    return parser


def run(argv: list[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)
    logging.basicConfig(
        level=getattr(logging, args.log_level),
        format="%(asctime)s %(levelname)s %(message)s",
    )
    run_started = datetime.now(timezone.utc)
    config = AppConfig.from_env()
    state = load_state(config.state_file)
    LOG.info("Starting run (dry_run=%s)", args.dry_run)
    LOG.info("Authentication mode: %s", config.auth_mode)
    LOG.info("Last successful run: %s", state.last_successful_run_utc or "none")

    try:
        graph = GraphClient(config)
        site_id = graph.resolve_site_id()
        drives = graph.list_drives(site_id)
    except GraphClientError as exc:
        LOG.error("Failed to initialize Graph access: %s", exc)
        return 2

    library_filter = {name.lower() for name in config.library_names}
    if library_filter:
        drives = [d for d in drives if d.get("name", "").lower() in library_filter]

    processed = 0
    failed = 0
    skipped = 0

    with tempfile.TemporaryDirectory(prefix="watermark_") as tmp_dir:
        tmp_root = Path(tmp_dir)
        for drive in drives:
            drive_id = drive["id"]
            drive_name = drive.get("name", drive_id)
            watermark_path = config.library_watermark_paths.get(
                drive_name.lower(), config.watermark_image_path
            )
            LOG.info("Scanning library: %s", drive_name)
            try:
                items = graph.iter_files(drive_id)
            except GraphClientError as exc:
                LOG.error("Failed to list files in %s: %s", drive_name, exc)
                failed += 1
                continue

            for item in items:
                file_name = item.get("name", "")
                if not is_supported_extension(file_name):
                    skipped += 1
                    continue
                if not should_process(item.get("createdDateTime"), state.last_successful_run_utc):
                    skipped += 1
                    continue

                LOG.info("Processing %s", item.get("webUrl", file_name))
                item_id = item["id"]
                source_path = tmp_root / f"{item_id}_{file_name}"
                output_path = tmp_root / f"{item_id}_watermarked_{file_name}"
                try:
                    file_bytes = graph.download_file(drive_id, item_id)
                    source_path.write_bytes(file_bytes)
                    apply_watermark(source_path, output_path, watermark_path)
                    if not args.dry_run:
                        graph.upload_file(drive_id, item_id, output_path.read_bytes())
                    processed += 1
                except Exception as exc:  # noqa: BLE001
                    LOG.error("Failed file %s: %s", file_name, exc)
                    failed += 1

    if failed == 0:
        save_state(config.state_file, run_started)
        LOG.info("Run successful. Processed=%s skipped=%s failed=%s", processed, skipped, failed)
        return 0

    LOG.error("Run completed with errors. Processed=%s skipped=%s failed=%s", processed, skipped, failed)
    return 1


def main() -> None:
    sys.exit(run())


if __name__ == "__main__":
    main()
