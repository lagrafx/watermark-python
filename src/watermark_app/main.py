"""CLI entry point for SharePoint watermark automation."""

from __future__ import annotations

import argparse
import hashlib
import logging
import re
import sys
import tempfile
import uuid
from datetime import datetime, timezone
from pathlib import Path

from watermark_app.config import AppConfig
from watermark_app.graph import GraphClient, GraphClientError
from watermark_app.state import load_state, save_state
from watermark_app.watermarking import apply_watermark, is_supported_extension

LOG = logging.getLogger(__name__)


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Apply a PNG watermark to new Office/PDF files in SharePoint libraries."
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
    parser.add_argument(
        "--list-fields",
        action="store_true",
        help=(
            "List SharePoint metadata fields (internal + display names) for targeted "
            "libraries, then exit."
        ),
    )
    parser.add_argument(
        "--diagnose-libraries",
        action="store_true",
        help=(
            "Run read-only diagnostics for targeted libraries, including library identity, "
            "delta-state presence, watermark mapping, and required metadata fields."
        ),
    )
    parser.add_argument(
        "--write-probe",
        action="store_true",
        help=(
            "With --diagnose-libraries, create, update, and delete a tiny temporary "
            "probe file in each targeted library to verify write access."
        ),
    )
    parser.add_argument(
        "--first-file-only",
        action="store_true",
        help=(
            "Process only the first eligible file, then stop. This is intended for "
            "production troubleshooting and does not save Graph delta links."
        ),
    )
    parser.add_argument(
        "--save-diagnostics",
        nargs="?",
        const="diagnostics",
        default=None,
        metavar="DIR",
        help=(
            "Save per-file troubleshooting artifacts. For each attempted upload, "
            "the app writes the original download, locally watermarked output, and "
            "post-upload SharePoint download. If DIR is omitted, uses ./diagnostics."
        ),
    )
    return parser


def _run_write_probe(graph: GraphClient, drive_id: str, drive_name: str) -> bool:
    file_name = f"_watermark_app_write_probe_{uuid.uuid4().hex}.txt"
    created_item_id: str | None = None
    LOG.warning(
        "  write_probe=enabled; temporary file will be created, updated, and deleted: %s",
        file_name,
    )
    try:
        created = graph.create_root_file(
            drive_id,
            file_name,
            b"watermark app write probe - create\n",
        )
        created_item_id = created.get("id")
        if not created_item_id:
            LOG.error("  write_probe=create failed for %s: Graph returned no item id", drive_name)
            return False
        LOG.info("  write_probe=create ok item_id=%s", created_item_id)
    except GraphClientError as exc:
        LOG.error("  write_probe=create failed for %s: %s", drive_name, exc)
        return False

    try:
        graph.upload_file(
            drive_id,
            created_item_id,
            b"watermark app write probe - update\n",
        )
        LOG.info("  write_probe=update ok item_id=%s", created_item_id)
    except GraphClientError as exc:
        LOG.error("  write_probe=update failed for %s: %s", drive_name, exc)
        try:
            graph.delete_drive_item(drive_id, created_item_id)
            LOG.info("  write_probe=cleanup delete ok after update failure")
        except GraphClientError as cleanup_exc:
            LOG.error(
                "  write_probe=cleanup delete failed; remove manually if present: "
                "%s (%s)",
                file_name,
                cleanup_exc,
            )
        return False

    try:
        graph.delete_drive_item(drive_id, created_item_id)
        LOG.info("  write_probe=delete ok item_id=%s", created_item_id)
    except GraphClientError as exc:
        LOG.error(
            "  write_probe=delete failed for %s; remove manually if present: %s (%s)",
            drive_name,
            file_name,
            exc,
        )
        return False

    LOG.info("  write_probe=passed for %s", drive_name)
    return True


def _retry_items_for_drive(
    graph: GraphClient,
    drive_id: str,
    prior_failures: list[dict[str, str]],
) -> list[dict]:
    retry_items: list[dict] = []
    for failed_item in prior_failures:
        item_id = failed_item.get("id")
        if not item_id:
            continue
        try:
            item = graph.get_drive_item(drive_id, item_id)
        except GraphClientError as exc:
            LOG.warning(
                "Previously failed item is no longer readable; dropping retry item_id=%s "
                "name=%s error=%s",
                item_id,
                failed_item.get("name", ""),
                exc,
            )
            continue
        if "deleted" in item:
            LOG.warning(
                "Previously failed item appears deleted; dropping retry item_id=%s name=%s",
                item_id,
                failed_item.get("name", ""),
            )
            continue
        if "file" not in item:
            LOG.warning(
                "Previously failed item is no longer a file; dropping retry item_id=%s name=%s",
                item_id,
                failed_item.get("name", ""),
            )
            continue
        retry_items.append(item)
    return retry_items


def _merge_items(retry_items: list[dict], changed_items: list[dict]) -> list[dict]:
    merged: list[dict] = []
    seen: set[str] = set()
    for item in [*retry_items, *changed_items]:
        item_id = item.get("id")
        if not isinstance(item_id, str) or not item_id:
            merged.append(item)
            continue
        if item_id in seen:
            continue
        seen.add(item_id)
        merged.append(item)
    return merged


def _failed_item_record(item: dict, error: Exception) -> dict[str, str]:
    return {
        "id": str(item.get("id", "")),
        "name": str(item.get("name", "")),
        "webUrl": str(item.get("webUrl", "")),
        "error": str(error),
    }


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _safe_path_part(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", value.strip())
    return cleaned.strip("._") or "unnamed"


def _diagnostic_attempt_dir(
    diagnostics_root: Path | None,
    run_started: datetime,
    drive_name: str,
    item_id: str,
    file_name: str,
) -> Path | None:
    if diagnostics_root is None:
        return None
    timestamp = run_started.strftime("%Y%m%d-%H%M%S")
    path = (
        diagnostics_root
        / timestamp
        / _safe_path_part(drive_name)
        / f"{_safe_path_part(file_name)}_{_safe_path_part(item_id)}"
    )
    path.mkdir(parents=True, exist_ok=True)
    return path


def _write_diagnostic_bytes(
    attempt_dir: Path | None,
    label: str,
    file_name: str,
    data: bytes,
) -> None:
    if attempt_dir is None:
        return
    path = attempt_dir / f"{label}_{_safe_path_part(file_name)}"
    path.write_bytes(data)
    LOG.info("diagnostic_saved=%s bytes=%s", path, len(data))


def _verify_uploaded_content(
    graph: GraphClient,
    drive_id: str,
    item_id: str,
    file_name: str,
    expected_bytes: bytes,
    diagnostics_dir: Path | None = None,
) -> None:
    expected_hash = _sha256(expected_bytes)
    verify_bytes = graph.download_file(drive_id, item_id)
    _write_diagnostic_bytes(diagnostics_dir, "03_sharepoint_after_upload", file_name, verify_bytes)
    actual_hash = _sha256(verify_bytes)
    if actual_hash != expected_hash:
        LOG.error(
            "post_upload_verify=failed file=%s expected_sha256=%s actual_sha256=%s "
            "expected_bytes=%s actual_bytes=%s",
            file_name,
            expected_hash,
            actual_hash,
            len(expected_bytes),
            len(verify_bytes),
        )
        raise RuntimeError(
            "post-upload verification failed; SharePoint returned different bytes "
            "than the app uploaded"
        )
    LOG.info(
        "post_upload_verify=passed file=%s sha256=%s bytes=%s",
        file_name,
        actual_hash,
        len(verify_bytes),
    )


def _log_library_diagnostics(
    graph: GraphClient,
    drives: list[dict],
    config: AppConfig,
    drive_delta_links: dict[str, str] | None,
    write_probe: bool = False,
) -> int:
    failures = 0
    delta_links = drive_delta_links or {}
    for drive in drives:
        drive_id = drive["id"]
        drive_name = drive.get("name", drive_id)
        drive_key = drive_name.lower()
        watermark_path = config.library_watermark_paths.get(drive_key)

        LOG.info("Diagnostic library: %s", drive_name)
        LOG.info("The account or ID accessing this library is '%s'", graph.access_identity)
        LOG.info("  drive_id=%s", drive_id)
        LOG.info("  drive_webUrl=%s", drive.get("webUrl", "(not returned)"))
        LOG.info("  delta_state=%s", "present" if drive_id in delta_links else "missing")
        if watermark_path:
            LOG.info("  watermark_mapping=present path=%s", watermark_path)
        else:
            LOG.warning("  watermark_mapping=missing for configured target library")
            failures += 1

        try:
            details = graph.get_library_details(drive_id)
            list_facet = details.get("list", {})
            sharepoint_ids = details.get("sharepointIds", {})
            LOG.info("  list_id=%s", details.get("id", "(not returned)"))
            LOG.info("  list_displayName=%s", details.get("displayName", "(not returned)"))
            LOG.info("  list_name=%s", details.get("name", "(not returned)"))
            LOG.info("  list_webUrl=%s", details.get("webUrl", "(not returned)"))
            LOG.info("  list_template=%s", list_facet.get("template", "(not returned)"))
            LOG.info(
                "  list_contentTypesEnabled=%s",
                list_facet.get("contentTypesEnabled", "(not returned)"),
            )
            LOG.info("  list_hidden=%s", list_facet.get("hidden", "(not returned)"))
            LOG.info("  sharepoint_listId=%s", sharepoint_ids.get("listId", "(not returned)"))
        except GraphClientError as exc:
            LOG.error("  library_details=failed: %s", exc)
            failures += 1

        try:
            fields = graph.list_library_fields(drive_id)
        except GraphClientError as exc:
            LOG.error("  field_read=failed: %s", exc)
            failures += 1
            continue

        required_editable = [
            field
            for field in fields
            if field.get("required")
            and not field.get("readOnly")
            and not field.get("hidden")
        ]
        hidden_count = sum(1 for field in fields if field.get("hidden"))
        read_only_count = sum(1 for field in fields if field.get("readOnly"))
        LOG.info(
            "  field_read=ok total=%s required_editable=%s readOnly=%s hidden=%s",
            len(fields),
            len(required_editable),
            read_only_count,
            hidden_count,
        )
        if required_editable:
            LOG.warning(
                "  required editable metadata fields may block uploads if SharePoint "
                "requires values during file replacement:"
            )
            for field in sorted(required_editable, key=lambda f: (f.get("name") or "").lower()):
                LOG.warning(
                    "    field=%s displayName=%s type=%s",
                    field.get("name", ""),
                    field.get("displayName", ""),
                    field.get("columnGroup", field.get("type", "(not returned)")),
                )
        else:
            LOG.info("  required editable metadata fields: none detected")

        if write_probe and not _run_write_probe(graph, drive_id, drive_name):
            failures += 1

    if failures:
        LOG.error("Library diagnostics completed with %s warning/error condition(s).", failures)
        return 1
    LOG.info("Library diagnostics completed without detected configuration errors.")
    return 0


def run(argv: list[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)
    logging.basicConfig(
        level=getattr(logging, args.log_level),
        format="%(asctime)s %(levelname)s %(message)s",
    )
    run_started = datetime.now(timezone.utc)
    if args.write_probe and not args.diagnose_libraries:
        LOG.error("--write-probe must be used with --diagnose-libraries.")
        return 2
    if args.first_file_only and args.dry_run:
        LOG.warning(
            "--first-file-only is running with --dry-run; no uploads or state updates occur."
        )

    config = AppConfig.from_env()
    state = load_state(config.state_file)
    diagnostics_root = (
        Path(args.save_diagnostics).expanduser().resolve() if args.save_diagnostics else None
    )
    LOG.info("Starting run (dry_run=%s)", args.dry_run)
    LOG.info("Authentication mode: %s", config.auth_mode)
    LOG.info("Target site: %s%s", config.site_hostname, config.site_path)
    LOG.info("State file: %s (exists=%s)", config.state_file, config.state_file.exists())
    LOG.info(
        "Configured library filter: %s",
        ", ".join(config.library_names) if config.library_names else "(all libraries)",
    )
    LOG.info(
        "Configured watermark mappings: %s",
        ", ".join(sorted(config.library_watermark_paths.keys())),
    )
    LOG.info("Last successful run: %s", state.last_successful_run_utc or "none")
    LOG.info("Loaded processed item IDs: %s", len(state.processed_item_ids))
    LOG.info("Loaded drive delta links: %s", len(state.drive_delta_links or {}))
    LOG.info(
        "Loaded failed retry items: %s",
        sum(len(items) for items in (state.failed_items or {}).values()),
    )
    if diagnostics_root:
        LOG.info("Diagnostics output: %s", diagnostics_root)

    try:
        graph = GraphClient(config)
        site_id = graph.resolve_site_id()
        drives = graph.list_drives(site_id)
    except GraphClientError as exc:
        LOG.error("Failed to initialize Graph access: %s", exc)
        return 2

    available_drive_names = [d.get("name", d.get("id", "")) for d in drives]
    LOG.info(
        "Discovered SharePoint libraries/drives: %s",
        ", ".join(available_drive_names) if available_drive_names else "(none)",
    )

    library_filter = {name.lower() for name in config.library_names}
    if library_filter:
        drives = [d for d in drives if d.get("name", "").lower() in library_filter]

    if not drives:
        LOG.error(
            "No SharePoint libraries matched SP_LIBRARY_NAMES. Configured=%s Available=%s",
            ", ".join(config.library_names) if config.library_names else "(all libraries)",
            ", ".join(available_drive_names) if available_drive_names else "(none)",
        )
        return 2

    if args.list_fields:
        if not drives:
            LOG.warning("No libraries matched the current SP_LIBRARY_NAMES filter.")
            return 0
        for drive in drives:
            drive_id = drive["id"]
            drive_name = drive.get("name", drive_id)
            LOG.info("Library: %s", drive_name)
            try:
                fields = graph.list_library_fields(drive_id)
            except GraphClientError as exc:
                LOG.error("Failed to list fields for %s: %s", drive_name, exc)
                continue
            for field in sorted(fields, key=lambda f: (f.get("name") or "").lower()):
                internal_name = field.get("name", "")
                display_name = field.get("displayName", "")
                read_only = field.get("readOnly", False)
                hidden = field.get("hidden", False)
                LOG.info(
                    "  field=%s displayName=%s readOnly=%s hidden=%s",
                    internal_name,
                    display_name,
                    read_only,
                    hidden,
                )
        LOG.info("Field listing complete.")
        return 0

    if args.diagnose_libraries:
        return _log_library_diagnostics(
            graph=graph,
            drives=drives,
            config=config,
            drive_delta_links=state.drive_delta_links,
            write_probe=args.write_probe,
        )

    missing_mappings = [
        d.get("name", d["id"])
        for d in drives
        if d.get("name", "").lower() not in config.library_watermark_paths
    ]
    if missing_mappings:
        LOG.error(
            "Missing SP_LIBRARY_WATERMARKS entries for targeted libraries: %s",
            ", ".join(missing_mappings),
        )
        return 2

    processed = 0
    failed = 0
    skipped = 0
    stop_after_first_attempt = False

    with tempfile.TemporaryDirectory(prefix="watermark_") as tmp_dir:
        tmp_root = Path(tmp_dir)
        new_delta_links = dict(state.drive_delta_links or {})
        processed_item_ids = set(state.processed_item_ids)
        failed_items_by_drive: dict[str, list[dict[str, str]]] = {}
        for drive in drives:
            if stop_after_first_attempt:
                break
            drive_id = drive["id"]
            drive_name = drive.get("name", drive_id)
            watermark_path = config.library_watermark_paths[drive_name.lower()]
            prior_failed_items = (state.failed_items or {}).get(drive_id, [])
            retry_items = _retry_items_for_drive(graph, drive_id, prior_failed_items)
            if prior_failed_items:
                LOG.info(
                    "Retry queue for %s: loaded=%s readable_files=%s",
                    drive_name,
                    len(prior_failed_items),
                    len(retry_items),
                )
            LOG.info(
                "Scanning library: %s (drive_id=%s, watermark=%s)",
                drive_name,
                drive_id,
                watermark_path,
            )
            try:
                prior_delta_link = new_delta_links.get(drive_id)
                LOG.info(
                    "Delta mode for %s: %s",
                    drive_name,
                    "incremental (stored delta link found)"
                    if prior_delta_link
                    else "initial/full baseline (no stored delta link)",
                )
                items, latest_delta_link = graph.iter_changed_files(drive_id, prior_delta_link)
                new_delta_links[drive_id] = latest_delta_link
                LOG.info(
                    "Graph returned %s changed file item(s) for %s; latest delta link captured=%s",
                    len(items),
                    drive_name,
                    bool(latest_delta_link),
                )
            except GraphClientError as exc:
                LOG.error("Failed to list files in %s: %s", drive_name, exc)
                failed += 1
                if prior_failed_items:
                    failed_items_by_drive[drive_id] = prior_failed_items
                continue

            items = _merge_items(retry_items, items)
            if retry_items:
                LOG.info(
                    "Processing %s total candidate item(s) for %s after merging "
                    "%s retry item(s) with Graph delta results",
                    len(items),
                    drive_name,
                    len(retry_items),
                )

            library_processed = 0
            library_unsupported = 0
            library_already_processed = 0
            library_failed = 0
            library_failed_items: list[dict[str, str]] = []
            for item_index, item in enumerate(items, start=1):
                if item_index == 1 or item_index % 250 == 0 or item_index == len(items):
                    LOG.info(
                        "Progress for %s: evaluating item %s/%s "
                        "(processed=%s skipped_already_processed=%s "
                        "skipped_unsupported=%s failed=%s)",
                        drive_name,
                        item_index,
                        len(items),
                        library_processed,
                        library_already_processed,
                        library_unsupported,
                        library_failed,
                    )
                file_name = item.get("name", "")
                if not is_supported_extension(file_name):
                    LOG.debug("Skipping unsupported file type: %s", file_name)
                    library_unsupported += 1
                    skipped += 1
                    continue
                item_id = item.get("id")
                if item_id and item_id in processed_item_ids:
                    LOG.debug("Skipping already processed item: %s (%s)", file_name, item_id)
                    library_already_processed += 1
                    skipped += 1
                    continue

                LOG.info("Processing %s", item.get("webUrl", file_name))
                item_id = item["id"]
                source_path = tmp_root / f"{item_id}_{file_name}"
                output_path = tmp_root / f"{item_id}_watermarked_{file_name}"
                attempt_diagnostics_dir = _diagnostic_attempt_dir(
                    diagnostics_root,
                    run_started,
                    drive_name,
                    item_id,
                    file_name,
                )
                try:
                    file_bytes = graph.download_file(drive_id, item_id)
                    LOG.debug("Downloaded %s bytes for %s", len(file_bytes), file_name)
                    _write_diagnostic_bytes(
                        attempt_diagnostics_dir,
                        "01_original_download",
                        file_name,
                        file_bytes,
                    )
                    source_path.write_bytes(file_bytes)
                    apply_watermark(source_path, output_path, watermark_path)
                    output_bytes = output_path.read_bytes()
                    _write_diagnostic_bytes(
                        attempt_diagnostics_dir,
                        "02_local_watermarked",
                        file_name,
                        output_bytes,
                    )
                    LOG.debug(
                        "Watermarked %s: source_bytes=%s output_bytes=%s",
                        file_name,
                        len(file_bytes),
                        len(output_bytes),
                    )
                    if not args.dry_run:
                        graph.upload_file(drive_id, item_id, output_bytes)
                        LOG.info("Uploaded watermarked file: %s", file_name)
                        _verify_uploaded_content(
                            graph,
                            drive_id,
                            item_id,
                            file_name,
                            output_bytes,
                            attempt_diagnostics_dir,
                        )
                    else:
                        LOG.info("Dry run: would upload watermarked file: %s", file_name)
                    processed += 1
                    library_processed += 1
                    processed_item_ids.add(item_id)
                except Exception as exc:  # noqa: BLE001
                    LOG.error("Failed file %s: %s", file_name, exc)
                    library_failed_items.append(_failed_item_record(item, exc))
                    library_failed += 1
                    failed += 1
                if args.first_file_only:
                    LOG.info(
                        "first_file_only=stopping after first eligible file attempt: %s",
                        file_name,
                    )
                    stop_after_first_attempt = True
                    break
            if library_failed_items:
                failed_items_by_drive[drive_id] = library_failed_items
                LOG.error("Failed item report for %s:", drive_name)
                for failed_item in library_failed_items:
                    LOG.error(
                        "  failed item_id=%s name=%s url=%s error=%s",
                        failed_item.get("id", ""),
                        failed_item.get("name", ""),
                        failed_item.get("webUrl", ""),
                        failed_item.get("error", ""),
                    )
            LOG.info(
                "Library summary for %s: changed=%s processed=%s skipped_already_processed=%s "
                "skipped_unsupported=%s failed=%s",
                drive_name,
                len(items),
                library_processed,
                library_already_processed,
                library_unsupported,
                library_failed,
            )

    if args.dry_run:
        LOG.info("Dry run complete; state file not updated.")
    else:
        LOG.info(
            "Saving state to %s (processed_item_ids=%s, drive_delta_links=%s, "
            "failed_retry_items=%s)",
            config.state_file,
            len(processed_item_ids),
            0 if args.first_file_only else len(new_delta_links),
            sum(len(items) for items in failed_items_by_drive.values()),
        )
        save_state(
            config.state_file,
            run_started,
            processed_item_ids=processed_item_ids,
            drive_delta_links=state.drive_delta_links if args.first_file_only else new_delta_links,
            failed_items=failed_items_by_drive,
        )
        if args.first_file_only:
            LOG.info(
                "first_file_only=delta links were not advanced, so unprocessed files remain "
                "eligible for normal future runs."
            )
        LOG.info(
            "State saved: %s (exists=%s, bytes=%s)",
            config.state_file,
            config.state_file.exists(),
            config.state_file.stat().st_size if config.state_file.exists() else 0,
        )

    if failed == 0:
        LOG.info("Run successful. Processed=%s skipped=%s failed=%s", processed, skipped, failed)
        return 0

    LOG.error(
        "Run completed with errors. Successful files were checkpointed; failed files "
        "will be retried on the next run. Processed=%s skipped=%s failed=%s",
        processed,
        skipped,
        failed,
    )
    return 1


def main() -> None:
    sys.exit(run())


if __name__ == "__main__":
    main()
