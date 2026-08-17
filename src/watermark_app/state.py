"""Run-state persistence."""

from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path


@dataclass(frozen=True)
class RunState:
    last_successful_run_utc: datetime | None = None
    processed_item_ids: frozenset[str] = frozenset()
    drive_delta_links: dict[str, str] | None = None
    failed_items: dict[str, list[dict[str, str]]] | None = None


def _parse_iso_datetime(value: str | None) -> datetime | None:
    if not value:
        return None
    parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    return parsed.astimezone(timezone.utc)


def load_state(path: Path) -> RunState:
    if not path.exists():
        return RunState()
    data = json.loads(path.read_text(encoding="utf-8"))
    processed = data.get("processed_item_ids", [])
    if not isinstance(processed, list):
        processed = []
    raw_delta_links = data.get("drive_delta_links", {})
    if not isinstance(raw_delta_links, dict):
        raw_delta_links = {}
    delta_links: dict[str, str] = {}
    for drive_id, link in raw_delta_links.items():
        if isinstance(drive_id, str) and isinstance(link, str):
            delta_links[drive_id] = link
    raw_failed_items = data.get("failed_items", {})
    if not isinstance(raw_failed_items, dict):
        raw_failed_items = {}
    failed_items: dict[str, list[dict[str, str]]] = {}
    for drive_id, items in raw_failed_items.items():
        if not isinstance(drive_id, str) or not isinstance(items, list):
            continue
        parsed_items: list[dict[str, str]] = []
        for item in items:
            if not isinstance(item, dict):
                continue
            item_id = item.get("id")
            name = item.get("name", "")
            web_url = item.get("webUrl", "")
            if isinstance(item_id, str) and item_id:
                parsed_items.append(
                    {
                        "id": item_id,
                        "name": str(name),
                        "webUrl": str(web_url),
                    }
                )
        failed_items[drive_id] = parsed_items
    return RunState(
        last_successful_run_utc=_parse_iso_datetime(data.get("last_successful_run_utc")),
        processed_item_ids=frozenset(str(item_id) for item_id in processed),
        drive_delta_links=delta_links,
        failed_items=failed_items,
    )


def save_state(
    path: Path,
    run_started_utc: datetime,
    processed_item_ids: set[str] | None = None,
    drive_delta_links: dict[str, str] | None = None,
    failed_items: dict[str, list[dict[str, str]]] | None = None,
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload: dict[str, object] = {
        "last_successful_run_utc": run_started_utc.astimezone(timezone.utc).isoformat(),
        "processed_item_ids": sorted(processed_item_ids or set()),
    }
    if drive_delta_links is not None:
        payload["drive_delta_links"] = drive_delta_links
    if failed_items is not None:
        payload["failed_items"] = failed_items
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def should_process(
    item_id: str | None,
    created_datetime: str | None,
    last_successful_run_utc: datetime | None,
    processed_item_ids: frozenset[str],
) -> bool:
    if processed_item_ids:
        if not item_id:
            return False
        return item_id not in processed_item_ids
    if last_successful_run_utc is None:
        return True
    created = _parse_iso_datetime(created_datetime)
    if created is None:
        return False
    return created > last_successful_run_utc
