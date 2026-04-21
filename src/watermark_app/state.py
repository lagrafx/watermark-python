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
    return RunState(
        last_successful_run_utc=_parse_iso_datetime(data.get("last_successful_run_utc")),
        processed_item_ids=frozenset(str(item_id) for item_id in processed),
    )


def save_state(path: Path, run_started_utc: datetime, processed_item_ids: set[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "last_successful_run_utc": run_started_utc.astimezone(timezone.utc).isoformat(),
        "processed_item_ids": sorted(processed_item_ids),
    }
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
