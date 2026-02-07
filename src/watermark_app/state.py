"""Run-state persistence."""

from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path


@dataclass(frozen=True)
class RunState:
    last_successful_run_utc: datetime | None = None


def _parse_iso_datetime(value: str | None) -> datetime | None:
    if not value:
        return None
    parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    return parsed.astimezone(timezone.utc)


def load_state(path: Path) -> RunState:
    if not path.exists():
        return RunState()
    data = json.loads(path.read_text(encoding="utf-8"))
    return RunState(last_successful_run_utc=_parse_iso_datetime(data.get("last_successful_run_utc")))


def save_state(path: Path, run_started_utc: datetime) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {"last_successful_run_utc": run_started_utc.astimezone(timezone.utc).isoformat()}
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def should_process(created_datetime: str | None, last_successful_run_utc: datetime | None) -> bool:
    if last_successful_run_utc is None:
        return True
    created = _parse_iso_datetime(created_datetime)
    if created is None:
        return False
    return created > last_successful_run_utc
