from datetime import datetime, timezone

from watermark_app.state import load_state, save_state


def test_save_and_load_state(tmp_path):
    state_file = tmp_path / "state.json"
    ts = datetime(2026, 2, 7, 12, 0, tzinfo=timezone.utc)
    save_state(state_file, ts, {"abc", "def"})
    loaded = load_state(state_file)
    assert loaded.last_successful_run_utc == ts
    assert loaded.processed_item_ids == frozenset({"abc", "def"})
