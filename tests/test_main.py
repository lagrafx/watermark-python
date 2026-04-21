from datetime import datetime, timezone
from pathlib import Path

from watermark_app import main as main_module
from watermark_app.state import RunState, should_process
from watermark_app.watermarking import is_supported_extension


def test_supported_extensions() -> None:
    assert is_supported_extension("a.docx")
    assert is_supported_extension("a.xlsm")
    assert not is_supported_extension("a.pdf")


def test_should_process_with_no_state() -> None:
    assert should_process("2026-02-07T00:00:00Z", None)


def test_should_process_with_state() -> None:
    last_run = datetime(2026, 2, 7, 0, 0, tzinfo=timezone.utc)
    assert should_process("2026-02-07T00:01:00Z", last_run)
    assert not should_process("2026-02-06T23:59:00Z", last_run)


def test_run_dry_run_does_not_save_state(monkeypatch, tmp_path: Path) -> None:
    watermark = tmp_path / "wm.png"
    watermark.write_bytes(b"not-used")

    class DummyConfig:
        auth_mode = "certificate"
        state_file = tmp_path / "state.json"
        library_names = ["WatermarkTesting"]
        library_watermark_paths = {"watermarktesting": watermark}

    class DummyGraphClient:
        def __init__(self, _config):  # noqa: ANN001
            pass

        def resolve_site_id(self) -> str:
            return "site-id"

        def list_drives(self, _site_id: str) -> list[dict]:
            return [{"id": "drive-id", "name": "WatermarkTesting"}]

        def iter_files(self, _drive_id: str) -> list[dict]:
            return []

    state_saved = {"called": False}

    monkeypatch.setattr(main_module.AppConfig, "from_env", lambda: DummyConfig())
    monkeypatch.setattr(main_module, "GraphClient", DummyGraphClient)
    monkeypatch.setattr(main_module, "load_state", lambda _path: RunState(None))
    monkeypatch.setattr(
        main_module,
        "save_state",
        lambda _path, _run_started: state_saved.__setitem__("called", True),
    )

    rc = main_module.run(["--dry-run", "--log-level", "INFO"])

    assert rc == 0
    assert not state_saved["called"]


def test_run_non_dry_run_saves_state(monkeypatch, tmp_path: Path) -> None:
    watermark = tmp_path / "wm.png"
    watermark.write_bytes(b"not-used")

    class DummyConfig:
        auth_mode = "certificate"
        state_file = tmp_path / "state.json"
        library_names = ["WatermarkTesting"]
        library_watermark_paths = {"watermarktesting": watermark}

    class DummyGraphClient:
        def __init__(self, _config):  # noqa: ANN001
            pass

        def resolve_site_id(self) -> str:
            return "site-id"

        def list_drives(self, _site_id: str) -> list[dict]:
            return [{"id": "drive-id", "name": "WatermarkTesting"}]

        def iter_files(self, _drive_id: str) -> list[dict]:
            return []

    state_saved = {"called": False}

    monkeypatch.setattr(main_module.AppConfig, "from_env", lambda: DummyConfig())
    monkeypatch.setattr(main_module, "GraphClient", DummyGraphClient)
    monkeypatch.setattr(main_module, "load_state", lambda _path: RunState(None))
    monkeypatch.setattr(
        main_module,
        "save_state",
        lambda _path, _run_started: state_saved.__setitem__("called", True),
    )

    rc = main_module.run(["--log-level", "INFO"])

    assert rc == 0
    assert state_saved["called"]
