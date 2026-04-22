from datetime import datetime, timezone
from pathlib import Path

from watermark_app import main as main_module
from watermark_app.state import RunState, should_process
from watermark_app.watermarking import is_supported_extension


def test_supported_extensions() -> None:
    assert is_supported_extension("a.docx")
    assert is_supported_extension("a.xlsm")
    assert is_supported_extension("a.pptx")
    assert is_supported_extension("a.pptm")
    assert is_supported_extension("a.pdf")
    assert not is_supported_extension("a.txt")


def test_should_process_with_no_state() -> None:
    assert should_process("id-1", "2026-02-07T00:00:00Z", None, frozenset())


def test_should_process_with_state() -> None:
    last_run = datetime(2026, 2, 7, 0, 0, tzinfo=timezone.utc)
    assert should_process("id-1", "2026-02-07T00:01:00Z", last_run, frozenset())
    assert not should_process("id-1", "2026-02-06T23:59:00Z", last_run, frozenset())
    assert not should_process("id-1", "2026-02-08T00:00:00Z", last_run, frozenset({"id-1"}))


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

        def iter_changed_files(self, _drive_id: str, _delta_link: str | None = None):
            return [], "delta-1"

    state_saved = {"called": False}

    monkeypatch.setattr(main_module.AppConfig, "from_env", lambda: DummyConfig())
    monkeypatch.setattr(main_module, "GraphClient", DummyGraphClient)
    monkeypatch.setattr(
        main_module,
        "load_state",
        lambda _path: RunState(None, frozenset(), {"drive-id": "delta-0"}),
    )
    monkeypatch.setattr(
        main_module,
        "save_state",
        lambda _path, _run_started, **_kwargs: state_saved.__setitem__("called", True),
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

        def iter_changed_files(self, _drive_id: str, _delta_link: str | None = None):
            return [], "delta-1"

    state_saved = {"called": False}

    monkeypatch.setattr(main_module.AppConfig, "from_env", lambda: DummyConfig())
    monkeypatch.setattr(main_module, "GraphClient", DummyGraphClient)
    monkeypatch.setattr(
        main_module,
        "load_state",
        lambda _path: RunState(None, frozenset(), {"drive-id": "delta-0"}),
    )
    monkeypatch.setattr(
        main_module,
        "save_state",
        lambda _path, _run_started, **_kwargs: state_saved.__setitem__("called", True),
    )

    rc = main_module.run(["--log-level", "INFO"])

    assert rc == 0
    assert state_saved["called"]


def test_run_processes_only_new_files_based_on_delta(monkeypatch, tmp_path: Path) -> None:
    watermark = tmp_path / "wm.png"
    watermark.write_bytes(b"not-used")

    class DummyConfig:
        auth_mode = "certificate"
        state_file = tmp_path / "state.json"
        library_names = ["WatermarkTesting"]
        library_watermark_paths = {"watermarktesting": watermark}

    scenario = {"pass": 1}
    uploads: list[str] = []
    state_holder = {"state": RunState(None, frozenset(), {})}

    class DummyGraphClient:
        def __init__(self, _config):  # noqa: ANN001
            pass

        def resolve_site_id(self) -> str:
            return "site-id"

        def list_drives(self, _site_id: str) -> list[dict]:
            return [{"id": "drive-id", "name": "WatermarkTesting"}]

        def iter_changed_files(self, _drive_id: str, delta_link: str | None = None):
            if scenario["pass"] == 1:
                assert delta_link is None
                return [
                    {"id": "f1", "name": "a.docx", "createdDateTime": "2026-02-07T00:00:00Z"},
                    {"id": "f2", "name": "b.xlsx", "createdDateTime": "2026-02-07T00:00:00Z"},
                ], "delta-1"
            assert delta_link == "delta-1"
            return [
                {"id": "f1", "name": "a.docx", "createdDateTime": "2026-02-07T00:00:00Z"},
                {"id": "f2", "name": "b.xlsx", "createdDateTime": "2026-02-07T00:00:00Z"},
                {"id": "f3", "name": "c.docx", "createdDateTime": "2026-02-07T00:00:00Z"}
            ], "delta-2"

        def download_file(self, _drive_id: str, _item_id: str) -> bytes:
            return b"fake"

        def upload_file(self, _drive_id: str, item_id: str, _data: bytes) -> None:
            uploads.append(item_id)

    def fake_save_state(_path, run_started, **kwargs):  # noqa: ANN001
        ids = kwargs.get("processed_item_ids") or set()
        links = kwargs.get("drive_delta_links") or {}
        state_holder["state"] = RunState(run_started, frozenset(ids), dict(links))

    monkeypatch.setattr(main_module.AppConfig, "from_env", lambda: DummyConfig())
    monkeypatch.setattr(main_module, "GraphClient", DummyGraphClient)
    monkeypatch.setattr(main_module, "load_state", lambda _path: state_holder["state"])
    monkeypatch.setattr(main_module, "save_state", fake_save_state)
    monkeypatch.setattr(main_module, "apply_watermark", lambda _s, out, _w: out.write_bytes(b"wm"))

    rc1 = main_module.run(["--log-level", "INFO"])
    assert rc1 == 0
    assert uploads == ["f1", "f2"]
    assert state_holder["state"].drive_delta_links == {"drive-id": "delta-1"}

    scenario["pass"] = 2
    rc2 = main_module.run(["--log-level", "INFO"])
    assert rc2 == 0
    assert uploads == ["f1", "f2", "f3"]
    assert state_holder["state"].drive_delta_links == {"drive-id": "delta-2"}
    assert state_holder["state"].processed_item_ids == frozenset({"f1", "f2", "f3"})
