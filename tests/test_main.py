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
        site_hostname = "contoso.sharepoint.com"
        site_path = "/sites/Test"

    class DummyGraphClient:
        access_identity = "Watermark - Python"

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
        site_hostname = "contoso.sharepoint.com"
        site_path = "/sites/Test"

    class DummyGraphClient:
        access_identity = "Watermark - Python"

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


def test_run_fails_when_library_filter_matches_no_drives(monkeypatch, tmp_path: Path) -> None:
    watermark = tmp_path / "wm.png"
    watermark.write_bytes(b"not-used")

    class DummyConfig:
        auth_mode = "certificate"
        state_file = tmp_path / "state.json"
        library_names = ["Archive"]
        library_watermark_paths = {"archive": watermark}
        site_hostname = "contoso.sharepoint.com"
        site_path = "/sites/Test"

    class DummyGraphClient:
        access_identity = "Watermark - Python"

        def __init__(self, _config):  # noqa: ANN001
            pass

        def resolve_site_id(self) -> str:
            return "site-id"

        def list_drives(self, _site_id: str) -> list[dict]:
            return [{"id": "drive-id", "name": "WatermarkTesting"}]

    state_saved = {"called": False}

    monkeypatch.setattr(main_module.AppConfig, "from_env", lambda: DummyConfig())
    monkeypatch.setattr(main_module, "GraphClient", DummyGraphClient)
    monkeypatch.setattr(main_module, "load_state", lambda _path: RunState(None, frozenset(), {}))
    monkeypatch.setattr(
        main_module,
        "save_state",
        lambda _path, _run_started, **_kwargs: state_saved.__setitem__("called", True),
    )

    rc = main_module.run(["--log-level", "INFO"])

    assert rc == 2
    assert not state_saved["called"]


def test_run_processes_only_new_files_based_on_delta(monkeypatch, tmp_path: Path) -> None:
    watermark = tmp_path / "wm.png"
    watermark.write_bytes(b"not-used")

    class DummyConfig:
        auth_mode = "certificate"
        state_file = tmp_path / "state.json"
        library_names = ["WatermarkTesting"]
        library_watermark_paths = {"watermarktesting": watermark}
        site_hostname = "contoso.sharepoint.com"
        site_path = "/sites/Test"

    scenario = {"pass": 1}
    uploads: list[str] = []
    stored_bytes: dict[str, bytes] = {}
    state_holder = {"state": RunState(None, frozenset(), {})}

    class DummyGraphClient:
        access_identity = "Watermark - Python"

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

        def download_file(self, _drive_id: str, item_id: str) -> bytes:
            return stored_bytes.get(item_id, b"fake")

        def upload_file(self, _drive_id: str, item_id: str, _data: bytes) -> None:
            uploads.append(item_id)
            stored_bytes[item_id] = _data

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


def test_run_repair_watermarks_reprocesses_checkpointed_files_without_advancing_delta(
    monkeypatch,
    tmp_path: Path,
) -> None:
    watermark = tmp_path / "wm.png"
    watermark.write_bytes(b"not-used")
    uploads: list[str] = []
    stored_bytes = {"f1": b"old"}
    state_holder = {"state": RunState(None, frozenset({"f1"}), {"drive-id": "delta-old"})}

    class DummyConfig:
        auth_mode = "certificate"
        state_file = tmp_path / "state.json"
        library_names = ["Archive"]
        library_watermark_paths = {"archive": watermark}
        site_hostname = "contoso.sharepoint.com"
        site_path = "/sites/Test"

    class DummyGraphClient:
        access_identity = "Watermark - Python"

        def __init__(self, _config):  # noqa: ANN001
            pass

        def resolve_site_id(self) -> str:
            return "site-id"

        def list_drives(self, _site_id: str) -> list[dict]:
            return [{"id": "drive-id", "name": "Archive"}]

        def iter_files(self, _drive_id: str) -> list[dict]:
            return [
                {"id": "f1", "name": "already.docx", "webUrl": "https://sp/already.docx"},
                {"id": "txt-1", "name": "notes.txt", "webUrl": "https://sp/notes.txt"},
            ]

        def download_file(self, _drive_id: str, item_id: str) -> bytes:
            return stored_bytes[item_id]

        def upload_file(self, _drive_id: str, item_id: str, data: bytes) -> None:
            uploads.append(item_id)
            stored_bytes[item_id] = data

    def fake_save_state(_path, run_started, **kwargs):  # noqa: ANN001
        state_holder["state"] = RunState(
            run_started,
            frozenset(kwargs.get("processed_item_ids") or set()),
            dict(kwargs.get("drive_delta_links") or {}),
            dict(kwargs.get("failed_items") or {}),
        )

    monkeypatch.setattr(main_module.AppConfig, "from_env", lambda: DummyConfig())
    monkeypatch.setattr(main_module, "GraphClient", DummyGraphClient)
    monkeypatch.setattr(main_module, "load_state", lambda _path: state_holder["state"])
    monkeypatch.setattr(main_module, "save_state", fake_save_state)
    monkeypatch.setattr(main_module, "apply_watermark", lambda _s, out, _w: out.write_bytes(b"wm"))

    rc = main_module.run(["--repair-watermarks", "--log-level", "INFO"])

    assert rc == 0
    assert uploads == ["f1"]
    assert state_holder["state"].processed_item_ids == frozenset({"f1"})
    assert state_holder["state"].drive_delta_links == {"drive-id": "delta-old"}


def test_run_file_extension_filter_limits_processed_files(monkeypatch, tmp_path: Path) -> None:
    watermark = tmp_path / "wm.png"
    watermark.write_bytes(b"not-used")
    uploads: list[str] = []
    stored_bytes: dict[str, bytes] = {}
    state_holder = {"state": RunState(None, frozenset(), {})}

    class DummyConfig:
        auth_mode = "certificate"
        state_file = tmp_path / "state.json"
        library_names = ["Archive"]
        library_watermark_paths = {"archive": watermark}
        site_hostname = "contoso.sharepoint.com"
        site_path = "/sites/Test"

    class DummyGraphClient:
        access_identity = "Watermark - Python"

        def __init__(self, _config):  # noqa: ANN001
            pass

        def resolve_site_id(self) -> str:
            return "site-id"

        def list_drives(self, _site_id: str) -> list[dict]:
            return [{"id": "drive-id", "name": "Archive"}]

        def iter_changed_files(self, _drive_id: str, _delta_link: str | None = None):
            return [
                {"id": "doc-1", "name": "first.docx", "webUrl": "https://sp/first.docx"},
                {"id": "pdf-1", "name": "second.pdf", "webUrl": "https://sp/second.pdf"},
            ], "delta-1"

        def download_file(self, _drive_id: str, item_id: str) -> bytes:
            return stored_bytes.get(item_id, b"fake")

        def upload_file(self, _drive_id: str, item_id: str, data: bytes) -> None:
            uploads.append(item_id)
            stored_bytes[item_id] = data

    def fake_save_state(_path, run_started, **kwargs):  # noqa: ANN001
        state_holder["state"] = RunState(
            run_started,
            frozenset(kwargs.get("processed_item_ids") or set()),
            dict(kwargs.get("drive_delta_links") or {}),
            dict(kwargs.get("failed_items") or {}),
        )

    monkeypatch.setattr(main_module.AppConfig, "from_env", lambda: DummyConfig())
    monkeypatch.setattr(main_module, "GraphClient", DummyGraphClient)
    monkeypatch.setattr(main_module, "load_state", lambda _path: state_holder["state"])
    monkeypatch.setattr(main_module, "save_state", fake_save_state)
    monkeypatch.setattr(main_module, "apply_watermark", lambda _s, out, _w: out.write_bytes(b"wm"))

    rc = main_module.run(["--file-extension", "pdf", "--log-level", "INFO"])

    assert rc == 0
    assert uploads == ["pdf-1"]
    assert state_holder["state"].processed_item_ids == frozenset({"pdf-1"})


def test_run_file_name_filter_limits_processed_files(monkeypatch, tmp_path: Path) -> None:
    watermark = tmp_path / "wm.png"
    watermark.write_bytes(b"not-used")
    uploads: list[str] = []
    stored_bytes: dict[str, bytes] = {}
    state_holder = {"state": RunState(None, frozenset(), {})}

    class DummyConfig:
        auth_mode = "certificate"
        state_file = tmp_path / "state.json"
        library_names = ["WatermarkTesting"]
        library_watermark_paths = {"watermarktesting": watermark}
        site_hostname = "contoso.sharepoint.com"
        site_path = "/sites/Test"

    class DummyGraphClient:
        access_identity = "Watermark - Python"

        def __init__(self, _config):  # noqa: ANN001
            pass

        def resolve_site_id(self) -> str:
            return "site-id"

        def list_drives(self, _site_id: str) -> list[dict]:
            return [{"id": "drive-id", "name": "WatermarkTesting"}]

        def iter_changed_files(self, _drive_id: str, _delta_link: str | None = None):
            return [
                {"id": "doc-1", "name": "wrong.docx", "webUrl": "https://sp/wrong.docx"},
                {"id": "doc-2", "name": "Target Test.DOCX", "webUrl": "https://sp/target.docx"},
            ], "delta-1"

        def download_file(self, _drive_id: str, item_id: str) -> bytes:
            return stored_bytes.get(item_id, b"fake")

        def upload_file(self, _drive_id: str, item_id: str, data: bytes) -> None:
            uploads.append(item_id)
            stored_bytes[item_id] = data

    def fake_save_state(_path, run_started, **kwargs):  # noqa: ANN001
        state_holder["state"] = RunState(
            run_started,
            frozenset(kwargs.get("processed_item_ids") or set()),
            dict(kwargs.get("drive_delta_links") or {}),
            dict(kwargs.get("failed_items") or {}),
        )

    monkeypatch.setattr(main_module.AppConfig, "from_env", lambda: DummyConfig())
    monkeypatch.setattr(main_module, "GraphClient", DummyGraphClient)
    monkeypatch.setattr(main_module, "load_state", lambda _path: state_holder["state"])
    monkeypatch.setattr(main_module, "save_state", fake_save_state)
    monkeypatch.setattr(main_module, "apply_watermark", lambda _s, out, _w: out.write_bytes(b"wm"))

    rc = main_module.run(["--file-name", "target test.docx", "--log-level", "INFO"])

    assert rc == 0
    assert uploads == ["doc-2"]
    assert state_holder["state"].processed_item_ids == frozenset({"doc-2"})


def test_run_list_fields_mode_exits_without_saving_state(monkeypatch, tmp_path: Path) -> None:
    watermark = tmp_path / "wm.png"
    watermark.write_bytes(b"not-used")

    class DummyConfig:
        auth_mode = "certificate"
        state_file = tmp_path / "state.json"
        library_names = ["WatermarkTesting"]
        library_watermark_paths = {"watermarktesting": watermark}
        site_hostname = "contoso.sharepoint.com"
        site_path = "/sites/Test"

    class DummyGraphClient:
        access_identity = "Watermark - Python"

        def __init__(self, _config):  # noqa: ANN001
            pass

        def resolve_site_id(self) -> str:
            return "site-id"

        def list_drives(self, _site_id: str) -> list[dict]:
            return [{"id": "drive-id", "name": "WatermarkTesting"}]

        def list_library_fields(self, _drive_id: str) -> list[dict]:
            return [{"name": "RecordStatus", "displayName": "Record Status"}]

    state_saved = {"called": False}

    monkeypatch.setattr(main_module.AppConfig, "from_env", lambda: DummyConfig())
    monkeypatch.setattr(main_module, "GraphClient", DummyGraphClient)
    monkeypatch.setattr(main_module, "load_state", lambda _path: RunState(None, frozenset(), {}))
    monkeypatch.setattr(
        main_module,
        "save_state",
        lambda _path, _run_started, **_kwargs: state_saved.__setitem__("called", True),
    )

    rc = main_module.run(["--list-fields", "--log-level", "INFO"])

    assert rc == 0
    assert not state_saved["called"]


def test_run_diagnose_libraries_mode_exits_without_saving_state(
    monkeypatch,
    tmp_path: Path,
) -> None:
    watermark = tmp_path / "wm.png"
    watermark.write_bytes(b"not-used")

    class DummyConfig:
        auth_mode = "certificate"
        state_file = tmp_path / "state.json"
        library_names = ["Archive"]
        library_watermark_paths = {"archive": watermark}
        site_hostname = "contoso.sharepoint.com"
        site_path = "/sites/Test"

    class DummyGraphClient:
        access_identity = "Watermark - Python"

        def __init__(self, _config):  # noqa: ANN001
            pass

        def resolve_site_id(self) -> str:
            return "site-id"

        def list_drives(self, _site_id: str) -> list[dict]:
            return [{"id": "drive-id", "name": "Archive", "webUrl": "https://sp/Archive"}]

        def get_library_details(self, _drive_id: str) -> dict:
            return {
                "id": "list-id",
                "displayName": "Archive",
                "name": "Archive",
                "webUrl": "https://sp/Archive",
                "list": {
                    "template": "documentLibrary",
                    "contentTypesEnabled": False,
                    "hidden": False,
                },
                "sharepointIds": {"listId": "sp-list-id"},
            }

        def list_library_fields(self, _drive_id: str) -> list[dict]:
            return [
                {
                    "name": "RequiredCategory",
                    "displayName": "Required Category",
                    "required": True,
                    "readOnly": False,
                    "hidden": False,
                },
                {
                    "name": "ReadOnlyField",
                    "displayName": "Read Only Field",
                    "required": False,
                    "readOnly": True,
                    "hidden": False,
                },
            ]

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

    rc = main_module.run(["--diagnose-libraries", "--log-level", "INFO"])

    assert rc == 0
    assert not state_saved["called"]


def test_run_diagnose_libraries_fails_on_missing_watermark_mapping(
    monkeypatch,
    tmp_path: Path,
) -> None:
    class DummyConfig:
        auth_mode = "certificate"
        state_file = tmp_path / "state.json"
        library_names = ["Archive"]
        library_watermark_paths = {}
        site_hostname = "contoso.sharepoint.com"
        site_path = "/sites/Test"

    class DummyGraphClient:
        access_identity = "Watermark - Python"

        def __init__(self, _config):  # noqa: ANN001
            pass

        def resolve_site_id(self) -> str:
            return "site-id"

        def list_drives(self, _site_id: str) -> list[dict]:
            return [{"id": "drive-id", "name": "Archive"}]

        def get_library_details(self, _drive_id: str) -> dict:
            return {"id": "list-id", "displayName": "Archive", "list": {}}

        def list_library_fields(self, _drive_id: str) -> list[dict]:
            return []

    monkeypatch.setattr(main_module.AppConfig, "from_env", lambda: DummyConfig())
    monkeypatch.setattr(main_module, "GraphClient", DummyGraphClient)
    monkeypatch.setattr(main_module, "load_state", lambda _path: RunState(None, frozenset(), {}))

    rc = main_module.run(["--diagnose-libraries", "--log-level", "INFO"])

    assert rc == 1


def test_run_write_probe_requires_diagnose_libraries(monkeypatch) -> None:
    monkeypatch.setattr(
        main_module.AppConfig,
        "from_env",
        lambda: (_ for _ in ()).throw(AssertionError("config should not load")),
    )

    rc = main_module.run(["--write-probe", "--log-level", "INFO"])

    assert rc == 2


def test_run_diagnose_libraries_write_probe_create_update_delete(
    monkeypatch,
    tmp_path: Path,
) -> None:
    watermark = tmp_path / "wm.png"
    watermark.write_bytes(b"not-used")
    operations: list[tuple[str, str]] = []

    class DummyConfig:
        auth_mode = "certificate"
        state_file = tmp_path / "state.json"
        library_names = ["Archive"]
        library_watermark_paths = {"archive": watermark}
        site_hostname = "contoso.sharepoint.com"
        site_path = "/sites/Test"

    class DummyGraphClient:
        access_identity = "Watermark - Python"

        def __init__(self, _config):  # noqa: ANN001
            pass

        def resolve_site_id(self) -> str:
            return "site-id"

        def list_drives(self, _site_id: str) -> list[dict]:
            return [{"id": "drive-id", "name": "Archive"}]

        def get_library_details(self, _drive_id: str) -> dict:
            return {"id": "list-id", "displayName": "Archive", "list": {}}

        def list_library_fields(self, _drive_id: str) -> list[dict]:
            return []

        def create_root_file(self, _drive_id: str, file_name: str, _data: bytes) -> dict:
            operations.append(("create", file_name))
            return {"id": "probe-id"}

        def upload_file(self, _drive_id: str, item_id: str, _data: bytes) -> None:
            operations.append(("update", item_id))

        def delete_drive_item(self, _drive_id: str, item_id: str) -> None:
            operations.append(("delete", item_id))

    state_saved = {"called": False}

    monkeypatch.setattr(main_module.AppConfig, "from_env", lambda: DummyConfig())
    monkeypatch.setattr(main_module, "GraphClient", DummyGraphClient)
    monkeypatch.setattr(main_module, "load_state", lambda _path: RunState(None, frozenset(), {}))
    monkeypatch.setattr(
        main_module,
        "save_state",
        lambda _path, _run_started, **_kwargs: state_saved.__setitem__("called", True),
    )

    rc = main_module.run(["--diagnose-libraries", "--write-probe", "--log-level", "INFO"])

    assert rc == 0
    assert [operation[0] for operation in operations] == ["create", "update", "delete"]
    assert operations[0][1].startswith("_watermark_app_write_probe_")
    assert operations[0][1].endswith(".txt")
    assert operations[1] == ("update", "probe-id")
    assert operations[2] == ("delete", "probe-id")
    assert not state_saved["called"]


def test_run_checkpoints_successes_and_retries_failed_files(
    monkeypatch,
    tmp_path: Path,
) -> None:
    watermark = tmp_path / "wm.png"
    watermark.write_bytes(b"not-used")

    class DummyConfig:
        auth_mode = "certificate"
        state_file = tmp_path / "state.json"
        library_names = ["Archive"]
        library_watermark_paths = {"archive": watermark}
        site_hostname = "contoso.sharepoint.com"
        site_path = "/sites/Test"

    scenario = {"pass": 1}
    uploads: list[str] = []
    stored_bytes: dict[str, bytes] = {}
    state_holder = {"state": RunState(None, frozenset(), {})}

    class DummyGraphClient:
        access_identity = "Watermark - Python"

        def __init__(self, _config):  # noqa: ANN001
            pass

        def resolve_site_id(self) -> str:
            return "site-id"

        def list_drives(self, _site_id: str) -> list[dict]:
            return [{"id": "drive-id", "name": "Archive"}]

        def iter_changed_files(self, _drive_id: str, delta_link: str | None = None):
            if scenario["pass"] == 1:
                assert delta_link is None
                return [
                    {"id": "ok-1", "name": "ok.docx", "webUrl": "https://sp/ok.docx"},
                    {"id": "bad-1", "name": "bad.docx", "webUrl": "https://sp/bad.docx"},
                ], "delta-1"
            assert delta_link == "delta-1"
            return [
                {"id": "ok-2", "name": "new.docx", "webUrl": "https://sp/new.docx"},
            ], "delta-2"

        def get_drive_item(self, _drive_id: str, item_id: str) -> dict:
            assert item_id == "bad-1"
            return {"id": "bad-1", "name": "bad.docx", "webUrl": "https://sp/bad.docx", "file": {}}

        def download_file(self, _drive_id: str, item_id: str) -> bytes:
            return stored_bytes.get(item_id, b"fake")

        def upload_file(self, _drive_id: str, item_id: str, _data: bytes) -> None:
            uploads.append(item_id)
            if scenario["pass"] == 1 and item_id == "bad-1":
                raise RuntimeError("locked")
            stored_bytes[item_id] = _data

    def fake_save_state(_path, run_started, **kwargs):  # noqa: ANN001
        ids = kwargs.get("processed_item_ids") or set()
        links = kwargs.get("drive_delta_links") or {}
        failed_items = kwargs.get("failed_items") or {}
        state_holder["state"] = RunState(
            run_started,
            frozenset(ids),
            dict(links),
            dict(failed_items),
        )

    monkeypatch.setattr(main_module.AppConfig, "from_env", lambda: DummyConfig())
    monkeypatch.setattr(main_module, "GraphClient", DummyGraphClient)
    monkeypatch.setattr(main_module, "load_state", lambda _path: state_holder["state"])
    monkeypatch.setattr(main_module, "save_state", fake_save_state)
    monkeypatch.setattr(main_module, "apply_watermark", lambda _s, out, _w: out.write_bytes(b"wm"))

    rc1 = main_module.run(["--log-level", "INFO"])

    assert rc1 == 1
    assert uploads == ["ok-1", "bad-1"]
    assert state_holder["state"].processed_item_ids == frozenset({"ok-1"})
    assert state_holder["state"].drive_delta_links == {"drive-id": "delta-1"}
    assert state_holder["state"].failed_items == {
        "drive-id": [
            {
                "id": "bad-1",
                "name": "bad.docx",
                "webUrl": "https://sp/bad.docx",
                "error": "locked",
            }
        ]
    }

    scenario["pass"] = 2
    rc2 = main_module.run(["--log-level", "INFO"])

    assert rc2 == 0
    assert uploads == ["ok-1", "bad-1", "bad-1", "ok-2"]
    assert state_holder["state"].processed_item_ids == frozenset({"ok-1", "bad-1", "ok-2"})
    assert state_holder["state"].drive_delta_links == {"drive-id": "delta-2"}
    assert state_holder["state"].failed_items == {}


def test_run_first_file_only_processes_one_file_without_advancing_delta(
    monkeypatch,
    tmp_path: Path,
) -> None:
    watermark = tmp_path / "wm.png"
    watermark.write_bytes(b"not-used")
    uploads: list[str] = []
    stored_bytes: dict[str, bytes] = {}
    state_holder = {"state": RunState(None, frozenset(), {"drive-id": "delta-old"})}

    class DummyConfig:
        auth_mode = "certificate"
        state_file = tmp_path / "state.json"
        library_names = ["Archive"]
        library_watermark_paths = {"archive": watermark}
        site_hostname = "contoso.sharepoint.com"
        site_path = "/sites/Test"

    class DummyGraphClient:
        access_identity = "Watermark - Python"

        def __init__(self, _config):  # noqa: ANN001
            pass

        def resolve_site_id(self) -> str:
            return "site-id"

        def list_drives(self, _site_id: str) -> list[dict]:
            return [{"id": "drive-id", "name": "Archive"}]

        def iter_changed_files(self, _drive_id: str, delta_link: str | None = None):
            assert delta_link == "delta-old"
            return [
                {"id": "txt-1", "name": "notes.txt", "webUrl": "https://sp/notes.txt"},
                {"id": "doc-1", "name": "first.docx", "webUrl": "https://sp/first.docx"},
                {"id": "doc-2", "name": "second.docx", "webUrl": "https://sp/second.docx"},
            ], "delta-new"

        def download_file(self, _drive_id: str, item_id: str) -> bytes:
            return stored_bytes.get(item_id, b"fake")

        def upload_file(self, _drive_id: str, item_id: str, _data: bytes) -> None:
            uploads.append(item_id)
            stored_bytes[item_id] = _data

    def fake_save_state(_path, run_started, **kwargs):  # noqa: ANN001
        ids = kwargs.get("processed_item_ids") or set()
        links = kwargs.get("drive_delta_links") or {}
        failed_items = kwargs.get("failed_items") or {}
        state_holder["state"] = RunState(
            run_started,
            frozenset(ids),
            dict(links),
            dict(failed_items),
        )

    monkeypatch.setattr(main_module.AppConfig, "from_env", lambda: DummyConfig())
    monkeypatch.setattr(main_module, "GraphClient", DummyGraphClient)
    monkeypatch.setattr(main_module, "load_state", lambda _path: state_holder["state"])
    monkeypatch.setattr(main_module, "save_state", fake_save_state)
    monkeypatch.setattr(main_module, "apply_watermark", lambda _s, out, _w: out.write_bytes(b"wm"))

    rc = main_module.run(["--first-file-only", "--log-level", "INFO"])

    assert rc == 0
    assert uploads == ["doc-1"]
    assert state_holder["state"].processed_item_ids == frozenset({"doc-1"})
    assert state_holder["state"].drive_delta_links == {"drive-id": "delta-old"}
    assert state_holder["state"].failed_items == {}


def test_run_fails_when_post_upload_verification_returns_different_bytes(
    monkeypatch,
    tmp_path: Path,
) -> None:
    watermark = tmp_path / "wm.png"
    watermark.write_bytes(b"not-used")
    state_holder = {"state": RunState(None, frozenset(), {})}

    class DummyConfig:
        auth_mode = "certificate"
        state_file = tmp_path / "state.json"
        library_names = ["Archive"]
        library_watermark_paths = {"archive": watermark}
        site_hostname = "contoso.sharepoint.com"
        site_path = "/sites/Test"

    class DummyGraphClient:
        access_identity = "Watermark - Python"

        def __init__(self, _config):  # noqa: ANN001
            pass

        def resolve_site_id(self) -> str:
            return "site-id"

        def list_drives(self, _site_id: str) -> list[dict]:
            return [{"id": "drive-id", "name": "Archive"}]

        def iter_changed_files(self, _drive_id: str, _delta_link: str | None = None):
            return [
                {"id": "doc-1", "name": "first.docx", "webUrl": "https://sp/first.docx"}
            ], "delta-1"

        def download_file(self, _drive_id: str, _item_id: str) -> bytes:
            return b"old bytes"

        def upload_file(self, _drive_id: str, _item_id: str, _data: bytes) -> None:
            pass

    def fake_save_state(_path, run_started, **kwargs):  # noqa: ANN001
        state_holder["state"] = RunState(
            run_started,
            frozenset(kwargs.get("processed_item_ids") or set()),
            dict(kwargs.get("drive_delta_links") or {}),
            dict(kwargs.get("failed_items") or {}),
        )

    monkeypatch.setattr(main_module.AppConfig, "from_env", lambda: DummyConfig())
    monkeypatch.setattr(main_module, "GraphClient", DummyGraphClient)
    monkeypatch.setattr(main_module, "load_state", lambda _path: state_holder["state"])
    monkeypatch.setattr(main_module, "save_state", fake_save_state)
    monkeypatch.setattr(main_module, "apply_watermark", lambda _s, out, _w: out.write_bytes(b"wm"))

    rc = main_module.run(["--log-level", "INFO"])

    assert rc == 1
    assert state_holder["state"].processed_item_ids == frozenset()
    assert state_holder["state"].failed_items == {
        "drive-id": [
            {
                "id": "doc-1",
                "name": "first.docx",
                "webUrl": "https://sp/first.docx",
                "error": (
                    "post-upload verification failed; SharePoint returned different bytes "
                    "than the app uploaded"
                ),
            }
        ]
    }


def test_run_save_diagnostics_writes_original_output_and_post_upload_files(
    monkeypatch,
    tmp_path: Path,
) -> None:
    watermark = tmp_path / "wm.png"
    watermark.write_bytes(b"not-used")
    diagnostics_dir = tmp_path / "diagnostics"
    state_holder = {"state": RunState(None, frozenset(), {})}
    downloads = iter([b"original", b"sharepoint-after-upload"])

    class DummyConfig:
        auth_mode = "certificate"
        state_file = tmp_path / "state.json"
        library_names = ["Archive"]
        library_watermark_paths = {"archive": watermark}
        site_hostname = "contoso.sharepoint.com"
        site_path = "/sites/Test"

    class DummyGraphClient:
        access_identity = "Watermark - Python"

        def __init__(self, _config):  # noqa: ANN001
            pass

        def resolve_site_id(self) -> str:
            return "site-id"

        def list_drives(self, _site_id: str) -> list[dict]:
            return [{"id": "drive-id", "name": "Archive"}]

        def iter_changed_files(self, _drive_id: str, _delta_link: str | None = None):
            return [
                {"id": "doc-1", "name": "first.docx", "webUrl": "https://sp/first.docx"}
            ], "delta-1"

        def download_file(self, _drive_id: str, _item_id: str) -> bytes:
            return next(downloads)

        def upload_file(self, _drive_id: str, _item_id: str, _data: bytes) -> None:
            pass

    monkeypatch.setattr(main_module.AppConfig, "from_env", lambda: DummyConfig())
    monkeypatch.setattr(main_module, "GraphClient", DummyGraphClient)
    monkeypatch.setattr(main_module, "load_state", lambda _path: state_holder["state"])
    monkeypatch.setattr(main_module, "save_state", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(
        main_module,
        "apply_watermark",
        lambda _source, output, _watermark: output.write_bytes(b"watermarked"),
    )

    rc = main_module.run(
        ["--first-file-only", "--save-diagnostics", str(diagnostics_dir), "--log-level", "INFO"]
    )

    assert rc == 1
    saved_files = {path.name: path.read_bytes() for path in diagnostics_dir.rglob("*.docx")}
    assert saved_files == {
        "01_original_download_first.docx": b"original",
        "02_local_watermarked_first.docx": b"watermarked",
        "03_sharepoint_after_upload_first.docx": b"sharepoint-after-upload",
    }
