from pathlib import Path

import pytest

from watermark_app.config import AppConfig


def _set_base_env(monkeypatch: pytest.MonkeyPatch, watermark_file: Path) -> None:
    monkeypatch.setenv("AZURE_TENANT_ID", "tenant-id")
    monkeypatch.setenv("AZURE_CLIENT_ID", "client-id")
    monkeypatch.setenv("AZURE_CLIENT_SECRET", "client-secret")
    monkeypatch.setenv("SP_SITE_HOSTNAME", "contoso.sharepoint.com")
    monkeypatch.setenv("SP_SITE_PATH", "/sites/Finance")
    monkeypatch.setenv("WATERMARK_IMAGE_PATH", str(watermark_file))
    monkeypatch.setenv("SP_LIBRARY_WATERMARKS", "")
    monkeypatch.delenv("CLOUD_ENV", raising=False)


def test_from_env_defaults_to_commercial(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    watermark = tmp_path / "wm.png"
    watermark.write_bytes(b"fake")
    _set_base_env(monkeypatch, watermark)

    config = AppConfig.from_env()

    assert config.cloud_env == "commercial"
    assert config.authority_host == "https://login.microsoftonline.com"
    assert config.graph_base_url == "https://graph.microsoft.com/v1.0"
    assert config.graph_scope == "https://graph.microsoft.com/.default"


def test_from_env_gcch_endpoints(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    watermark = tmp_path / "wm.png"
    watermark.write_bytes(b"fake")
    _set_base_env(monkeypatch, watermark)
    monkeypatch.setenv("CLOUD_ENV", "gcch")
    monkeypatch.setenv("SP_SITE_HOSTNAME", "contoso.sharepoint.us")

    config = AppConfig.from_env()

    assert config.cloud_env == "gcch"
    assert config.authority_host == "https://login.microsoftonline.us"
    assert config.graph_base_url == "https://graph.microsoft.us/v1.0"
    assert config.graph_scope == "https://graph.microsoft.us/.default"


def test_from_env_rejects_unsupported_cloud(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    watermark = tmp_path / "wm.png"
    watermark.write_bytes(b"fake")
    _set_base_env(monkeypatch, watermark)
    monkeypatch.setenv("CLOUD_ENV", "dod")

    with pytest.raises(ValueError, match="Unsupported CLOUD_ENV"):
        AppConfig.from_env()


def test_from_env_library_watermark_map(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    default_wm = tmp_path / "default.png"
    docs_wm = tmp_path / "docs.png"
    legal_wm = tmp_path / "legal.png"
    default_wm.write_bytes(b"default")
    docs_wm.write_bytes(b"docs")
    legal_wm.write_bytes(b"legal")
    _set_base_env(monkeypatch, default_wm)
    monkeypatch.setenv(
        "SP_LIBRARY_WATERMARKS",
        f"Documents={docs_wm};Legal Docs={legal_wm}",
    )

    config = AppConfig.from_env()

    assert config.library_watermark_paths["documents"] == docs_wm.resolve()
    assert config.library_watermark_paths["legal docs"] == legal_wm.resolve()


def test_from_env_rejects_invalid_library_watermark_map(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    default_wm = tmp_path / "default.png"
    default_wm.write_bytes(b"default")
    _set_base_env(monkeypatch, default_wm)
    monkeypatch.setenv("SP_LIBRARY_WATERMARKS", "Documents")

    with pytest.raises(ValueError, match="Invalid SP_LIBRARY_WATERMARKS entry"):
        AppConfig.from_env()
