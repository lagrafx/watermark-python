from pathlib import Path

import pytest

from watermark_app.config import AppConfig


def _set_base_env(monkeypatch: pytest.MonkeyPatch, watermark_file: Path) -> None:
    monkeypatch.setenv("AZURE_TENANT_ID", "tenant-id")
    monkeypatch.setenv("AZURE_CLIENT_ID", "client-id")
    monkeypatch.setenv("AZURE_CLIENT_SECRET", "client-secret")
    monkeypatch.delenv("AZURE_CLIENT_CERT_PFX_PATH", raising=False)
    monkeypatch.delenv("AZURE_CLIENT_CERT_PFX_PASSWORD", raising=False)
    monkeypatch.setenv("SP_SITE_HOSTNAME", "contoso.sharepoint.com")
    monkeypatch.setenv("SP_SITE_PATH", "/sites/Finance")
    monkeypatch.setenv("SP_LIBRARY_NAMES", "Documents")
    monkeypatch.setenv("SP_LIBRARY_WATERMARKS", f"Documents={watermark_file}")
    monkeypatch.delenv("CLOUD_ENV", raising=False)


def test_from_env_defaults_to_commercial(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    watermark = tmp_path / "wm.png"
    watermark.write_bytes(b"fake")
    _set_base_env(monkeypatch, watermark)

    config = AppConfig.from_env()

    assert config.cloud_env == "commercial"
    assert config.auth_mode == "client_secret"
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


def test_from_env_certificate_auth_selected_when_pfx_is_set(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    watermark = tmp_path / "wm.png"
    cert = tmp_path / "app.pfx"
    watermark.write_bytes(b"fake")
    cert.write_bytes(b"fake")
    _set_base_env(monkeypatch, watermark)
    monkeypatch.setenv("AZURE_CLIENT_CERT_PFX_PATH", str(cert))
    monkeypatch.setenv("AZURE_CLIENT_CERT_PFX_PASSWORD", "secret-passphrase")

    config = AppConfig.from_env()

    assert config.auth_mode == "certificate"
    assert config.client_cert_pfx_path == cert.resolve()
    assert config.client_cert_pfx_password == "secret-passphrase"


def test_from_env_requires_secret_or_cert(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    watermark = tmp_path / "wm.png"
    watermark.write_bytes(b"fake")
    _set_base_env(monkeypatch, watermark)
    monkeypatch.setenv("AZURE_CLIENT_SECRET", "")

    with pytest.raises(ValueError, match="Missing app credential"):
        AppConfig.from_env()


def test_from_env_library_watermark_map(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    docs_wm = tmp_path / "docs.png"
    legal_wm = tmp_path / "legal.png"
    docs_wm.write_bytes(b"docs")
    legal_wm.write_bytes(b"legal")
    _set_base_env(monkeypatch, docs_wm)
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
    docs_wm = tmp_path / "docs.png"
    docs_wm.write_bytes(b"docs")
    _set_base_env(monkeypatch, docs_wm)
    monkeypatch.setenv("SP_LIBRARY_WATERMARKS", "Documents")

    with pytest.raises(ValueError, match="Invalid SP_LIBRARY_WATERMARKS entry"):
        AppConfig.from_env()


def test_from_env_requires_library_watermark_mapping(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    docs_wm = tmp_path / "docs.png"
    docs_wm.write_bytes(b"docs")
    _set_base_env(monkeypatch, docs_wm)
    monkeypatch.setenv("SP_LIBRARY_WATERMARKS", "")

    with pytest.raises(ValueError, match="SP_LIBRARY_WATERMARKS is required"):
        AppConfig.from_env()
