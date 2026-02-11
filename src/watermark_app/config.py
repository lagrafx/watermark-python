"""Configuration loading for the watermark application."""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv


def _required_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise ValueError(f"Missing required environment variable: {name}")
    return value


def _resolve_cloud_env(raw: str) -> tuple[str, str, str, str]:
    normalized = raw.strip().lower().replace("_", "-")
    if normalized in {"", "commercial"}:
        return (
            "commercial",
            "https://login.microsoftonline.com",
            "https://graph.microsoft.com/v1.0",
            "https://graph.microsoft.com/.default",
        )
    if normalized in {"gcch", "gcc-high", "gcc high"}:
        return (
            "gcch",
            "https://login.microsoftonline.us",
            "https://graph.microsoft.us/v1.0",
            "https://graph.microsoft.us/.default",
        )
    raise ValueError(
        f"Unsupported CLOUD_ENV '{raw}'. Supported values: commercial, gcch"
    )


def _parse_library_watermarks(raw: str) -> dict[str, Path]:
    mapping: dict[str, Path] = {}
    if not raw.strip():
        return mapping

    for pair in raw.split(";"):
        pair = pair.strip()
        if not pair:
            continue
        if "=" not in pair:
            raise ValueError(
                "Invalid SP_LIBRARY_WATERMARKS entry. Expected format "
                "'Library Name=C:\\path\\file.png;Other Library=C:\\path\\other.png'"
            )
        library_name, watermark_raw = pair.split("=", 1)
        library_name = library_name.strip()
        watermark_raw = watermark_raw.strip()
        if not library_name or not watermark_raw:
            raise ValueError(
                "Invalid SP_LIBRARY_WATERMARKS entry. Library name and path are required."
            )
        watermark_path = Path(watermark_raw).expanduser().resolve()
        if not watermark_path.exists():
            raise ValueError(
                f"Watermark image for library '{library_name}' not found: {watermark_path}"
            )
        mapping[library_name.lower()] = watermark_path
    return mapping


@dataclass(frozen=True)
class AppConfig:
    tenant_id: str
    client_id: str
    client_secret: str
    cloud_env: str
    authority_host: str
    graph_base_url: str
    graph_scope: str
    site_hostname: str
    site_path: str
    watermark_image_path: Path
    library_watermark_paths: dict[str, Path]
    state_file: Path
    library_names: tuple[str, ...]

    @classmethod
    def from_env(cls) -> "AppConfig":
        load_dotenv()
        cloud_env, authority_host, graph_base_url, graph_scope = _resolve_cloud_env(
            os.getenv("CLOUD_ENV", "")
        )
        watermark_path = Path(_required_env("WATERMARK_IMAGE_PATH")).expanduser().resolve()
        if not watermark_path.exists():
            raise ValueError(f"Watermark image not found: {watermark_path}")

        libraries_csv = os.getenv("SP_LIBRARY_NAMES", "").strip()
        library_names = tuple(
            part.strip() for part in libraries_csv.split(",") if part.strip()
        )
        library_watermark_paths = _parse_library_watermarks(
            os.getenv("SP_LIBRARY_WATERMARKS", "")
        )
        state_file = Path(os.getenv("STATE_FILE", ".watermark_state.json")).expanduser().resolve()

        return cls(
            tenant_id=_required_env("AZURE_TENANT_ID"),
            client_id=_required_env("AZURE_CLIENT_ID"),
            client_secret=_required_env("AZURE_CLIENT_SECRET"),
            cloud_env=cloud_env,
            authority_host=authority_host,
            graph_base_url=graph_base_url,
            graph_scope=graph_scope,
            site_hostname=_required_env("SP_SITE_HOSTNAME"),
            site_path=_required_env("SP_SITE_PATH"),
            watermark_image_path=watermark_path,
            library_watermark_paths=library_watermark_paths,
            state_file=state_file,
            library_names=library_names,
        )
