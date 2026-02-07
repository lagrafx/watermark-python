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


@dataclass(frozen=True)
class AppConfig:
    tenant_id: str
    client_id: str
    client_secret: str
    site_hostname: str
    site_path: str
    watermark_image_path: Path
    state_file: Path
    library_names: tuple[str, ...]

    @classmethod
    def from_env(cls) -> "AppConfig":
        load_dotenv()
        watermark_path = Path(_required_env("WATERMARK_IMAGE_PATH")).expanduser().resolve()
        if not watermark_path.exists():
            raise ValueError(f"Watermark image not found: {watermark_path}")

        libraries_csv = os.getenv("SP_LIBRARY_NAMES", "").strip()
        library_names = tuple(
            part.strip() for part in libraries_csv.split(",") if part.strip()
        )
        state_file = Path(os.getenv("STATE_FILE", ".watermark_state.json")).expanduser().resolve()

        return cls(
            tenant_id=_required_env("AZURE_TENANT_ID"),
            client_id=_required_env("AZURE_CLIENT_ID"),
            client_secret=_required_env("AZURE_CLIENT_SECRET"),
            site_hostname=_required_env("SP_SITE_HOSTNAME"),
            site_path=_required_env("SP_SITE_PATH"),
            watermark_image_path=watermark_path,
            state_file=state_file,
            library_names=library_names,
        )
