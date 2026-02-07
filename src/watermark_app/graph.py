"""Microsoft Graph client for SharePoint document libraries."""

from __future__ import annotations

from dataclasses import dataclass

import msal
import requests

from watermark_app.config import AppConfig

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


class GraphClientError(RuntimeError):
    """Raised for Microsoft Graph request/authentication failures."""


@dataclass
class GraphClient:
    config: AppConfig

    def __post_init__(self) -> None:
        authority = f"https://login.microsoftonline.com/{self.config.tenant_id}"
        self._msal_app = msal.ConfidentialClientApplication(
            client_id=self.config.client_id,
            client_credential=self.config.client_secret,
            authority=authority,
        )
        token_result = self._msal_app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )
        token = token_result.get("access_token")
        if not token:
            raise GraphClientError(f"Failed to acquire access token: {token_result}")
        self._headers = {"Authorization": f"Bearer {token}"}

    def resolve_site_id(self) -> str:
        site_path = self.config.site_path
        if not site_path.startswith("/"):
            site_path = "/" + site_path
        url = f"{GRAPH_BASE}/sites/{self.config.site_hostname}:{site_path}"
        response = requests.get(url, headers=self._headers, timeout=60)
        self._raise_for_error(response, "resolve site")
        return response.json()["id"]

    def list_drives(self, site_id: str) -> list[dict]:
        response = requests.get(
            f"{GRAPH_BASE}/sites/{site_id}/drives",
            headers=self._headers,
            timeout=60,
        )
        self._raise_for_error(response, "list drives")
        return response.json().get("value", [])

    def iter_files(self, drive_id: str) -> list[dict]:
        files: list[dict] = []
        queue: list[str] = [f"{GRAPH_BASE}/drives/{drive_id}/root/children"]
        while queue:
            url = queue.pop(0)
            response = requests.get(url, headers=self._headers, timeout=60)
            self._raise_for_error(response, "list drive items")
            payload = response.json()
            for item in payload.get("value", []):
                if "folder" in item:
                    queue.append(f"{GRAPH_BASE}/drives/{drive_id}/items/{item['id']}/children")
                elif "file" in item:
                    files.append(item)
            next_link = payload.get("@odata.nextLink")
            if next_link:
                queue.append(next_link)
        return files

    def download_file(self, drive_id: str, item_id: str) -> bytes:
        response = requests.get(
            f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/content",
            headers=self._headers,
            timeout=120,
        )
        self._raise_for_error(response, "download file")
        return response.content

    def upload_file(self, drive_id: str, item_id: str, data: bytes) -> None:
        response = requests.put(
            f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/content",
            headers={**self._headers, "Content-Type": "application/octet-stream"},
            data=data,
            timeout=120,
        )
        self._raise_for_error(response, "upload file")

    @staticmethod
    def _raise_for_error(response: requests.Response, operation: str) -> None:
        if response.ok:
            return
        try:
            detail = response.json()
        except Exception:  # noqa: BLE001
            detail = response.text
        raise GraphClientError(f"Failed to {operation}: HTTP {response.status_code} {detail}")
