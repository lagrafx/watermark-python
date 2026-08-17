"""Microsoft Graph client for SharePoint document libraries."""

from __future__ import annotations

import base64
import json
import logging
from dataclasses import dataclass

import msal
import requests

from watermark_app.config import AppConfig

LOG = logging.getLogger(__name__)


class GraphClientError(RuntimeError):
    """Raised for Microsoft Graph request/authentication failures."""


@dataclass
class GraphClient:
    config: AppConfig

    def __post_init__(self) -> None:
        authority = f"{self.config.authority_host}/{self.config.tenant_id}"
        client_credential: str | dict[str, str]
        if self.config.auth_mode == "certificate":
            if not self.config.client_cert_pfx_path:
                raise GraphClientError("Certificate auth selected but no PFX path was configured.")
            client_credential = {
                "private_key_pfx_path": str(self.config.client_cert_pfx_path),
            }
            if self.config.client_cert_pfx_password:
                client_credential["passphrase"] = self.config.client_cert_pfx_password
        else:
            if not self.config.client_secret:
                raise GraphClientError(
                    "Client secret auth selected but AZURE_CLIENT_SECRET is empty."
                )
            client_credential = self.config.client_secret

        self._msal_app = msal.ConfidentialClientApplication(
            client_id=self.config.client_id,
            client_credential=client_credential,
            authority=authority,
        )
        self._scope = [self.config.graph_scope]
        token_result = self._acquire_access_token()
        token = token_result.get("access_token")
        if not token:
            raise GraphClientError(f"Failed to acquire access token: {token_result}")
        self._headers = {"Authorization": f"Bearer {token}"}
        self._token_claims = self._decode_token_claims(token)

    @property
    def access_identity(self) -> str:
        """Return the application identity used for Graph access."""
        for claim in ("app_displayname", "appid", "azp", "client_id"):
            value = self._token_claims.get(claim)
            if isinstance(value, str) and value.strip():
                return value
        return self.config.client_id

    def resolve_site_id(self) -> str:
        site_path = self.config.site_path
        if not site_path.startswith("/"):
            site_path = "/" + site_path
        url = f"{self.config.graph_base_url}/sites/{self.config.site_hostname}:{site_path}"
        response = self._request("GET", url, operation="resolve site", timeout=60)
        self._raise_for_error(response, "resolve site")
        return response.json()["id"]

    def list_drives(self, site_id: str) -> list[dict]:
        response = self._request(
            "GET",
            f"{self.config.graph_base_url}/sites/{site_id}/drives",
            operation="list drives",
            timeout=60,
        )
        self._raise_for_error(response, "list drives")
        return response.json().get("value", [])

    def list_library_fields(self, drive_id: str) -> list[dict]:
        response = self._request(
            "GET",
            f"{self.config.graph_base_url}/drives/{drive_id}/list/columns",
            operation="list library fields",
            timeout=60,
        )
        self._raise_for_error(response, "list library fields")
        return response.json().get("value", [])

    def get_library_details(self, drive_id: str) -> dict:
        response = self._request(
            "GET",
            f"{self.config.graph_base_url}/drives/{drive_id}/list",
            operation="get library details",
            timeout=60,
        )
        self._raise_for_error(response, "get library details")
        return response.json()

    def iter_files(self, drive_id: str) -> list[dict]:
        files: list[dict] = []
        queue: list[str] = [f"{self.config.graph_base_url}/drives/{drive_id}/root/children"]
        while queue:
            url = queue.pop(0)
            response = self._request("GET", url, operation="list drive items", timeout=60)
            self._raise_for_error(response, "list drive items")
            payload = response.json()
            for item in payload.get("value", []):
                if "folder" in item:
                    queue.append(
                        f"{self.config.graph_base_url}/drives/{drive_id}/items/{item['id']}/children"
                    )
                elif "file" in item:
                    files.append(item)
            next_link = payload.get("@odata.nextLink")
            if next_link:
                queue.append(next_link)
        return files

    def iter_changed_files(
        self, drive_id: str, delta_link: str | None = None
    ) -> tuple[list[dict], str]:
        files: list[dict] = []
        if delta_link:
            url = delta_link
        else:
            url = f"{self.config.graph_base_url}/drives/{drive_id}/root/delta"

        final_delta_link: str | None = None
        page_number = 0
        total_raw_items = 0
        total_files = 0
        total_folders = 0
        total_deleted = 0
        total_other = 0
        while url:
            page_number += 1
            response = self._request("GET", url, operation="list changed drive items", timeout=60)
            self._raise_for_error(response, "list changed drive items")
            payload = response.json()
            values = payload.get("value", [])
            if not isinstance(values, list):
                raise GraphClientError(
                    "Failed to list changed drive items: response 'value' was not a list"
                )

            page_files = 0
            page_folders = 0
            page_deleted = 0
            page_other = 0
            for item in values:
                if not isinstance(item, dict):
                    page_other += 1
                    continue
                if "deleted" in item:
                    page_deleted += 1
                    continue
                if "file" in item and "deleted" not in item:
                    page_files += 1
                    files.append(item)
                    continue
                if "folder" in item:
                    page_folders += 1
                    continue
                page_other += 1

            next_link = payload.get("@odata.nextLink")
            page_raw_items = len(values)
            total_raw_items += page_raw_items
            total_files += page_files
            total_folders += page_folders
            total_deleted += page_deleted
            total_other += page_other
            LOG.info(
                "Delta page %s for drive %s: raw_items=%s files=%s folders=%s "
                "deleted=%s other=%s next_link=%s delta_link=%s",
                page_number,
                drive_id,
                page_raw_items,
                page_files,
                page_folders,
                page_deleted,
                page_other,
                bool(next_link),
                bool(payload.get("@odata.deltaLink")),
            )
            if next_link:
                url = next_link
                continue
            final_delta_link = payload.get("@odata.deltaLink")
            break

        if not final_delta_link:
            raise GraphClientError("Failed to list changed drive items: missing @odata.deltaLink")
        LOG.info(
            "Delta complete for drive %s: pages=%s raw_items=%s files=%s folders=%s "
            "deleted=%s other=%s",
            drive_id,
            page_number,
            total_raw_items,
            total_files,
            total_folders,
            total_deleted,
            total_other,
        )
        return files, final_delta_link

    def download_file(self, drive_id: str, item_id: str) -> bytes:
        response = self._request(
            "GET",
            f"{self.config.graph_base_url}/drives/{drive_id}/items/{item_id}/content",
            operation="download file",
            timeout=120,
        )
        self._raise_for_error(response, "download file")
        return response.content

    def get_drive_item(self, drive_id: str, item_id: str) -> dict:
        response = self._request(
            "GET",
            f"{self.config.graph_base_url}/drives/{drive_id}/items/{item_id}",
            operation="get drive item",
            timeout=60,
        )
        self._raise_for_error(response, "get drive item")
        return response.json()

    def upload_file(self, drive_id: str, item_id: str, data: bytes) -> None:
        response = self._request(
            "PUT",
            f"{self.config.graph_base_url}/drives/{drive_id}/items/{item_id}/content",
            operation="upload file",
            headers={"Content-Type": "application/octet-stream"},
            data=data,
            timeout=120,
        )
        self._raise_for_error(response, "upload file")

    def create_root_file(self, drive_id: str, file_name: str, data: bytes) -> dict:
        response = self._request(
            "PUT",
            f"{self.config.graph_base_url}/drives/{drive_id}/root:/{file_name}:/content",
            operation="create root file",
            headers={"Content-Type": "text/plain"},
            data=data,
            timeout=120,
        )
        self._raise_for_error(response, "create root file")
        return response.json()

    def delete_drive_item(self, drive_id: str, item_id: str) -> None:
        response = self._request(
            "DELETE",
            f"{self.config.graph_base_url}/drives/{drive_id}/items/{item_id}",
            operation="delete drive item",
            timeout=120,
        )
        self._raise_for_error(response, "delete drive item")

    def _request(
        self,
        method: str,
        url: str,
        operation: str,
        timeout: int,
        headers: dict[str, str] | None = None,
        **kwargs: object,
    ) -> requests.Response:
        response = requests.request(
            method=method,
            url=url,
            headers={**self._headers, **(headers or {})},
            timeout=timeout,
            **kwargs,
        )
        if response.status_code == 401 and self._is_invalid_token_error(response):
            token_result = self._acquire_access_token()
            token = token_result.get("access_token")
            if not token:
                raise GraphClientError(f"Failed to refresh access token: {token_result}")
            self._headers = {"Authorization": f"Bearer {token}"}
            self._token_claims = self._decode_token_claims(token)
            response = requests.request(
                method=method,
                url=url,
                headers={**self._headers, **(headers or {})},
                timeout=timeout,
                **kwargs,
            )
        return response

    def _acquire_access_token(self) -> dict:
        return self._msal_app.acquire_token_for_client(scopes=self._scope)

    @staticmethod
    def _decode_token_claims(token: str) -> dict:
        parts = token.split(".")
        if len(parts) < 2:
            return {}
        payload = parts[1]
        padding = "=" * (-len(payload) % 4)
        try:
            decoded = base64.urlsafe_b64decode(f"{payload}{padding}")
            claims = json.loads(decoded)
        except Exception:  # noqa: BLE001
            return {}
        return claims if isinstance(claims, dict) else {}

    @staticmethod
    def _is_invalid_token_error(response: requests.Response) -> bool:
        try:
            payload = response.json()
        except Exception:  # noqa: BLE001
            return False
        if not isinstance(payload, dict):
            return False
        error = payload.get("error")
        if not isinstance(error, dict):
            return False
        return error.get("code") == "InvalidAuthenticationToken"

    @staticmethod
    def _raise_for_error(response: requests.Response, operation: str) -> None:
        if response.ok:
            return
        detail = None
        try:
            detail = response.json()
        except Exception:  # noqa: BLE001
            detail = response.text
        hint = ""
        if response.status_code == 403:
            hint = (
                " Hint: Access denied. If using Graph Application permission "
                "'Sites.Selected', grant this app site-level permission to the target "
                "SharePoint site (for example, write access)."
            )
        raise GraphClientError(
            f"Failed to {operation}: HTTP {response.status_code} {detail}{hint}"
        )
