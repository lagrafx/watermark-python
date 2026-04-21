"""Microsoft Graph client for SharePoint document libraries."""

from __future__ import annotations

from dataclasses import dataclass

import msal
import requests

from watermark_app.config import AppConfig


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
        while url:
            response = self._request("GET", url, operation="list changed drive items", timeout=60)
            self._raise_for_error(response, "list changed drive items")
            payload = response.json()
            for item in payload.get("value", []):
                if "file" in item and "deleted" not in item:
                    files.append(item)
            next_link = payload.get("@odata.nextLink")
            if next_link:
                url = next_link
                continue
            final_delta_link = payload.get("@odata.deltaLink")
            break

        if not final_delta_link:
            raise GraphClientError("Failed to list changed drive items: missing @odata.deltaLink")
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
