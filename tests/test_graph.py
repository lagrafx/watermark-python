import pytest

from watermark_app.graph import GraphClient, GraphClientError


class _DummyResponse:
    def __init__(
        self,
        ok: bool,
        status_code: int,
        payload: dict | None = None,
        text: str = "",
    ) -> None:
        self.ok = ok
        self.status_code = status_code
        self._payload = payload
        self.text = text

    def json(self) -> dict:
        if self._payload is None:
            raise ValueError("No JSON payload")
        return self._payload


def test_raise_for_error_adds_sites_selected_hint_on_403() -> None:
    response = _DummyResponse(
        ok=False,
        status_code=403,
        payload={"error": {"code": "accessDenied", "message": "Access denied"}},
    )

    try:
        GraphClient._raise_for_error(response, "resolve site")
        raise AssertionError("Expected GraphClientError")
    except GraphClientError as exc:
        message = str(exc)
        assert "Failed to resolve site: HTTP 403" in message
        assert "Sites.Selected" in message
        assert "site-level permission" in message


def test_raise_for_error_no_hint_for_non_403() -> None:
    response = _DummyResponse(ok=False, status_code=404, payload={"error": "notFound"})

    try:
        GraphClient._raise_for_error(response, "resolve site")
        raise AssertionError("Expected GraphClientError")
    except GraphClientError as exc:
        message = str(exc)
        assert "Failed to resolve site: HTTP 404" in message
        assert "Sites.Selected" not in message


def test_request_refreshes_token_and_retries_once(monkeypatch: pytest.MonkeyPatch) -> None:
    client = GraphClient.__new__(GraphClient)
    client._headers = {"Authorization": "Bearer old-token"}

    calls: list[dict] = []

    def fake_request(method, url, headers, timeout, **kwargs):  # noqa: ANN001
        calls.append({"method": method, "url": url, "headers": headers, "timeout": timeout})
        if len(calls) == 1:
            return _DummyResponse(
                ok=False,
                status_code=401,
                payload={"error": {"code": "InvalidAuthenticationToken"}},
            )
        return _DummyResponse(ok=True, status_code=200, payload={"ok": True})

    monkeypatch.setattr("watermark_app.graph.requests.request", fake_request)
    monkeypatch.setattr(
        client,
        "_acquire_access_token",
        lambda: {"access_token": "new-token"},
    )

    response = client._request(
        method="GET",
        url="https://graph.microsoft.us/v1.0/me",
        operation="test operation",
        timeout=60,
    )

    assert response.status_code == 200
    assert len(calls) == 2
    assert calls[0]["headers"]["Authorization"] == "Bearer old-token"
    assert calls[1]["headers"]["Authorization"] == "Bearer new-token"


def test_request_does_not_retry_for_non_token_401(monkeypatch: pytest.MonkeyPatch) -> None:
    client = GraphClient.__new__(GraphClient)
    client._headers = {"Authorization": "Bearer old-token"}

    calls: list[dict] = []

    def fake_request(method, url, headers, timeout, **kwargs):  # noqa: ANN001
        calls.append({"method": method, "url": url, "headers": headers, "timeout": timeout})
        return _DummyResponse(
            ok=False,
            status_code=401,
            payload={"error": {"code": "accessDenied"}},
        )

    monkeypatch.setattr("watermark_app.graph.requests.request", fake_request)
    monkeypatch.setattr(
        client,
        "_acquire_access_token",
        lambda: {"access_token": "new-token"},
    )

    response = client._request(
        method="GET",
        url="https://graph.microsoft.us/v1.0/me",
        operation="test operation",
        timeout=60,
    )

    assert response.status_code == 401
    assert len(calls) == 1


def test_iter_changed_files_uses_delta_pagination(monkeypatch: pytest.MonkeyPatch) -> None:
    client = GraphClient.__new__(GraphClient)
    client.config = type("Cfg", (), {"graph_base_url": "https://graph.microsoft.us/v1.0"})()

    calls: list[str] = []

    def fake_request(method, url, operation, timeout):  # noqa: ANN001
        assert method == "GET"
        assert operation == "list changed drive items"
        calls.append(url)
        if len(calls) == 1:
            return _DummyResponse(
                ok=True,
                status_code=200,
                payload={
                    "value": [{"id": "f1", "file": {"mimeType": "x"}}],
                    "@odata.nextLink": "https://graph.microsoft.us/v1.0/drives/delta-next",
                },
            )
        return _DummyResponse(
            ok=True,
            status_code=200,
            payload={
                "value": [
                    {"id": "f2", "file": {"mimeType": "x"}},
                    {"id": "d1", "deleted": {}},
                ],
                "@odata.deltaLink": "https://graph.microsoft.us/v1.0/drives/delta-final",
            },
        )

    monkeypatch.setattr(client, "_request", fake_request)
    monkeypatch.setattr(client, "_raise_for_error", lambda _response, _operation: None)

    items, delta_link = client.iter_changed_files("drive-id")

    assert calls[0] == "https://graph.microsoft.us/v1.0/drives/drive-id/root/delta"
    assert calls[1] == "https://graph.microsoft.us/v1.0/drives/delta-next"
    assert [item["id"] for item in items] == ["f1", "f2"]
    assert delta_link == "https://graph.microsoft.us/v1.0/drives/delta-final"


def test_list_library_fields_calls_expected_endpoint(monkeypatch: pytest.MonkeyPatch) -> None:
    client = GraphClient.__new__(GraphClient)
    client.config = type("Cfg", (), {"graph_base_url": "https://graph.microsoft.us/v1.0"})()
    seen = {"url": None}

    def fake_request(method, url, operation, timeout):  # noqa: ANN001
        seen["url"] = url
        assert method == "GET"
        assert operation == "list library fields"
        assert timeout == 60
        return _DummyResponse(
            ok=True,
            status_code=200,
            payload={"value": [{"name": "RecordStatus", "displayName": "Record Status"}]},
        )

    monkeypatch.setattr(client, "_request", fake_request)
    monkeypatch.setattr(client, "_raise_for_error", lambda _response, _operation: None)

    fields = client.list_library_fields("drive-id")

    assert seen["url"] == "https://graph.microsoft.us/v1.0/drives/drive-id/list/columns"
    assert fields == [{"name": "RecordStatus", "displayName": "Record Status"}]
