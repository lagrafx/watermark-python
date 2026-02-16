from watermark_app.graph import GraphClientError
from watermark_app.graph import GraphClient


class _DummyResponse:
    def __init__(self, ok: bool, status_code: int, payload: dict | None = None, text: str = "") -> None:
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
