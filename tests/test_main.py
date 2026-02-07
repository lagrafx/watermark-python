from datetime import datetime, timezone

from watermark_app.state import should_process
from watermark_app.watermarking import is_supported_extension


def test_supported_extensions() -> None:
    assert is_supported_extension("a.docx")
    assert is_supported_extension("a.xlsm")
    assert not is_supported_extension("a.pdf")


def test_should_process_with_no_state() -> None:
    assert should_process("2026-02-07T00:00:00Z", None)


def test_should_process_with_state() -> None:
    last_run = datetime(2026, 2, 7, 0, 0, tzinfo=timezone.utc)
    assert should_process("2026-02-07T00:01:00Z", last_run)
    assert not should_process("2026-02-06T23:59:00Z", last_run)
