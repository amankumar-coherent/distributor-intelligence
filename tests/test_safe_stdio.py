"""safe_stdio must swallow broken console handles (Windows OSError 22)."""
from __future__ import annotations

import io
import sys

from vendor_intel.utils.safe_stdio import SafeTextIO, install_safe_stdio


class _BrokenStream(io.StringIO):
    def write(self, s: str) -> int:  # type: ignore[override]
        raise OSError(22, "Incorrect function", None, 1, None)

    def flush(self) -> None:
        raise OSError(22, "Incorrect function", None, 1, None)


def test_safe_writer_swallows_oserror_22() -> None:
    safe = SafeTextIO(_BrokenStream())
    assert safe.write("hello") == 5
    safe.flush()  # must not raise
    safe.writelines(["a", "b"])


def test_install_safe_stdio_idempotent(monkeypatch) -> None:
    buf = io.StringIO()
    monkeypatch.setattr(sys, "stdout", buf)
    monkeypatch.setattr(sys, "stderr", io.StringIO())
    assert install_safe_stdio() is True
    assert isinstance(sys.stdout, SafeTextIO)
    assert install_safe_stdio() is False
    print("ok", flush=True)
    assert "ok" in buf.getvalue()
