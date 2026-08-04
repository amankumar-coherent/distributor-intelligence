"""Prevent Windows Cursor/console pipe crashes from killing the batch process.

On Windows under heavy concurrent ``print(..., flush=True)`` into a Cursor /
VS Code terminal, the console handle can become invalid. Python then raises:

    OSError(22, 'Incorrect function', ...)
    lost sys.stderr

and the whole ``run_batch`` process dies mid-market.

Wrapping stdout/stderr so write/flush never raise keeps the pipeline alive;
lines may be dropped if the terminal is broken, but progress still checkpoints.
"""
from __future__ import annotations

import sys
from typing import Any, TextIO


class SafeTextIO:
    """Proxy around a text stream that swallows broken-pipe / invalid-handle errors."""

    __slots__ = ("_stream",)

    def __init__(self, stream: TextIO) -> None:
        self._stream = stream

    def write(self, s: Any) -> int:
        try:
            n = self._stream.write(s)
            return int(n) if n is not None else (len(s) if isinstance(s, str) else 0)
        except (OSError, ValueError, AttributeError):
            return len(s) if isinstance(s, str) else 0

    def writelines(self, lines: Any) -> None:
        for line in lines:
            self.write(line)

    def flush(self) -> None:
        try:
            self._stream.flush()
        except (OSError, ValueError, AttributeError):
            pass

    def reconfigure(self, *args: Any, **kwargs: Any) -> None:
        fn = getattr(self._stream, "reconfigure", None)
        if callable(fn):
            try:
                fn(*args, **kwargs)
            except (OSError, ValueError, AttributeError):
                pass

    @property
    def encoding(self) -> str:
        return getattr(self._stream, "encoding", None) or "utf-8"

    @property
    def errors(self) -> str:
        return getattr(self._stream, "errors", None) or "replace"

    @property
    def buffer(self) -> Any:
        return getattr(self._stream, "buffer", None)

    def isatty(self) -> bool:
        try:
            return bool(self._stream.isatty())
        except (OSError, ValueError, AttributeError):
            return False

    def fileno(self) -> int:
        try:
            return int(self._stream.fileno())
        except (OSError, ValueError, AttributeError):
            return -1

    def __getattr__(self, name: str) -> Any:
        return getattr(self._stream, name)


def install_safe_stdio(*, force: bool = False) -> bool:
    """Wrap ``sys.stdout`` / ``sys.stderr`` if not already wrapped. Returns True if installed."""
    changed = False
    out = sys.stdout
    err = sys.stderr
    if force or not isinstance(out, SafeTextIO):
        if out is not None:
            sys.stdout = SafeTextIO(out)  # type: ignore[assignment]
            changed = True
    if force or not isinstance(err, SafeTextIO):
        if err is not None:
            sys.stderr = SafeTextIO(err)  # type: ignore[assignment]
            changed = True
    # Prefer replace over strict so emoji/dashes never raise UnicodeEncodeError
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")  # type: ignore[attr-defined]
        except Exception:
            pass
    return changed
