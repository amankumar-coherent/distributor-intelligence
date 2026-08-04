"""Terminal + file logging for avocado scrape / re-enrich runs."""
from __future__ import annotations

import asyncio
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import TextIO

_log_lock = asyncio.Lock()
_print_lock: asyncio.Lock | None = None
_log_file: TextIO | None = None
_progress_path: Path | None = None
_run_t0: float | None = None


class _TeeStream:
    """Mirror writes to terminal and a log file."""

    def __init__(self, stream: TextIO, log_path: Path) -> None:
        self._stream = stream
        log_path.parent.mkdir(parents=True, exist_ok=True)
        self._file = log_path.open("a", encoding="utf-8")

    def write(self, data: str) -> int:
        if not data:
            return 0
        self._stream.write(data)
        self._file.write(data)
        self._file.flush()
        if hasattr(self._stream, "flush"):
            self._stream.flush()
        return len(data)

    def flush(self) -> None:
        self._stream.flush()
        self._file.flush()

    def isatty(self) -> bool:
        return bool(getattr(self._stream, "isatty", lambda: False)())

    def fileno(self) -> int:
        return self._stream.fileno()

    @property
    def encoding(self) -> str:
        return getattr(self._stream, "encoding", "utf-8")

    def reconfigure(self, **kwargs: object) -> None:
        if hasattr(self._stream, "reconfigure"):
            self._stream.reconfigure(**kwargs)  # type: ignore[attr-defined]


def setup_scrape_logging(
    log_path: Path,
    *,
    progress_path: Path | None = None,
) -> Path:
    """Tee stdout/stderr to log_path; return resolved log path."""
    global _log_file, _progress_path, _run_t0, _print_lock
    _run_t0 = time.perf_counter()
    _progress_path = progress_path
    resolved = log_path.resolve()
    resolved.parent.mkdir(parents=True, exist_ok=True)
    banner = (
        f"\n{'=' * 72}\n"
        f"  SCRAPE RUN START  {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S UTC')}\n"
        f"  Log file: {resolved}\n"
    )
    if progress_path:
        banner += f"  Progress: {progress_path.resolve()}\n"
        progress_path.parent.mkdir(parents=True, exist_ok=True)
        progress_path.write_text("", encoding="utf-8")
    banner += f"{'=' * 72}\n"
    with resolved.open("a", encoding="utf-8") as f:
        f.write(banner)
    sys.stdout.write(banner)
    sys.stdout.flush()
    sys.stdout = _TeeStream(sys.stdout, resolved)  # type: ignore[assignment]
    sys.stderr = _TeeStream(sys.stderr, resolved)  # type: ignore[assignment]
    _log_file = resolved.open("a", encoding="utf-8")
    _print_lock = asyncio.Lock()
    return resolved


def log(msg: str, *, level: str = "INFO") -> None:
    """Timestamped line to terminal + tee log."""
    ts = datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] [{level}] {msg}"
    print(line, flush=True)


async def alog(msg: str, *, level: str = "INFO") -> None:
    """Thread-safe timestamped log (use after concurrent crawls interleave)."""
    if _print_lock is None:
        log(msg, level=level)
        return
    async with _print_lock:
        log(msg, level=level)


def log_section(title: str) -> None:
    log(f"{'—' * 20} {title} {'—' * 20}")


def append_progress(record: dict) -> None:
    """Append one JSON line per completed company (safe to tail while running)."""
    if not _progress_path:
        return
    import json

    _progress_path.parent.mkdir(parents=True, exist_ok=True)
    with _progress_path.open("a", encoding="utf-8") as f:
        f.write(json.dumps(record, default=str) + "\n")


def elapsed_minutes() -> float:
    if _run_t0 is None:
        return 0.0
    return (time.perf_counter() - _run_t0) / 60.0
