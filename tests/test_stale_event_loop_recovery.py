"""One greenlet error permanently disabled AI Mode.

Live failure. Mid-run the log turned into an unbroken repetition of:

    query failed (Error: It looks like you are using Playwright Sync API
    inside the asyncio loop. Please use the Async API instead.)
    query failed (Error: BrowserType.launch_persistent_context: Opening in
    existing browser session ... the profile is already in use)

...until the failure streak hit 5/5, the session-reset budget was spent
("session already reset 3x"), and every chunk was dropped fail-closed.
402 companies were verified before it; the rest were lost.

Three separate defects had to line up, and each is pinned below.

1. Sync Patchright refuses to run when an event loop is installed on the
   calling thread -- even a CLOSED one. Something left a closed loop
   behind, so every later query failed for a reason no relaunch could fix.
   _clear_stale_event_loop removes a closed loop, and deliberately does
   NOT remove a live one (that error is legitimate).

2. Neither the greenlet error nor "profile is already in use" was in
   _DEAD_BROWSER_MARKERS, so the broken context was never torn down and
   was handed to the next query intact.

3. Teardown was gated on `attempt < retries`, so the final attempt left
   the dead context in place for the NEXT ask() call, and shutdown()
   returned before Chrome released the profile lock -- which is what
   produced the "already in use" half of the loop.
"""
from __future__ import annotations

import asyncio
import warnings
from pathlib import Path

import pytest

from vendor_intel.scraping import google_ai_mode as m


@pytest.fixture(autouse=True)
def _restore_loop():
    """Never leak a loop into another test."""
    yield
    try:
        asyncio.set_event_loop(None)
    except Exception:
        pass


# --- 1. the stale closed loop --------------------------------------------


def test_a_closed_loop_is_cleared():
    """The whole failure: a closed loop blocks sync Playwright forever."""
    loop = asyncio.new_event_loop()
    loop.close()
    asyncio.set_event_loop(loop)

    assert m._clear_stale_event_loop() is True
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        with pytest.raises(RuntimeError):
            asyncio.get_event_loop_policy().get_event_loop()


def test_an_open_loop_is_left_alone():
    """A LIVE loop means we really are being called from async code, where
    the sync API is the wrong tool. Silently clearing it would hide a real
    bug and could break the caller's own loop."""
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    try:
        assert m._clear_stale_event_loop() is False
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            assert asyncio.get_event_loop_policy().get_event_loop() is loop
    finally:
        loop.close()


def test_no_loop_is_a_no_op():
    asyncio.set_event_loop(None)
    assert m._clear_stale_event_loop() is False


def test_repeated_calls_are_stable():
    """Called once per query for thousands of queries."""
    for _ in range(5):
        loop = asyncio.new_event_loop()
        loop.close()
        asyncio.set_event_loop(loop)
        assert m._clear_stale_event_loop() is True
        assert m._clear_stale_event_loop() is False


def test_get_page_clears_before_touching_patchright():
    """Must run before the context-reuse path too: driving an existing
    context is blocked by a stale loop exactly as launching is. Placing
    the call after the `if _context is not None` early return would leave
    the common case broken."""
    src = Path(m.__file__).read_text(encoding="utf-8")
    fn = src[src.index("def _get_page():"):]
    clear_at = fn.index("_clear_stale_event_loop()")
    reuse_at = fn.index("if _context is not None")
    assert clear_at < reuse_at, "must clear before the reuse early-return"


def test_a_running_loop_is_detected_from_async_code():
    """get_running_loop is the only reliable "am I really inside async"
    check; the installed-loop lookup cannot tell."""
    async def check():
        return m._clear_stale_event_loop()

    assert asyncio.run(check()) is False


# --- 2. recognising the unusable-browser errors --------------------------


@pytest.mark.parametrize(
    "message",
    [
        # verbatim from the failing run
        "It looks like you are using Playwright Sync API inside the asyncio loop. "
        "Please use the Async API instead.",
        "BrowserType.launch_persistent_context: Opening in existing browser session. "
        "This usually means that the profile is already in use",
        "EPIPE: broken pipe, write",
        # previously recognised, must not regress
        "Target page, context or browser has been closed",
        "Target closed",
        "Page crashed",
    ],
)
def test_unusable_browser_errors_are_recognised(message):
    assert m._looks_like_dead_browser(Exception(message)), message


@pytest.mark.parametrize(
    "message",
    [
        "net::ERR_NAME_NOT_RESOLVED",
        "Timeout 30000ms exceeded waiting for selector",
        "429 Too Many Requests",
    ],
)
def test_ordinary_failures_do_not_trigger_a_relaunch(message):
    """A relaunch costs ~10s plus a fresh CAPTCHA risk, so a page-level
    problem must not provoke one."""
    assert not m._looks_like_dead_browser(Exception(message)), message


def test_detection_is_case_insensitive():
    assert m._looks_like_dead_browser(Exception("TARGET CLOSED"))
    assert m._looks_like_dead_browser(Exception("Epipe"))


# --- 3. teardown is not gated on retries left ----------------------------


def test_teardown_is_not_conditional_on_remaining_retries():
    """A broken context outlives the ask() call that saw the error, so
    skipping teardown on the last attempt handed the dead context to the
    next query -- exactly the observed unbroken failure run."""
    src = Path(m.__file__).read_text(encoding="utf-8")
    assert "_looks_like_dead_browser(exc) and attempt < retries" not in src, (
        "teardown must not be gated on retries remaining"
    )
    assert "if _looks_like_dead_browser(exc):" in src


# --- shutdown must release the profile lock ------------------------------


def test_shutdown_waits_for_the_lock():
    """close() returns before Chrome exits; relaunching in that window is
    what produced "the profile is already in use"."""
    src = Path(m.__file__).read_text(encoding="utf-8")
    fn = src[src.index("def shutdown()"):]
    fn = fn[:fn.index("\ndef ")] if "\ndef " in fn else fn
    assert "_wait_for_profile_lock_release" in fn


def test_lock_wait_returns_immediately_when_free(tmp_path):
    import time

    started = time.monotonic()
    assert m._wait_for_profile_lock_release(tmp_path, timeout=5.0) is True
    assert time.monotonic() - started < 1.0, "must not sleep when already free"


def test_lock_wait_gives_up_rather_than_hanging(tmp_path):
    """A Windows profile sometimes keeps SingletonLock after exit, so a
    still-present lock must not block the run forever."""
    (tmp_path / "SingletonLock").write_text("held", encoding="utf-8")
    assert m._wait_for_profile_lock_release(tmp_path, timeout=0.5) is False


def test_lock_wait_notices_a_release(tmp_path):
    """The lock disappearing is the signal Chrome has let go."""
    import threading

    lock = tmp_path / "SingletonLock"
    lock.write_text("held", encoding="utf-8")
    threading.Timer(0.3, lock.unlink).start()
    assert m._wait_for_profile_lock_release(tmp_path, timeout=5.0) is True


# --- the diagnostic ------------------------------------------------------


def test_diagnostic_names_a_stale_loop():
    """The greenlet error alone is unactionable -- it never says which
    loop or who installed it."""
    loop = asyncio.new_event_loop()
    loop.close()
    asyncio.set_event_loop(loop)
    out = m._asyncio_diagnostic()
    assert "closed=True" in out
    assert "thread=" in out


def test_diagnostic_handles_no_loop():
    asyncio.set_event_loop(None)
    assert "no installed loop" in m._asyncio_diagnostic()


def test_diagnostic_never_raises():
    """It runs inside an exception handler; raising there would mask the
    original error."""
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    try:
        assert isinstance(m._asyncio_diagnostic(), str)
    finally:
        loop.close()
