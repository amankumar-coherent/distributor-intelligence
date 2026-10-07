"""Automatic session recovery when AI Mode keeps failing.

Requested after repeated manual restarts: "when many time come something
went wrong, ai cant generate then reset profile and session and run again
automatically".

Before this, the only auto-recovery was for a CRASHED browser
(_looks_like_dead_browser). A session that was alive but persistently
failing — every query CAPTCHA'd, refused, timing out, or returning an
empty answer pane — needed the user to notice and restart by hand.

The properties pinned here:
  1. A streak of failed queries rebuilds the session (browser + profile),
     because Google's flag rides on the profile's cookies.
  2. A successful answer clears the streak, so intermittent failures never
     accumulate into a needless reset.
  3. Resets are capped, so a broken environment cannot loop forever.
  4. Installed extensions survive a reset.
"""
from __future__ import annotations

from unittest.mock import patch

import pytest

from vendor_intel.scraping import google_ai_mode as aim


@pytest.fixture(autouse=True)
def _clean_counters(monkeypatch):
    """Every test starts from a healthy session."""
    monkeypatch.setattr(aim, "_consecutive_failures", 0)
    monkeypatch.setattr(aim, "_session_resets", 0)
    monkeypatch.setattr(aim, "_consecutive_captchas", 0)
    monkeypatch.setattr(aim, "_last_captcha_at", 0.0)
    monkeypatch.setattr(aim, "_last_refusal_at", 0.0)
    monkeypatch.setattr(aim.time, "sleep", lambda _s: None)


# --- configuration --------------------------------------------------------


def test_failure_threshold_and_reset_cap_are_sane():
    """Bounds, not preferences.

    The reset ceiling was 10 when a reset meant "the session looks dead".
    Verification now resets on every refusal and every rate-limit page,
    because both are session state that a browser restart cannot clear --
    a six-market pass is ~200 chunks, so a budget of 3 was spent long
    before the run finished and every later chunk was dropped. The
    workload legitimately needs tens of resets, so the upper bound guards
    against a typo (4000) rather than against a large-but-deliberate
    value.
    """
    assert 1 <= aim.max_consecutive_failures() <= 20
    assert 1 <= aim.max_session_resets() <= 100


def test_thresholds_are_env_overridable(monkeypatch):
    """Tunable from .env without a code change — and read at CALL time, so
    a value applied after import (which is when .env loads) still counts."""
    monkeypatch.setenv("GOOGLE_AI_MODE_MAX_FAILURES", "2")
    monkeypatch.setenv("GOOGLE_AI_MODE_MAX_RESETS", "1")
    assert aim.max_consecutive_failures() == 2
    assert aim.max_session_resets() == 1


# --- reset_session --------------------------------------------------------


def test_reset_session_closes_browser_and_rebuilds_profile(monkeypatch, tmp_path):
    calls = []
    monkeypatch.setattr(aim, "shutdown", lambda: calls.append("shutdown"))
    monkeypatch.setattr(aim, "active_profile_dir", lambda: str(tmp_path / "profile"))
    monkeypatch.setattr(
        aim, "_rebuild_profile_preserving_staged", lambda p: calls.append("rebuild")
    )
    monkeypatch.setenv("GOOGLE_AI_MODE_RESET_COOLOFF", "0")

    assert aim.reset_session("test") is True
    assert calls == ["shutdown", "rebuild"], calls
    assert aim._session_resets == 1


def test_reset_session_clears_every_failure_signal(monkeypatch, tmp_path):
    """The fresh session must be judged on its own results, not inherit the
    old session's CAPTCHA/refusal history."""
    monkeypatch.setattr(aim, "shutdown", lambda: None)
    monkeypatch.setattr(aim, "active_profile_dir", lambda: str(tmp_path / "p"))
    monkeypatch.setattr(aim, "_rebuild_profile_preserving_staged", lambda p: None)
    monkeypatch.setenv("GOOGLE_AI_MODE_RESET_COOLOFF", "0")
    monkeypatch.setattr(aim, "_consecutive_failures", 9)
    monkeypatch.setattr(aim, "_consecutive_captchas", 3)
    monkeypatch.setattr(aim, "_last_captcha_at", 12345.0)

    aim.reset_session("test")

    assert aim._consecutive_failures == 0
    assert aim._consecutive_captchas == 0
    assert aim.saw_captcha_recently(600.0) is False


def test_reset_session_is_capped(monkeypatch, tmp_path):
    """A permanently broken environment (no network, Chrome missing) must
    not loop forever rebuilding profiles."""
    monkeypatch.setattr(aim, "shutdown", lambda: None)
    monkeypatch.setattr(aim, "active_profile_dir", lambda: str(tmp_path / "p"))
    monkeypatch.setattr(aim, "_rebuild_profile_preserving_staged", lambda p: None)
    monkeypatch.setenv("GOOGLE_AI_MODE_RESET_COOLOFF", "0")
    monkeypatch.setattr(aim, "_session_resets", aim.max_session_resets())

    assert aim.reset_session("still broken") is False


def test_reset_session_survives_a_locked_profile(monkeypatch, tmp_path):
    """A profile dir Windows refuses to delete must not abort the run — the
    relaunch alone still clears in-memory state."""
    monkeypatch.setattr(aim, "shutdown", lambda: None)
    monkeypatch.setattr(aim, "active_profile_dir", lambda: str(tmp_path / "p"))
    monkeypatch.setenv("GOOGLE_AI_MODE_RESET_COOLOFF", "0")

    def _boom(_p):
        raise PermissionError("profile in use")

    monkeypatch.setattr(aim, "_rebuild_profile_preserving_staged", _boom)

    assert aim.reset_session("test") is True   # did not raise


def test_reset_preserves_installed_extensions(monkeypatch, tmp_path):
    """Extensions live in a SIBLING dir precisely so a reset cannot cost
    the user their solver."""
    monkeypatch.setenv("GOOGLE_AI_MODE_PROFILE_DIR", str(tmp_path / "profile"))
    src = tmp_path / "dl" / "solver"
    src.mkdir(parents=True)
    (src / "manifest.json").write_text('{"name":"solver","version":"1"}', encoding="utf-8")
    assert aim.install_extension(str(src)) is True

    (tmp_path / "profile" / "Default").mkdir(parents=True, exist_ok=True)
    aim._rebuild_profile_preserving_staged(tmp_path / "profile")

    assert aim.staged_extension_paths(), "a reset must not delete extensions"


# --- ask() integration ----------------------------------------------------


def _stub_ask_env(monkeypatch, answers):
    """Drive ask() through a fake page that yields `answers` in order."""
    seq = list(answers)

    class _Page:
        url = "https://www.google.com/search?udm=50"

        def goto(self, *_a, **_k):
            pass

        def wait_for_timeout(self, _ms):
            pass

    monkeypatch.setattr(aim, "_get_page", lambda: _Page())
    monkeypatch.setattr(aim, "_pace_query", lambda: None)
    monkeypatch.setattr(aim, "_wait_for_response", lambda _p: None)
    monkeypatch.setattr(aim, "_page_is_captcha", lambda _p: False)
    monkeypatch.setattr(
        aim, "_extract_answer_text", lambda _p: seq.pop(0) if seq else ""
    )


def test_streak_of_empty_answers_triggers_a_reset(monkeypatch):
    """The reported symptom: AI Mode keeps returning nothing usable."""
    _stub_ask_env(monkeypatch, [])          # always empty
    monkeypatch.setattr(aim, "max_consecutive_failures", lambda: 3)
    resets = []
    monkeypatch.setattr(
        aim, "reset_session", lambda reason="": (resets.append(reason), False)[1]
    )

    for _ in range(3):
        aim.ask("prompt", retries=0)

    assert resets, "a failure streak must rebuild the session"
    assert "consecutive" in resets[0]


def test_successful_answer_clears_the_failure_streak(monkeypatch):
    """Intermittent failures must not accumulate into a needless reset."""
    _stub_ask_env(monkeypatch, ["", "", "REAL ANSWER"])
    monkeypatch.setattr(aim, "max_consecutive_failures", lambda: 5)

    aim.ask("p", retries=0)
    aim.ask("p", retries=0)
    assert aim._consecutive_failures == 2

    assert aim.ask("p", retries=0) == "REAL ANSWER"
    assert aim._consecutive_failures == 0, "a good answer must reset the streak"


def test_reset_retries_the_prompt_that_tripped_it(monkeypatch):
    """The caller should get a real answer, not an empty result from
    whichever query happened to be last in the bad streak."""
    _stub_ask_env(monkeypatch, [])
    monkeypatch.setattr(aim, "max_consecutive_failures", lambda: 1)

    def _fake_reset(reason=""):
        # After "resetting", answers start working again.
        monkeypatch.setattr(aim, "_extract_answer_text", lambda _p: "AFTER RESET")
        return True

    monkeypatch.setattr(aim, "reset_session", _fake_reset)

    assert aim.ask("prompt", retries=0) == "AFTER RESET"


def test_blocked_session_can_still_reset(monkeypatch):
    """Regression: the consecutive-CAPTCHA circuit breaker returned ""
    BEFORE the retry loop, so a fully flagged profile — the case a reset
    actually fixes — could never trigger auto-recovery."""
    monkeypatch.setattr(aim, "_consecutive_captchas", aim._MAX_CONSECUTIVE_CAPTCHAS)
    called = []

    def _fake_reset(reason=""):
        called.append(reason)
        return False        # cap reached; ask() then returns ""

    monkeypatch.setattr(aim, "reset_session", _fake_reset)

    assert aim.ask("prompt") == ""
    assert called, "a flagged profile must attempt a session reset"
    assert "CAPTCHA" in called[0]


def test_reset_recursion_is_bounded(monkeypatch):
    """The post-reset retry uses retries=0 and must not recurse into
    another reset, or a broken environment would stack frames."""
    _stub_ask_env(monkeypatch, [])
    monkeypatch.setattr(aim, "max_consecutive_failures", lambda: 1)
    depth = {"n": 0}

    def _fake_reset(reason=""):
        depth["n"] += 1
        return depth["n"] <= 1       # allow one reset, then refuse

    monkeypatch.setattr(aim, "reset_session", _fake_reset)

    assert aim.ask("prompt", retries=0) == ""
    assert depth["n"] <= 2, f"reset called {depth['n']} times — unbounded"


def test_session_stats_reports_counters(monkeypatch):
    monkeypatch.setattr(aim, "_consecutive_failures", 2)
    monkeypatch.setattr(aim, "_session_resets", 1)
    s = aim.session_stats()
    assert s["consecutive_failures"] == 2
    assert s["session_resets"] == 1
    assert s["max_failures"] == aim.max_consecutive_failures()


# --- "Something went wrong, and an AI response wasn't generated" ----------
#
# Reported live with a screenshot. Google served the page normally; the
# generation itself failed. Distinct from a content refusal, and distinct
# from a CAPTCHA.


def test_google_side_error_is_detected():
    """The exact on-screen text must be recognised, or it flows downstream
    as a "real answer" and fails JSON parsing on every single query."""
    assert aim.is_ai_mode_refusal(
        "Something went wrong, and an AI response wasn't generated."
    )


def test_google_side_error_detected_after_a_long_echoed_prompt():
    """The page echoes our prompt first, so a 5 KB prompt plus a one-line
    failure message must still be classified by the TAIL of the page."""
    page_text = ("You are a precise market-research analyst. " * 120 +
                 "\nSomething went wrong, and an AI response wasn\'t generated.")
    assert len(page_text) > 600
    assert aim.is_ai_mode_refusal(page_text)


def test_repeated_google_side_errors_trigger_a_session_reset(monkeypatch):
    """End-to-end for the reported symptom: N of these in a row must
    rebuild the session without the user intervening."""
    _stub_ask_env(monkeypatch, [])
    monkeypatch.setattr(
        aim, "_extract_answer_text",
        lambda _p: "Something went wrong, and an AI response wasn't generated.",
    )
    monkeypatch.setattr(aim, "max_consecutive_failures", lambda: 3)
    monkeypatch.setenv("GOOGLE_AI_MODE_TRANSIENT_BACKOFF", "0")
    resets = []
    monkeypatch.setattr(
        aim, "reset_session", lambda reason="": (resets.append(reason), False)[1]
    )

    for _ in range(3):
        aim.ask("prompt", retries=0)

    assert resets, "repeated Google-side failures must reset the session"


def test_transient_error_backs_off_before_retrying(monkeypatch):
    """An immediate retry usually returns the identical error, so a short
    wait converts a wasted attempt into a likely success."""
    _stub_ask_env(monkeypatch, [])
    monkeypatch.setattr(
        aim, "_extract_answer_text",
        lambda _p: "Something went wrong, and an AI response wasn't generated.",
    )
    monkeypatch.setenv("GOOGLE_AI_MODE_TRANSIENT_BACKOFF", "20")
    slept = []
    monkeypatch.setattr(aim.time, "sleep", lambda s: slept.append(s))

    aim.ask("prompt", retries=1)

    assert any(x == 20 for x in slept), f"expected a 20s backoff, slept {slept}"


def test_content_refusal_does_not_back_off(monkeypatch):
    """Only the Google-side error is transient. A content refusal should be
    re-asked promptly — waiting cannot change the model's mind."""
    _stub_ask_env(monkeypatch, [])
    monkeypatch.setattr(
        aim, "_extract_answer_text",
        lambda _p: "It looks like there's no response available for this search.",
    )
    monkeypatch.setenv("GOOGLE_AI_MODE_TRANSIENT_BACKOFF", "20")
    slept = []
    monkeypatch.setattr(aim.time, "sleep", lambda s: slept.append(s))

    aim.ask("prompt", retries=1)

    assert 20 not in slept, "a content refusal must not pay the transient backoff"


def test_reset_attempts_stop_once_the_cap_is_reached(monkeypatch):
    """Regression: the streak counter kept climbing past the threshold
    (6/5, 7/5, ...), so every later query re-attempted a reset that was
    already refused — pure log noise and wasted time."""
    _stub_ask_env(monkeypatch, [])
    monkeypatch.setattr(aim, "max_consecutive_failures", lambda: 2)
    attempts = []
    monkeypatch.setattr(
        aim, "reset_session", lambda reason="": (attempts.append(reason), False)[1]
    )

    for _ in range(10):
        aim.ask("prompt", retries=0)

    # 10 queries at a threshold of 2 would be 9 attempts if the streak were
    # never cleared; clearing it caps attempts at one per full streak.
    assert len(attempts) <= 5, f"{len(attempts)} reset attempts — streak not cleared"


# --- persistent Google-side failure must reset the session ---------------
#
# Observed live: 8 consecutive "Something went wrong, and an AI response
# wasn't generated" pages, each retried after a 20s backoff, and ZERO
# AUTO-RECOVERY lines. Cause: a refusal loops inside ask() via `continue`,
# so it rarely reaches the per-call failure counter -- a wedged session
# retried indefinitely instead of being rebuilt.

TRANSIENT = "Something went wrong, and an AI response wasn't generated."


def test_transient_streak_triggers_a_session_reset(monkeypatch):
    _stub_ask_env(monkeypatch, [])
    monkeypatch.setattr(aim, "_extract_answer_text", lambda _p: TRANSIENT)
    monkeypatch.setattr(aim, "max_consecutive_transient", lambda: 3)
    monkeypatch.setenv("GOOGLE_AI_MODE_TRANSIENT_BACKOFF", "0")
    monkeypatch.setattr(aim, "_consecutive_transient", 0)

    resets = []
    monkeypatch.setattr(
        aim, "reset_session", lambda reason="": (resets.append(reason), False)[1]
    )

    # The streak persists ACROSS calls, which is how it behaves in a real
    # run: batch after batch each getting the same Google-side failure.
    for _ in range(3):
        aim.ask("prompt", retries=0)

    assert resets, "a wedged session must be rebuilt, not retried forever"
    assert "generation failures" in resets[0]


def test_transient_backoff_applies_before_the_reset_threshold(monkeypatch):
    """The first couple of failures are genuinely transient, so back off
    and retry rather than rebuilding the session immediately."""
    _stub_ask_env(monkeypatch, [])
    monkeypatch.setattr(aim, "_extract_answer_text", lambda _p: TRANSIENT)
    monkeypatch.setattr(aim, "max_consecutive_transient", lambda: 99)
    monkeypatch.setenv("GOOGLE_AI_MODE_TRANSIENT_BACKOFF", "20")
    monkeypatch.setattr(aim, "_consecutive_transient", 0)
    slept = []
    monkeypatch.setattr(aim.time, "sleep", lambda s: slept.append(s))

    aim.ask("prompt", retries=1)

    assert 20 in slept


def test_a_real_answer_clears_the_transient_streak(monkeypatch):
    """Intermittent Google errors must not accumulate toward a reset."""
    seq = [TRANSIENT, "REAL ANSWER"]
    _stub_ask_env(monkeypatch, [])
    monkeypatch.setattr(aim, "_extract_answer_text", lambda _p: seq.pop(0) if seq else "")
    monkeypatch.setenv("GOOGLE_AI_MODE_TRANSIENT_BACKOFF", "0")
    monkeypatch.setattr(aim, "_consecutive_transient", 0)

    assert aim.ask("prompt", retries=2) == "REAL ANSWER"
    assert aim._consecutive_transient == 0


def test_transient_threshold_is_env_tunable(monkeypatch):
    monkeypatch.setenv("GOOGLE_AI_MODE_MAX_TRANSIENT", "5")
    assert aim.max_consecutive_transient() == 5
    monkeypatch.delenv("GOOGLE_AI_MODE_MAX_TRANSIENT")
    assert aim.max_consecutive_transient() == aim._DEFAULT_MAX_TRANSIENT


def test_reset_clears_the_transient_streak(monkeypatch, tmp_path):
    monkeypatch.setattr(aim, "shutdown", lambda: None)
    monkeypatch.setattr(aim, "active_profile_dir", lambda: str(tmp_path / "p"))
    monkeypatch.setattr(aim, "_rebuild_profile_preserving_staged", lambda p: None)
    monkeypatch.setenv("GOOGLE_AI_MODE_RESET_COOLOFF", "0")
    monkeypatch.setattr(aim, "_consecutive_transient", 9)

    aim.reset_session("test")

    assert aim._consecutive_transient == 0


def test_google_page_chrome_alone_is_not_a_transient_error():
    """The log used to print the first 60 chars, which is Google's nav
    boilerplate -- that hid the real message and made a genuine failure
    look unexplained. The nav text on its own is not an error."""
    chrome_only = ("Skip to main content\nAccessibility help\nAI Mode\nAll\n"
                   "Images\nVideos\nNews\nMore\nSign in")
    assert not aim.is_ai_mode_refusal(chrome_only)


def test_error_is_detected_behind_the_page_chrome():
    """The real page is chrome + echoed prompt + the error, so detection
    must look at the TAIL."""
    full = ("Skip to main content\nAccessibility help\nAI Mode\nAll\nImages\n"
            "Videos\nNews\nMore\nSign in\nAI Mode conversation: test\ntest\n"
            + TRANSIENT + "\nAI Mode response is ready")
    assert aim.is_ai_mode_refusal(full)
