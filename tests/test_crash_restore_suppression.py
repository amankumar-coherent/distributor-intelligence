"""A restored-crash session showed a modal and eight blank tabs.

Live failure, reported with a screenshot: after AUTO-RECOVERY reset the
profile, the new window came up with eight about:blank tabs and a
"Chromium didn't shut down correctly / Restore pages?" bubble, and nothing
loaded. The modal matters because it sits above the page and swallows the
navigation, so the pipeline reads an empty answer pane, counts a failure,
and eventually burns a session reset on a problem one pref byte caused.

Three causes, all pinned here:

  1. _seed_profile_from_real_chrome copies Preferences from the user's
     REAL Chrome, which is normally RUNNING. A running Chrome's
     Preferences says exit_type="Crashed" and carries its open-tab list,
     so every fresh automation profile inherited a crash it never had.
  2. The flag is also set legitimately -- Ctrl+C, a sleeping machine, or
     reset_session() tearing the browser down mid-query -- so it must be
     cleared on every launch, not only when seeding.
  3. _get_page() returned pages[0]. On a restored session that is a stale
     about:blank that nothing navigates, so the caller waited on a blank
     page while any real answer loaded in a different tab.
"""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from vendor_intel.scraping import google_ai_mode as m


# --- clearing the crash flag ---------------------------------------------


def _profile(tmp_path: Path, **profile_prefs) -> Path:
    d = tmp_path / "prof"
    (d / "Default").mkdir(parents=True)
    (d / "Default" / "Preferences").write_text(
        json.dumps({"profile": profile_prefs or {"exit_type": "Crashed"}}),
        encoding="utf-8",
    )
    return d


def _prefs(profile: Path) -> dict:
    return json.loads(
        (profile / "Default" / "Preferences").read_text(encoding="utf-8"))


@pytest.mark.parametrize("bad", ["Crashed", "SessionEnded", "", "Killed"])
def test_any_non_normal_exit_type_is_normalised(tmp_path, bad):
    """Chrome offers restore for ANY value that is not "Normal"."""
    p = _profile(tmp_path, exit_type=bad)
    m._mark_profile_exited_cleanly(p)
    assert _prefs(p)["profile"]["exit_type"] == "Normal"


def test_exited_cleanly_is_set(tmp_path):
    p = _profile(tmp_path, exit_type="Crashed", exited_cleanly=False)
    m._mark_profile_exited_cleanly(p)
    assert _prefs(p)["profile"]["exited_cleanly"] is True


def test_copied_startup_tabs_are_dropped(tmp_path):
    """The eight about:blank tabs came from the real profile's session
    list, which restore replays even after the crash flag is cleared."""
    p = tmp_path / "prof"
    (p / "Default").mkdir(parents=True)
    (p / "Default" / "Preferences").write_text(
        json.dumps({
            "profile": {"exit_type": "Crashed"},
            "session": {"restore_on_startup": 1,
                        "startup_urls": ["https://a.example",
                                         "https://b.example"]},
        }),
        encoding="utf-8",
    )
    m._mark_profile_exited_cleanly(p)
    session = _prefs(p)["session"]
    assert session["startup_urls"] == []
    assert session["restore_on_startup"] == 5, "5 == New Tab Page"


def test_session_files_are_removed(tmp_path):
    """Chrome replays tabs from these files regardless of the pref."""
    p = _profile(tmp_path)
    d = p / "Default"
    (d / "Sessions").mkdir()
    (d / "Sessions" / "Session_1").write_text("x", encoding="utf-8")
    names = ("Current Session", "Current Tabs", "Last Session", "Last Tabs")
    for name in names:
        (d / name).write_text("x", encoding="utf-8")

    m._mark_profile_exited_cleanly(p)

    assert not (d / "Sessions").exists()
    for name in names:
        assert not (d / name).exists(), name


def test_other_preferences_survive(tmp_path):
    """Clearing the crash flag must not cost the seeded session -- the
    point of seeding is to arrive warm rather than CAPTCHA-prone."""
    p = tmp_path / "prof"
    (p / "Default").mkdir(parents=True)
    (p / "Default" / "Preferences").write_text(
        json.dumps({
            "profile": {"exit_type": "Crashed", "name": "Person 1"},
            "extensions": {"settings": {"abc": {"location": 4}}},
        }),
        encoding="utf-8",
    )
    m._mark_profile_exited_cleanly(p)
    data = _prefs(p)
    assert data["profile"]["name"] == "Person 1"
    assert data["extensions"]["settings"]["abc"]["location"] == 4


def test_missing_preferences_is_not_an_error(tmp_path):
    """A brand-new profile has no Preferences yet; launch must proceed."""
    d = tmp_path / "empty"
    d.mkdir()
    m._mark_profile_exited_cleanly(d)


def test_corrupt_preferences_is_not_an_error(tmp_path):
    p = tmp_path / "prof"
    (p / "Default").mkdir(parents=True)
    (p / "Default" / "Preferences").write_text("{not json", encoding="utf-8")
    m._mark_profile_exited_cleanly(p)


def test_seeding_clears_the_flag_it_just_copied(tmp_path, monkeypatch):
    """The originating bug: the source Chrome is running, so the
    Preferences copied in says "Crashed"."""
    src_root = tmp_path / "real"
    (src_root / "Default").mkdir(parents=True)
    (src_root / "Default" / "Preferences").write_text(
        json.dumps({"profile": {"exit_type": "Crashed"}}), encoding="utf-8")
    monkeypatch.setenv("CHATGPT_SCRAPER_CHROME_PROFILE_DIR", str(src_root))
    monkeypatch.setenv("CHATGPT_SCRAPER_CHROME_PROFILE_DIRECTORY", "Default")

    dest = tmp_path / "auto"
    dest.mkdir()
    m._seed_profile_from_real_chrome(dest)

    assert _prefs(dest)["profile"]["exit_type"] == "Normal"


# --- launch flags ---------------------------------------------------------


def test_launch_suppresses_the_restore_bubble():
    """Even with a clean pref, an unclean kill between the write and the
    launch can re-raise the modal, so the flags are belt-and-braces."""
    src = Path(m.__file__).read_text(encoding="utf-8")
    for flag in ("--hide-crash-restore-bubble",
                 "--disable-session-crashed-bubble"):
        assert flag in src, flag


def test_launch_skips_first_run_prompts():
    src = Path(m.__file__).read_text(encoding="utf-8")
    assert "--no-first-run" in src
    assert "--no-default-browser-check" in src


def test_existing_profile_is_cleaned_before_launch():
    """Not just on seeding -- a Ctrl+C'd run leaves the flag behind."""
    src = Path(m.__file__).read_text(encoding="utf-8")
    launch = src[src.index("def _get_page():"):]
    assert "_mark_profile_exited_cleanly(profile)" in launch


def test_reset_cleans_a_profile_it_could_not_delete():
    """rmtree fails when Chrome still holds a file open; the leftover
    Preferences would otherwise keep the crash flag."""
    src = Path(m.__file__).read_text(encoding="utf-8")
    fn = src[src.index("def _rebuild_profile_preserving_staged"):]
    fn = fn[:fn.index("\ndef ")]
    assert "_mark_profile_exited_cleanly" in fn


# --- picking the right tab -----------------------------------------------


class _FakePage:
    def __init__(self, url="about:blank"):
        self.url = url
        self.closed = False

    def is_closed(self):
        return self.closed

    def close(self):
        self.closed = True


class _FakeContext:
    def __init__(self, pages):
        self.pages = list(pages)
        self.created = 0

    def new_page(self):
        self.created += 1
        page = _FakePage()
        self.pages.append(page)
        return page


def test_extra_restored_tabs_are_closed():
    """The reported screenshot had eight of them."""
    ctx = _FakeContext([_FakePage() for _ in range(8)])
    kept = m._single_page(ctx)
    assert sum(1 for p in ctx.pages if not p.closed) == 1
    assert not kept.closed


def test_a_page_with_a_real_url_is_preferred():
    """pages[0] on a restored session is a stale blank tab that nothing
    navigates -- waiting on it hangs forever."""
    blank_a = _FakePage()
    real = _FakePage("https://www.google.com/search?udm=50")
    blank_b = _FakePage()
    kept = m._single_page(_FakeContext([blank_a, real, blank_b]))
    assert kept is real
    assert blank_a.closed and blank_b.closed


def test_a_single_page_is_returned_untouched():
    only = _FakePage("https://www.google.com/")
    ctx = _FakeContext([only])
    assert m._single_page(ctx) is only
    assert not only.closed
    assert ctx.created == 0


def test_a_new_page_is_created_when_there_are_none():
    ctx = _FakeContext([])
    page = m._single_page(ctx)
    assert page in ctx.pages
    assert ctx.created == 1


def test_already_closed_pages_are_ignored():
    dead = _FakePage()
    dead.closed = True
    live = _FakePage("https://www.google.com/")
    assert m._single_page(_FakeContext([dead, live])) is live


def test_all_pages_closed_creates_a_fresh_one():
    dead = _FakePage()
    dead.closed = True
    ctx = _FakeContext([dead])
    page = m._single_page(ctx)
    assert ctx.created == 1
    assert not page.closed
