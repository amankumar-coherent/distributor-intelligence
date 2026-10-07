"""Two pipelines on one browser profile is never valid.

Live failure. Two copies of the watchdog script started at the same second
and produced two run_all_markets trees, two discovery processes and two
verifiers, all launching Chrome against ONE persistent profile. The log
became an endless repetition of:

    BrowserType.launch_persistent_context: Opening in existing browser
    session. This usually means that the profile is already in use

and the session-reset budget was spent inside a minute, because no amount
of relaunching can win a lock that another LIVE process legitimately
holds. The existing shutdown() lock-wait could not help: the lock was not
stale.

The markets share one profile and one IP by design -- that is why
run_all_markets runs them strictly one at a time. So concurrency here is
always a bug, and detecting it at startup turns an unbreakable loop into
one clear error.

The guard is deliberately ADVISORY rather than a hard mutex: it refuses
only when the recorded PID is still alive. A crashed run leaves a stale
record, which is taken over rather than treated as fatal -- everything in
this pipeline is built to resume, and refusing to start after a crash
would be a worse failure than the one being fixed.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest

from vendor_intel.scraping import google_ai_mode as m


@pytest.fixture
def profile(tmp_path: Path) -> Path:
    p = tmp_path / "ai_mode_chrome_profile"
    p.mkdir()
    yield p
    m._release_profile(p)


# --- claiming -------------------------------------------------------------


def test_a_free_profile_can_be_claimed(profile):
    m._claim_profile(profile)
    assert m._owner_file(profile).read_text(encoding="utf-8").strip() == str(os.getpid())


def test_reclaiming_our_own_profile_is_fine(profile):
    """_get_page() can be called many times in one run."""
    m._claim_profile(profile)
    m._claim_profile(profile)      # must not raise
    m._claim_profile(profile)


def test_a_live_peer_warns_but_never_blocks(profile, capsys):
    """Warn, never raise.

    This check originally raised ProfileInUse, and the cure proved worse
    than the disease: it refused the verifier subprocess because the
    verifier's own PARENT held the claim, producing 159 failures, every
    chunk dropped fail-closed, and an entire Solar Rooftop verification
    pass lost.

    Duplicate runs are prevented at the LAUNCHER (the PID lock in
    run_four_markets.ps1), where the duplication actually happens. Down
    here the useful contribution is naming the other process -- and a
    warning does that without any way to take down a run that would
    otherwise have worked.
    """
    import subprocess

    peer = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(30)"])
    try:
        m._owner_file(profile).write_text(str(peer.pid), encoding="utf-8")
        m._claim_profile(profile)          # must NOT raise
        out = capsys.readouterr().out
        assert str(peer.pid) in out, "the warning must name the other process"
        assert "Continuing anyway" in out
    finally:
        peer.kill()
        peer.wait(timeout=10)


def test_nothing_in_the_claim_path_can_raise(profile):
    """The property that matters after two self-inflicted outages: no
    input to _claim_profile may abort a run."""
    import subprocess

    peer = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(30)"])
    try:
        for value in (str(peer.pid), str(os.getppid()), "999999",
                      "not-a-pid", "", "-1", "0"):
            m._owner_file(profile).write_text(value, encoding="utf-8")
            m._claim_profile(profile)      # must never raise
    finally:
        peer.kill()
        peer.wait(timeout=10)


def test_an_ancestor_is_not_a_competitor(profile):
    """Self-inflicted deadlock, seen live.

    run_market_research runs the verifier as a SUBPROCESS -- deliberately,
    so the verifier drives its own Patchright session rather than
    contending for the parent's browser handle. The parent stays alive the
    whole time, blocked in subprocess.run(), still holding the claim it
    took during discovery.

    The guard refused the child with "process 32508 is already driving
    ai_mode_chrome_profile_chromium" -- 32508 being the child's own
    parent. Nothing was competing; the two run strictly in sequence. It
    burned the whole reset budget (AUTO-RECOVERY 1/3, 2/3, 3/3) without a
    single successful query: precisely the failure the guard exists to
    prevent, caused by the guard."""
    m._owner_file(profile).write_text(str(os.getppid()), encoding="utf-8")
    m._claim_profile(profile)      # must NOT raise


def test_the_ancestor_walk_finds_the_real_parent():
    """The first implementation shelled out to `wmic`, which is
    deprecated and absent from Windows 11, so it silently returned only
    our own PID and left the deadlock in place. Anything that reports no
    ancestors re-breaks the exemption above."""
    pids = m._ancestor_pids()
    assert os.getpid() in pids
    assert os.getppid() in pids, (
        "the parent must be found, or an ancestor-held claim deadlocks"
    )


def test_a_stale_pid_is_taken_over(profile):
    """A crashed run leaves a record behind. Resumability matters more
    than strictness, so this must NOT be fatal."""
    m._owner_file(profile).write_text("999999", encoding="utf-8")
    m._claim_profile(profile)      # must not raise
    assert m._owner_file(profile).read_text(encoding="utf-8").strip() == str(os.getpid())


def test_a_corrupt_owner_file_is_taken_over(profile):
    m._owner_file(profile).write_text("not-a-pid", encoding="utf-8")
    m._claim_profile(profile)      # must not raise


def test_an_empty_owner_file_is_taken_over(profile):
    m._owner_file(profile).write_text("", encoding="utf-8")
    m._claim_profile(profile)


def test_the_guard_can_be_overridden(profile, monkeypatch):
    """An escape hatch, because a legitimate reason may exist that this
    code cannot foresee -- two different profiles, say."""
    monkeypatch.setenv("GOOGLE_AI_MODE_ALLOW_CONCURRENT", "true")
    m._owner_file(profile).write_text(str(os.getppid()), encoding="utf-8")
    m._claim_profile(profile)      # must not raise


# --- releasing ------------------------------------------------------------


def test_releasing_removes_our_claim(profile):
    m._claim_profile(profile)
    m._release_profile(profile)
    assert not m._owner_file(profile).exists()


def test_releasing_does_not_steal_a_peers_claim(profile):
    """shutdown() runs on every exit path, including one taken while a
    peer owns the profile. It must not clear their record."""
    m._owner_file(profile).write_text("424242", encoding="utf-8")
    m._release_profile(profile)
    assert m._owner_file(profile).read_text(encoding="utf-8").strip() == "424242"


def test_releasing_an_unclaimed_profile_is_a_no_op(profile):
    m._release_profile(profile)     # must not raise


# --- where the record lives ----------------------------------------------


def test_the_record_lives_outside_the_profile(profile):
    """reset_session() deletes the whole profile directory. A record
    inside it would vanish on every reset, silently disabling the guard
    exactly when repeated resets make a collision most likely."""
    owner = m._owner_file(profile)
    assert profile not in owner.parents
    assert owner.parent == profile.parent


# --- liveness detection ---------------------------------------------------


def test_our_own_pid_does_not_count_as_a_peer():
    """Otherwise the guard would refuse to let us reuse our own profile."""
    assert m._pid_is_alive(os.getpid()) is False


def test_a_nonexistent_pid_is_not_alive():
    assert m._pid_is_alive(999999) is False


def test_a_real_other_process_is_alive():
    """The parent process is by definition running."""
    assert m._pid_is_alive(os.getppid()) is True


@pytest.mark.parametrize("bad", [0, -1, -999])
def test_invalid_pids_are_not_alive(bad):
    assert m._pid_is_alive(bad) is False


# --- wiring ---------------------------------------------------------------


def test_the_claim_happens_before_chrome_is_launched():
    """Claiming after launch would be useless: the collision has already
    happened by then."""
    src = Path(m.__file__).read_text(encoding="utf-8")
    fn = src[src.index("def _get_page():"):]
    claim_at = fn.index("_claim_profile(profile)")
    launch_at = fn.index("launch_persistent_context")
    assert claim_at < launch_at


def test_shutdown_releases_the_claim():
    src = Path(m.__file__).read_text(encoding="utf-8")
    fn = src[src.index("def shutdown()"):]
    fn = fn[:fn.index("\ndef ")] if "\ndef " in fn else fn
    assert "_release_profile(profile)" in fn
