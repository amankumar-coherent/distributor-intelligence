"""Google AI Mode (udm=50) client, driven directly by Patchright.

Adapted from the reference implementation in "Brands V2/scraper/
scrape_ai_mode.py". The mechanics (persistent Chrome profile, link-inlining
text extraction, CAPTCHA detection, "Thinking..." wait) are kept as-is
because they encode hard-won details about how the page actually behaves.

What differs from the reference: it scrapes brand lists for a fixed
QUERY_TEMPLATE, whereas here every query is composed by the caller (and
usually by an LLM) for whatever the pipeline currently needs — screening a
batch of companies, filling export columns, identifying a Kompass category.
So this module deliberately exposes a generic ask(prompt) -> text, with no
opinion about the prompt or the shape of the answer.

PATCHRIGHT, NOT PLAIN PLAYWRIGHT — this is the whole reason this backend
works at all. An earlier version of this module used plain Playwright and
was abandoned because anonymous udm=50 traffic went straight to /sorry/
(CAPTCHA) on the very first query, and signing in to rescue it was
impossible (Google refuses OAuth from a WebDriver-controlled browser).
Patchright (github.com/Kaliiiiiiiiii-Vinyzu/patchright-python) is a
drop-in, stealth-patched Playwright fork that removes the automation
fingerprints Google detects. Re-verified live after the swap: anonymous
AI Mode answered in ~5s with no CAPTCHA and no sign-in, and correctly
followed a batched "reply with ONLY a JSON array" instruction. Do NOT
revert the import to `playwright.sync_api` — that reintroduces the
instant-CAPTCHA failure this module was previously written off for.

Why AI Mode rather than the AI Overview extension relay in
clients/google_ai_scraper.py:
  - AI Mode is a real generative answer, so it CAN follow an instruction
    like "classify these 25 companies and reply as JSON". The AI Overview
    path is search-backed: the same prompt returns "About 0 results"
    (verified) because no web page matches that text.
  - No login wall and no per-account usage cap, unlike the ChatGPT-account
    backends (the Codex/openai-oauth proxy hits "usage limit reached" on a
    free plan, and the ChatGPT web UI demands sign-in mid-run).

Chrome's omnibox rewrites udm=50 into an internal page, so the URL has to be
navigated by Patchright directly — that bypass is the other half of the trick.
"""
from __future__ import annotations

import json
import os
import re
import sys
import shutil
import threading
import time
import urllib.parse
from pathlib import Path

RESPONSE_TIMEOUT_MS = 90_000
THINKING_POLL_MS = 1500
CAPTCHA_WAIT_SECONDS = 60
# After this many CAPTCHAs in a row, stop asking. Google has flagged the
# profile/IP at that point and keeps challenging every query — solving them
# by hand doesn't clear it, so continuing just burns 60s per query. Reset to
# zero as soon as any query returns a real answer.
_MAX_CONSECUTIVE_CAPTCHAS = 3
_consecutive_captchas = 0

# Total CAPTCHAs seen, and when the last one happened. Kept SEPARATELY from
# _consecutive_captchas, which is reset to zero the moment a CAPTCHA is
# solved — that reset is correct for the circuit breaker but it destroys
# the only evidence that Google is throttling us. Callers deciding "was
# this empty result caused by blocking, or is the market really out of
# companies?" need a signal that a solve does NOT erase, or a solved
# CAPTCHA looks identical to no CAPTCHA at all. (Observed live: a run
# logged "CAPTCHA cleared" and then "no sign of blocking — treating as
# genuinely exhausted" for the same batch, stopping at 467/500.)
_captcha_total = 0
_last_captcha_at = 0.0

# How many times ONE query may hit a CAPTCHA, be solved, and be re-sent.
# Solving must not consume the caller's retry budget: with retries=1 the
# loop previously spent attempt 0 on the CAPTCHA and attempt 1 re-reading
# the same /sorry/ page, so a successful hand-solve still returned "" and
# three of those in a row made the batch runner declare the market
# "exhausted" (observed: a run stopped at 318 companies that way).
_MAX_CAPTCHA_ROUNDS_PER_QUERY = 3

# Consecutive failed queries (no usable answer for ANY reason: CAPTCHA,
# refusal, timeout, empty). After this many in a row the SESSION itself is
# treated as poisoned and rebuilt automatically — browser closed, profile
# reset, cookies discarded, re-seeded from the real Chrome profile.
#
# Why a separate counter from _consecutive_captchas: that one only counts
# CAPTCHAs, and it is reset by a successful solve. A session can fail
# persistently WITHOUT tripping it — every query timing out, or AI Mode
# refusing everything, or the answer pane rendering empty. Those states
# previously required the user to notice and restart by hand ("many times
# come something went wrong, ai cant generate"), because the only
# auto-recovery was for a crashed browser.
#
# Resetting the profile is the strongest lever available in-process:
# Google's flag rides on the session cookies, so discarding them and
# re-seeding gives the run a genuinely fresh identity.
_DEFAULT_MAX_CONSECUTIVE_FAILURES = 5
_consecutive_failures = 0
_session_resets = 0
# Cap resets per process so a permanently broken environment (no network,
# Chrome uninstalled) cannot loop forever rebuilding profiles.
_DEFAULT_MAX_SESSION_RESETS = 3


def _env_int(name: str, default: int) -> int:
    """Read an int env var at CALL time, not import time.

    Module-level os.getenv() runs before apply_env_overrides() has loaded
    .env, so a constant initialised that way always captures its default
    and the .env setting is silently ignored (this exact bug made
    MARKET_RESEARCH_BLOCK_COOLOFF=0 appear to do nothing).
    """
    raw = (os.getenv(name) or "").strip()
    if not raw:
        return default
    try:
        return max(1, int(raw))
    except ValueError:
        return default


def max_consecutive_failures() -> int:
    return _env_int("GOOGLE_AI_MODE_MAX_FAILURES", _DEFAULT_MAX_CONSECUTIVE_FAILURES)


def max_session_resets() -> int:
    return _env_int("GOOGLE_AI_MODE_MAX_RESETS", _DEFAULT_MAX_SESSION_RESETS)

# Minimum gap between consecutive AI Mode queries, with jitter.
#
# Without this, queries fired back-to-back (~10-16s apart, purely however
# long generation took) and Google eventually flagged the IP with "Our
# systems have detected unusual traffic from your computer network" —
# observed live at ~200 companies into a run, and the CAPTCHA it then
# serves often won't even load ("Cannot contact reCAPTCHA"), so it cannot
# be solved by hand either. Pacing is the only lever that actually reduces
# how often that happens, since the trigger is request RATE from one IP.
#
# Jitter matters as much as the delay: a perfectly regular interval is
# itself a bot signal. Override with GOOGLE_AI_MODE_QUERY_DELAY (seconds,
# 0 disables) if you want to trade speed against block risk.
#
# Default is deliberately conservative (30s + up to 10s jitter, so ~35s
# average): it cuts the request rate to roughly 80/hour, versus ~257/hour
# with no pacing at all, which is what got the IP flagged. On a long run
# the extra minutes are cheap next to losing the whole run to a block that
# can take hours to clear.
_DEFAULT_QUERY_DELAY_SECONDS = 10.0
_QUERY_JITTER_SECONDS = 5.0
_last_query_at = 0.0


def query_delay_seconds() -> float:
    raw = (os.getenv("GOOGLE_AI_MODE_QUERY_DELAY") or "").strip()
    if not raw:
        return _DEFAULT_QUERY_DELAY_SECONDS
    try:
        return max(0.0, float(raw))
    except ValueError:
        return _DEFAULT_QUERY_DELAY_SECONDS


def _pace_query() -> None:
    """Sleep just long enough that consecutive queries stay `delay` apart.

    Measures from the END of the previous query, so generation time already
    counts toward the gap — a slow 16s answer doesn't then also wait the
    full delay on top.
    """
    global _last_query_at
    delay = query_delay_seconds()
    if delay <= 0:
        _last_query_at = time.monotonic()
        return
    import random

    target = delay + random.uniform(0, _QUERY_JITTER_SECONDS)
    elapsed = time.monotonic() - _last_query_at
    if _last_query_at and elapsed < target:
        wait = target - elapsed
        _log(f"pacing: waiting {wait:.1f}s before next query (anti-rate-limit)")
        time.sleep(wait)
    _last_query_at = time.monotonic()

CAPTCHA_MARKERS = (
    "unusual traffic",
    "detected unusual traffic",
    "/sorry/index",
    "our systems have detected",
)

# One browser per process, reused across calls: launching Chrome costs
# several seconds, and a warm profile is also what keeps Google from
# CAPTCHA-ing us (a brand-new empty profile trips it immediately).
_browser_lock = threading.RLock()
_context = None
_playwright = None
# Set once if the automation profile had to be rebuilt to get an unpacked
# extension to register. Guards against a rebuild loop if it still fails.
_rebuilt_for_extension = False


def _log(msg: str) -> None:
    """Print a progress line, never crashing the run over an encoding issue.

    A Windows cp1252 console cannot encode most non-Western characters, and
    print() RAISES on an unmappable one — which killed a run mid-batch on a
    Turkish company name containing "S" with cedilla (U+015E). Since this
    pipeline deliberately searches worldwide, non-ASCII names are normal
    input, so logging must degrade rather than abort.
    """
    text = f"  [ai-mode] {msg}"
    try:
        print(text, flush=True)
    except UnicodeEncodeError:
        enc = (getattr(sys.stdout, "encoding", None) or "ascii")
        print(text.encode(enc, errors="replace").decode(enc), flush=True)


def ai_mode_enabled() -> bool:
    return (os.getenv("GOOGLE_AI_MODE_ENABLED", "false") or "false").strip().lower() in (
        "1",
        "true",
        "yes",
        "on",
    )


def ai_mode_headless() -> bool:
    """Headless is opt-in: a CAPTCHA can only be solved by hand in a visible
    window, and the profile has to pass one at least once before headless is
    safe (see the reference project's README)."""
    return (os.getenv("GOOGLE_AI_MODE_HEADLESS", "false") or "false").strip().lower() in (
        "1",
        "true",
        "yes",
        "on",
    )


def extensions_require_chromium() -> bool:
    """Whether to switch to bundled Chromium when extensions are requested.

    MEASURED on this machine (Chrome 152, throwaway profiles, ground truth
    read from chrome://extensions-internals rather than a worker count):

      real Chrome  + --load-extension                    -> NOT loaded
      real Chrome  + --load-extension + disable-except    -> NOT loaded
      bundled Chromium + the same flags                   -> LOADED
                          ("Free AI ReCaptcha Solver ... [COMMAND_LINE]")

    Cause: this Chrome is enterprise-managed (HKLM\\SOFTWARE\\Policies\\
    Google\\Chrome\\ExtensionInstallForcelist is set), and managed Chrome
    refuses command-line-loaded unpacked extensions outright. No flag
    combination or profile state recovers it — an earlier conclusion that
    a fresh profile fixes this was wrong; the apparent success was Chrome's
    own component worker (network_speech_synthesis) being miscounted.

    So: if extensions are requested, Chromium is the only build that can
    honour them. The trade-off is real and the caller is warned — Chromium
    gets CAPTCHA'd far more readily than real Chrome. Set
    GOOGLE_AI_MODE_EXTENSIONS_NEED_CHROMIUM=false to keep real Chrome and
    accept that the extension will not load.
    """
    raw = (os.getenv("GOOGLE_AI_MODE_EXTENSIONS_NEED_CHROMIUM") or "true").strip().lower()
    return raw in ("1", "true", "yes", "on")


def ai_mode_browser_channel() -> str | None:
    """Which browser build to drive: real Chrome (default) or bundled Chromium.

    Returns the Playwright `channel` value, or None for bundled Chromium
    (Playwright's own build, installed via `patchright install chromium`).

    Prefer real Chrome ("chrome", the default). Measured on this machine:
    real Chrome answers AI Mode with no CAPTCHA, while bundled Chromium got
    CAPTCHA'd on essentially every query — Chromium's fingerprint is much
    easier for Google to flag, and no CAPTCHA-solver extension reliably
    clears it.

    Chromium is kept only as an escape hatch — but it IS currently the only
    build that loads the CAPTCHA-solver extension here.

    CORRECTION (re-measured against chrome://extensions-internals, which is
    the only trustworthy signal — a service-worker count includes Chrome's
    own components): the claim that once sat here, that real Chrome loads
    an unpacked extension "as long as the automation profile is CREATED
    with the extension present", is WRONG on this machine. Tested directly:
    the profile was moved aside, `--install` staged the extension, and a
    brand-new profile was created with it present. Result:

      real Chrome, fresh profile + staged extension
          -> extensions-internals: 0 non-component extensions
      bundled Chromium, same staged extension
          -> extensions-internals: 1 ("Free AI ReCaptcha Solver by Raptor
             ... [COMMAND_LINE]"), 1 background worker

    The registry does have HKLM\\SOFTWARE\\Policies\\Google\\Chrome\\
    ExtensionInstallForcelist set (no Blocklist/Allowlist), i.e. this
    Chrome is enterprise-managed, matching the older
    extensions_require_chromium() note.

    So the choice is real and unavoidable: real Chrome (few CAPTCHAs, NO
    solver) or Chromium (solver active, many more CAPTCHAs).

    Set GOOGLE_AI_MODE_BROWSER=chromium to use the bundled build.
    """
    raw = (os.getenv("GOOGLE_AI_MODE_BROWSER") or "chrome").strip().lower()
    return None if raw == "chromium" else "chrome"


def active_profile_dir() -> str:
    """The profile directory the next launch will actually use.

    Resolves the extensions-force-Chromium switch, so callers (the CLI's
    --list/--reset-profile, tests) report on the SAME directory the
    pipeline will open rather than the Chrome-channel default.
    """
    channel = ai_mode_browser_channel()
    if (
        (extension_paths() or staged_extension_paths())
        and channel == "chrome"
        and extensions_require_chromium()
    ):
        channel = None
    return _profile_dir_for(channel)


def _profile_dir_for(channel: str | None) -> str:
    """Profile path for a specific browser build.

    Chrome and Chromium must not share a profile directory: they are
    different builds with incompatible profile state (and different cookie
    encryption), so pointing both at one directory corrupts it. Chromium
    therefore gets its own suffixed directory.
    """
    base = _profile_dir()
    if channel is None:
        return f"{base}_chromium"
    return base


def _profile_dir() -> str:
    """Persistent Chrome profile for automation.

    Deliberately NOT the user's real profile directory: Chrome refuses
    remote debugging against the default profile, and locking it would stop
    the user browsing. A clone seeded from the real profile keeps cookies
    and history, which is what makes Google treat the session as human.
    """
    configured = (os.getenv("GOOGLE_AI_MODE_PROFILE_DIR") or "").strip()
    if configured:
        return configured
    from vendor_intel.config import _project_root

    return str(_project_root() / "data" / "ai_mode_chrome_profile")


def _seed_profile_from_real_chrome(dest: Path) -> None:
    """Copy cookies/history from the user's real Chrome profile on first use.

    Best-effort: a fresh profile still works, it's just likelier to hit a
    CAPTCHA on the first few queries.

    "Local State" MUST be copied along with the per-profile files, and it
    lives one level UP (in the User Data root, not inside Profile N).
    Chrome encrypts the cookie DB with an `os_crypt.encrypted_key` stored
    there; without it Chrome generates a brand-new key and every copied
    cookie silently fails to decrypt — the profile looks signed OUT even
    though Cookies was copied byte-for-byte (observed exactly that: the
    seeded profile still showed Google's "Sign in" button). Copy it only
    when seeding a fresh profile, never over an existing one, or Chrome
    would lose the key for cookies it wrote itself.

    Note this only carries the SESSION. It does not make Google's OAuth
    sign-in flow work inside an automated browser — that is separately
    blocked ("This browser or app may not be secure"), which is why AI Mode
    is designed to work anonymously in the first place.
    """
    src_root = (os.getenv("CHATGPT_SCRAPER_CHROME_PROFILE_DIR") or "").strip()
    src_name = (os.getenv("CHATGPT_SCRAPER_CHROME_PROFILE_DIRECTORY") or "Default").strip()
    if not src_root:
        return
    src = Path(src_root) / src_name
    if not src.is_dir():
        return

    # Cookie-decryption key, from the User Data ROOT (sibling of Profile N).
    try:
        local_state = Path(src_root) / "Local State"
        if local_state.is_file():
            shutil.copy2(local_state, dest / "Local State")
    except Exception:
        pass  # non-fatal; cookies just won't decrypt, same as before

    dest_profile = dest / "Default"
    dest_profile.mkdir(parents=True, exist_ok=True)
    for name in ("Cookies", "History", "Web Data", "Preferences", "Login Data", "Network"):
        s = src / name
        try:
            if s.is_dir():
                shutil.copytree(s, dest_profile / name, dirs_exist_ok=True)
            elif s.is_file():
                shutil.copy2(s, dest_profile / name)
        except Exception:
            pass  # non-fatal; a cold profile still functions

    # The real Chrome is usually RUNNING while we copy, so its live
    # Preferences records exit_type="Crashed" and carries its open-tab
    # session list. Seeding those verbatim made every fresh automation
    # session boot into "Chromium didn't shut down correctly / Restore
    # pages?" with a row of restored about:blank tabs, and the modal
    # blocked navigation — the page never loaded. Only the SESSION
    # (cookies) is wanted from the real profile, never its crash state.
    _mark_profile_exited_cleanly(dest)


def _mark_profile_exited_cleanly(profile: Path) -> None:
    """Clear the crash flag so Chrome does not show "Restore pages?".

    Chrome decides to offer session restore from `profile.exit_type` in
    Preferences. Any value other than "Normal" — set when Chrome is killed,
    when the machine sleeps mid-run, or (the common case here) inherited by
    copying a RUNNING Chrome's Preferences — raises a modal on the next
    launch. That modal sits on top of the page and swallows the navigation,
    so the pipeline sees a blank page and reports an empty answer.

    Also drops the copied startup/session tab list, which is what produced
    the row of restored about:blank tabs.

    Best-effort: if Preferences is unreadable the launch still proceeds,
    it just may show the bubble.
    """
    prefs = profile / "Default" / "Preferences"
    if not prefs.is_file():
        return
    try:
        data = json.loads(prefs.read_text(encoding="utf-8", errors="replace"))
    except (json.JSONDecodeError, OSError):
        return

    prof = data.setdefault("profile", {})
    prof["exit_type"] = "Normal"
    prof["exited_cleanly"] = True
    session = data.setdefault("session", {})
    session["restore_on_startup"] = 5      # 5 == open the New Tab Page
    session["startup_urls"] = []

    try:
        prefs.write_text(json.dumps(data), encoding="utf-8")
    except OSError:
        pass

    # Chrome replays tabs from these regardless of the crash flag, so a
    # copied session would still reopen the old window's tabs.
    for name in ("Sessions", "Session Storage"):
        shutil.rmtree(profile / "Default" / name, ignore_errors=True)
    for name in ("Current Session", "Current Tabs", "Last Session", "Last Tabs"):
        try:
            (profile / "Default" / name).unlink(missing_ok=True)
        except OSError:
            pass


def keep_profile() -> bool:
    """Whether to PRESERVE the automation profile instead of rebuilding it.

    Default TRUE. This is what makes the profile "better" in the sense the
    user asked for: extensions they add by hand through Chrome's own
    chrome://extensions UI survive across runs.

    The old behaviour (rebuild whenever an extension didn't register) was
    correct for `--load-extension`, which Chrome treats as a
    command-line-only load and never records in the profile. But it also
    silently deleted anything the user installed manually, which made
    hand-adding an extension impossible. With the profile preserved,
    manual installs persist exactly like they do in normal Chrome.

    Set GOOGLE_AI_MODE_KEEP_PROFILE=false to restore the old rebuild
    behaviour (useful only if a profile gets into a bad state; prefer
    `--reset-profile` on manage_ai_mode_extensions.py, which is explicit).
    """
    raw = (os.getenv("GOOGLE_AI_MODE_KEEP_PROFILE") or "true").strip().lower()
    return raw in ("1", "true", "yes", "on")


def allow_all_extensions() -> bool:
    """Whether to let the profile's OWN installed extensions run.

    Default TRUE, and this is the second half of the fix. Passing
    `--disable-extensions-except=<path>` disables every extension the
    profile itself has installed, which is why chrome://extensions looked
    completely EMPTY in the automation window and why "Load unpacked"
    appeared to do nothing: the extension was installed fine, then
    immediately suppressed by the flag.

    With this on we omit that flag entirely, so anything installed in the
    profile — by hand or by install_extension() — loads normally.

    Set GOOGLE_AI_MODE_ALLOW_ALL_EXTENSIONS=false to go back to loading
    ONLY the GOOGLE_AI_MODE_EXTENSIONS paths and nothing else.
    """
    raw = (os.getenv("GOOGLE_AI_MODE_ALLOW_ALL_EXTENSIONS") or "true").strip().lower()
    return raw in ("1", "true", "yes", "on")


def extension_paths() -> list[str]:
    """Unpacked Chrome extensions to load into the automation browser.

    Set GOOGLE_AI_MODE_EXTENSIONS to one or more paths, separated by ";"
    (Windows-friendly, since paths contain ":"). Each path must be the
    folder that directly CONTAINS manifest.json — pointing at a repo root
    whose manifest lives in a subfolder silently loads nothing, which is
    the most common mistake here.

    Intended for a CAPTCHA solver: Google shows an "unusual traffic"
    interstitial during long AI Mode runs, and a solver extension can clear
    it without a human present.

    Note this is now only ONE of two ways an extension can be present.
    Extensions installed into the profile itself (by hand via
    chrome://extensions, or by install_extension()) load without appearing
    here, because keep_profile()/allow_all_extensions() no longer wipe or
    suppress them.
    """
    raw = (os.getenv("GOOGLE_AI_MODE_EXTENSIONS") or "").strip()
    if not raw:
        return []

    valid: list[str] = []
    for part in raw.split(";"):
        candidate = part.strip().strip('"')
        if not candidate:
            continue
        path = Path(candidate)
        if not path.is_dir():
            _log(f"extension path is not a directory, skipping: {path}")
            continue
        if not (path / "manifest.json").is_file():
            # Try the common one-level-down layout before giving up, since
            # most extension repos keep the manifest in extension/.
            nested = path / "extension"
            if (nested / "manifest.json").is_file():
                _log(f"using {nested} (manifest.json is in the extension/ subfolder)")
                valid.append(str(nested.resolve()))
                continue
            _log(f"no manifest.json in {path} — not a loadable extension, skipping")
            continue
        valid.append(str(path.resolve()))
    return valid


def _extension_identity(path: str | Path) -> str:
    """Identify an extension by its manifest name, for dedupe purposes.

    Needed because the same extension can be referenced by two different
    paths at once — the original download folder (via .env) and the copy
    made by install_extension(). Loading both registers two service
    workers for one extension, which is wasteful and lets two solver
    instances fight over the same CAPTCHA. Falls back to the folder name
    if the manifest can't be read.
    """
    p = Path(path)
    try:
        data = json.loads((p / "manifest.json").read_text(encoding="utf-8", errors="replace"))
        name = str(data.get("name") or "").strip()
        if name:
            return name.casefold()
    except (json.JSONDecodeError, OSError):
        pass
    return p.name.casefold()


def unpacked_extension_id(path: str | Path) -> str:
    """Chrome's extension ID for an unpacked extension at `path`.

    Chrome derives it deterministically: SHA-256 of the absolute path (as
    Chrome spells it), first 16 bytes, each nibble mapped 0-f -> a-p. This
    lets us tell OUR extension's service worker apart from Chrome's own
    component extensions (pdf, hangout_services,
    network_speech_synthesis), which always run and otherwise make a
    worker count look successful when the solver never loaded at all.

    Best-effort: returns "" if the ID can't be derived, and callers treat
    that as "can't verify" rather than failure.
    """
    import hashlib

    try:
        # Chrome hashes the path in its native form. On Windows that is the
        # UTF-16LE bytes of the drive-letter-uppercased absolute path.
        resolved = Path(path).resolve()
        raw = str(resolved)
        if os.name == "nt":
            if len(raw) > 1 and raw[1] == ":":
                raw = raw[0].upper() + raw[1:]
            digest = hashlib.sha256(raw.encode("utf-16-le")).hexdigest()
        else:
            digest = hashlib.sha256(raw.encode("utf-8")).hexdigest()
        return "".join(chr(ord("a") + int(c, 16)) for c in digest[:32])
    except (OSError, ValueError):
        return ""


def _own_extension_workers(context, loaded_paths) -> list[str]:
    """Service-worker/background-page URLs belonging to OUR extensions.

    Takes the exact paths that were passed to Chrome, rather than
    re-deriving them from env/staging: the launch path dedupes those two
    sources, so re-deriving can produce a different set than Chrome
    actually received and mis-attribute (or miss) a worker.

    Filters out Chrome's component extensions by matching derived unpacked
    extension IDs, so "extension active" means the solver really started
    rather than "Chrome booted its own built-ins".
    """
    try:
        urls = [w.url for w in context.service_workers]
        urls += [p.url for p in context.background_pages]
    except Exception:
        return []

    expected_ids = {unpacked_extension_id(p) for p in loaded_paths}
    expected_ids.discard("")
    if not expected_ids:
        # Can't derive IDs — fall back to excluding Chrome's known
        # component extensions rather than trusting a bare count.
        return [u for u in urls if not any(c in u for c in _COMPONENT_EXTENSION_IDS)]
    return [u for u in urls if any(eid in u for eid in expected_ids)]


# Chrome's own always-present component extensions. Counting these as a
# loaded solver is a false positive.
_COMPONENT_EXTENSION_IDS = (
    "mhjfbmdgcfjbbpaeojofohoefgiehjai",  # pdf viewer
    "nkeimhogjdpnpccoofpliimaahmaaome",  # hangout services
    "fignfifoniblkonapihmkfakmlgkbkcf",  # network speech synthesis
    "neajdppkdcdipfabeoofebfddakdcjhd",  # google network speech
)


def _rebuild_profile_preserving_staged(profile: Path) -> None:
    """Delete the automation profile but keep staged extensions.

    A rebuild exists to clear Chrome's browsing state (cookies, the
    "already ran without this extension" record). It should never cost the
    user their installed extensions, so the `_unpacked_extensions` staging
    directory is moved aside and restored.
    """
    # Staged extensions live OUTSIDE the profile (see _staging_root), so a
    # reset simply cannot touch them — no move-aside/restore dance, and no
    # risk of recreating the profile dir and defeating the first_run check.
    shutil.rmtree(profile, ignore_errors=True)
    # If the tree could not be removed (a file still held open by the
    # closing Chrome), the leftover Preferences keeps exit_type="Crashed"
    # and the replacement session boots into the restore modal. Clear it.
    if profile.exists():
        _mark_profile_exited_cleanly(profile)
    if staged_extension_paths():
        _log("staged extensions are stored outside the profile — unaffected")


def _installed_extension_ids(profile: Path) -> list[str]:
    """Extension IDs actually installed IN the automation profile.

    Reads the profile's own Preferences, which is where Chrome records
    every extension installed through the UI (or seeded by
    install_extension()). Unlike `--load-extension` paths, these persist
    across runs, so this is the list that tells you whether hand-added
    extensions survived.

    Chrome's own component extensions are filtered out — they are always
    present and would make an empty profile look populated.
    """
    prefs = profile / "Default" / "Preferences"
    if not prefs.is_file():
        return []
    try:
        data = json.loads(prefs.read_text(encoding="utf-8", errors="replace"))
    except (json.JSONDecodeError, OSError):
        return []

    settings = (data.get("extensions") or {}).get("settings") or {}
    found: list[str] = []
    for ext_id, meta in settings.items():
        if not isinstance(meta, dict):
            continue
        # location 5 == component/built-in; 10 == external policy. Only
        # user-visible installs (unpacked=4, normal=1) are interesting.
        if meta.get("location") in (5, 10):
            continue
        if meta.get("state") == 0:  # explicitly disabled by the user
            continue
        found.append(ext_id)
    return found


def _staging_root() -> Path:
    """Where installed extension copies live.

    Deliberately a SIBLING of the profile, not inside it. Keeping it inside
    caused two distinct bugs: a profile reset had to move it out and back
    (and restoring it recreated the profile directory, which defeated the
    `first_run` check), and Chrome vs Chromium use separate profiles, so a
    copy staged for one was invisible to the other. As a sibling it is
    shared by both builds and untouched by any profile reset.
    """
    base = Path(_profile_dir())
    return base.parent / f"{base.name}_extensions"


def install_extension(source: str) -> bool:
    """Install an unpacked extension INTO the automation profile so it
    persists across runs.

    Why this exists: `--load-extension` is a command-line-only load. Chrome
    never records it in the profile, so it vanishes on any rebuild and can
    never be managed from chrome://extensions. Copying the extension into
    the profile's own Extensions directory and registering it in
    Preferences makes it a real, persistent, user-manageable install.

    Returns True if the extension was staged successfully. The caller
    should then launch the browser once so Chrome picks it up.
    """
    src = Path(source.strip().strip('"'))
    if not (src / "manifest.json").is_file():
        nested = src / "extension"
        if (nested / "manifest.json").is_file():
            src = nested
        else:
            _log(f"no manifest.json in {src} (or its extension/ subfolder)")
            return False

    # Name the staged copy after the extension, not the folder. Selecting
    # `captcha-raptor\extension` would otherwise stage it as plain
    # "extension", so a second install of the same layout would collide and
    # the listing would be unreadable.
    stage_name = src.name
    if stage_name.lower() in ("extension", "extensions", "src", "public", "dist", "build"):
        stage_name = src.parent.name or stage_name

    staging = _staging_root() / stage_name
    staging.parent.mkdir(parents=True, exist_ok=True)
    if staging.exists():
        shutil.rmtree(staging, ignore_errors=True)
    try:
        shutil.copytree(src, staging)
    except OSError as exc:
        _log(f"could not copy extension into the profile: {exc}")
        return False

    _log(f"installed extension: {staging}")
    _log(
        "loaded from that copy on every run, so it survives profile resets "
        "and the original download folder can be moved or deleted"
    )
    return True


def staged_extension_paths() -> list[str]:
    """Extensions previously staged into the profile by install_extension().

    Loaded on every launch, which is what gives a staged extension the
    persistence a raw --load-extension path never had: the copy lives
    inside the profile directory itself.
    """
    staging_root = _staging_root()
    if not staging_root.is_dir():
        return []
    out: list[str] = []
    for child in sorted(staging_root.iterdir()):
        if (child / "manifest.json").is_file():
            out.append(str(child.resolve()))
    return out


def open_profile_for_extensions(timeout_seconds: int = 900) -> bool:
    """Open the automation profile's Chrome with extension management fully
    usable, so extensions can be added by hand and will persist.

    This is the interactive counterpart to install_extension(). It differs
    from a normal pipeline launch in exactly the ways that previously made
    hand-adding an extension impossible:

      * no --disable-extensions-except, so chrome://extensions shows the
        profile's real extension list instead of appearing empty;
      * the profile is preserved afterwards, so a manual install survives.

    Leave the window open, switch on Developer mode, use "Load unpacked",
    then close the window (or press Ctrl+C here) when done.
    """
    page = _get_page()
    profile = Path(active_profile_dir())
    try:
        page.goto("chrome://extensions/", wait_until="domcontentloaded", timeout=30_000)
    except Exception:
        pass  # chrome:// navigation can report oddly; the tab still opens

    _log("=" * 68)
    _log("Automation profile is open on chrome://extensions.")
    _log("  1. Turn ON 'Developer mode' (top-right toggle)")
    _log("  2. Click 'Load unpacked'")
    _log("  3. Select the folder that CONTAINS manifest.json, e.g.")
    _log(r"     ...\chrome-scraper-extensions\captcha-raptor\extension")
    _log("")
    _log("This install PERSISTS — the profile is no longer wiped between runs.")
    _log(f"Waiting up to {timeout_seconds}s. Press Ctrl+C when you're done.")
    _log("=" * 68)

    deadline = time.monotonic() + timeout_seconds
    try:
        while time.monotonic() < deadline:
            time.sleep(3)
            if _context is None:
                break
            try:
                if not _context.pages:  # user closed the window
                    break
            except Exception:
                break
    except KeyboardInterrupt:
        _log("done — saving the profile")

    shutdown()
    ids = _installed_extension_ids(profile)
    if ids:
        _log(f"profile now has {len(ids)} installed extension(s): {', '.join(ids)}")
        return True
    _log("no user-installed extensions recorded in the profile yet")
    return False


def _single_page(context):
    """Return ONE page, closing any extras.

    A restored session (or a leftover crash-restore window) opens a row of
    about:blank tabs. Blindly taking pages[0] then picked a stale tab that
    nothing navigated, so the pipeline waited on a permanently blank page.
    Prefer a page that already holds a real URL, close the rest, and keep
    exactly one alive so the caller always drives the page it can see.
    """
    try:
        pages = [pg for pg in context.pages if not pg.is_closed()]
    except Exception:
        pages = []
    if not pages:
        return context.new_page()

    real = [pg for pg in pages if (pg.url or "about:blank") not in ("about:blank", "")]
    keep = real[0] if real else pages[0]
    if len(pages) > 1:
        _log(f"closing {len(pages) - 1} leftover tab(s) from a restored session")
    for pg in pages:
        if pg is keep:
            continue
        try:
            pg.close()
        except Exception:
            pass
    return keep


def _clear_stale_event_loop() -> bool:
    """Uninstall a CLOSED event loop from this thread. Returns True if one
    was removed.

    Sync Patchright refuses to run when an event loop is installed on the
    calling thread — "It looks like you are using Playwright Sync API
    inside the asyncio loop." That check does not care whether the loop is
    still usable, so a loop some earlier async helper left behind after
    closing it permanently disables AI Mode for the rest of the process:
    every query fails identically, the failure streak fills up, and the
    session-reset budget is spent on a problem no relaunch can fix.

    Only a CLOSED loop is removed. A live loop means we are genuinely
    being called from async code, where the sync API really is the wrong
    tool and the error is correct.
    """
    import asyncio
    import warnings

    try:
        asyncio.get_running_loop()
        return False            # genuinely inside a running loop
    except RuntimeError:
        pass

    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            loop = asyncio.get_event_loop_policy().get_event_loop()
    except Exception:  # noqa: BLE001
        return False            # nothing installed: the good case

    if loop is None or not loop.is_closed():
        return False
    try:
        asyncio.set_event_loop(None)
    except Exception:  # noqa: BLE001
        return False
    _log("cleared a closed event loop left on this thread (was blocking sync Playwright)")
    return True


class ProfileInUse(RuntimeError):
    """Another live process is already driving the automation profile."""


def _owner_file(profile: Path) -> Path:
    # Beside the profile, not inside it: a profile rebuild must not
    # silently drop the ownership record.
    return profile.parent / f"{profile.name}.owner"


def _pid_is_alive(pid: int) -> bool:
    if pid <= 0 or pid == os.getpid():
        return False
    try:
        import ctypes

        # PROCESS_QUERY_LIMITED_INFORMATION; works without elevation.
        handle = ctypes.windll.kernel32.OpenProcess(0x1000, False, pid)
        if not handle:
            return False
        try:
            code = ctypes.c_ulong()
            if ctypes.windll.kernel32.GetExitCodeProcess(handle, ctypes.byref(code)):
                return code.value == 259          # STILL_ACTIVE
            return True
        finally:
            ctypes.windll.kernel32.CloseHandle(handle)
    except Exception:  # noqa: BLE001
        # Non-Windows, or the API is unavailable. Fall back to os.kill(0).
        try:
            os.kill(pid, 0)
            return True
        except (OSError, ProcessLookupError):
            return False


def _ancestor_pids(limit: int = 12) -> set[int]:
    """PIDs of this process and its ancestors.

    An ancestor holding the profile claim is not a competitor: the only
    way that happens is a parent that launched us and is now blocked
    waiting for us -- which is exactly how run_market_research invokes the
    verifier. Treating it as a rival deadlocks the pair.
    """
    pids = {os.getpid()}
    try:
        import psutil  # type: ignore

        proc = psutil.Process()
        for parent in proc.parents()[:limit]:
            pids.add(parent.pid)
        return pids
    except Exception:  # noqa: BLE001
        pass

    # No psutil. Walk the chain with PowerShell CIM rather than wmic:
    # wmic is deprecated and absent from Windows 11, so the wmic version
    # of this silently returned only our own PID -- which left the
    # self-deadlock in place.
    #
    # One PowerShell call builds the whole chain; spawning one per level
    # cost ~200ms each and this runs on every browser launch.
    try:
        import subprocess

        script = (
            "$p=" + str(os.getpid()) + ";"
            "for($i=0;$i -lt " + str(limit) + ";$i++){"
            "$q=Get-CimInstance Win32_Process -Filter \"ProcessId=$p\";"
            "if(-not $q){break};"
            "$p=$q.ParentProcessId;"
            "if($p -le 0){break};"
            "Write-Output $p}"
        )
        out = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", script],
            capture_output=True, text=True, timeout=15,
        ).stdout
        for line in (out or "").split():
            try:
                pid = int(line.strip())
            except ValueError:
                continue
            if pid > 0:
                pids.add(pid)
    except Exception:  # noqa: BLE001
        pass
    return pids


def _claim_profile(profile: Path) -> None:
    """Record this process as the profile's owner, or raise ProfileInUse.

    Concurrency on one persistent profile is never valid: the markets
    share a single profile and a single IP by design. Detecting it here
    turns an unbreakable relaunch loop into one clear error at startup.
    """
    if os.getenv("GOOGLE_AI_MODE_ALLOW_CONCURRENT", "").strip().lower() in (
        "1", "true", "yes",
    ):
        return

    owner = _owner_file(profile)
    try:
        recorded = int((owner.read_text(encoding="utf-8").split() or ["0"])[0])
    except (OSError, ValueError):
        recorded = 0

    # An ancestor is not a competitor: run_market_research holds the
    # claim from discovery while it waits on the verifier SUBPROCESS it
    # spawned, and the two run strictly one after the other. Refusing the
    # child there deadlocked the pair and burned the whole session-reset
    # budget without a single query succeeding.
    if recorded and recorded not in _ancestor_pids() and _pid_is_alive(recorded):
        # WARN, never raise.
        #
        # This check used to raise ProfileInUse, and the cure was worse
        # than the disease: it refused the verifier subprocess because the
        # verifier's own PARENT held the claim, producing 159 failures,
        # every chunk dropped fail-closed and an entire verification pass
        # lost. The ancestor exemption above fixes that specific case, but
        # any relationship the walk cannot see -- a shell wrapper in
        # between, a recycled PID, a detached relauncher -- would turn a
        # working run into a total outage.
        #
        # Duplicate runs are properly prevented at the LAUNCHER (the PID
        # lock in run_four_markets.ps1), which is where the duplication
        # happens and where refusing to start is the right answer. Down
        # here, the useful contribution is naming the other process; if a
        # collision is real, Patchright reports "profile is already in
        # use" and _looks_like_dead_browser triggers a relaunch anyway.
        _log(
            f"WARNING: process {recorded} also claims {profile.name}. "
            "If two pipelines really are running, stop one — they collide "
            "on the profile lock. Continuing anyway."
        )

    try:
        owner.parent.mkdir(parents=True, exist_ok=True)
        owner.write_text(str(os.getpid()), encoding="utf-8")
    except OSError:
        pass          # advisory only; never block the run on this


def _release_profile(profile: Path) -> None:
    """Drop our ownership claim, if it is still ours."""
    owner = _owner_file(profile)
    try:
        if int(owner.read_text(encoding="utf-8").strip() or 0) == os.getpid():
            owner.unlink(missing_ok=True)
    except (OSError, ValueError):
        pass


def _get_page():
    """Launch (once) and return the shared Playwright page."""
    global _context, _playwright, _rebuilt_for_extension
    # Must happen BEFORE any Patchright call, including the reuse path:
    # driving an existing context is just as blocked by a stale loop as
    # launching a new one.
    _clear_stale_event_loop()
    with _browser_lock:
        if _context is not None:
            return _single_page(_context)

        # Patchright, not playwright — see module docstring. Plain
        # Playwright's automation fingerprints get this instantly CAPTCHA'd.
        from patchright.sync_api import sync_playwright

        # Resolve the browser build FIRST: extensions can force a switch to
        # Chromium, and Chromium needs its own profile directory (sharing
        # one with Chrome corrupts it).
        channel = ai_mode_browser_channel()
        _wanted_extensions = extension_paths() or staged_extension_paths()
        if _wanted_extensions and channel == "chrome" and extensions_require_chromium():
            channel = None
        profile = Path(_profile_dir_for(channel))
        # "First run" means Chrome has no STATE here, not merely that the
        # directory is absent: a reset can leave the directory present but
        # empty, and treating that as an existing profile skips the
        # re-seeding step that keeps the session looking human.
        first_run = not (profile / "Default").exists()
        profile.mkdir(parents=True, exist_ok=True)
        # Before launching Chrome: if a peer process already owns this
        # profile, launching produces "the profile is already in use" on
        # every attempt forever. Fail loudly once instead.
        _claim_profile(profile)
        if first_run:
            _seed_profile_from_real_chrome(profile)
            _log(f"created automation profile at {profile}")

        # Clear the crash flag on EVERY launch, not only when seeding.
        # A run killed with Ctrl+C, a machine that slept, or a
        # reset_session() that tore the browser down mid-query all leave
        # exit_type="Crashed" behind — and the resulting "Restore pages?"
        # modal blocks the first navigation, which surfaces as an empty
        # answer and then as a pointless auto-recovery loop.
        if not first_run:
            _mark_profile_exited_cleanly(profile)

        args = [
            "--profile-directory=Default",
            # Belt-and-braces against the restore prompt: even with a clean
            # exit_type, an unclean kill between the pref write and the
            # launch can re-raise it. These make Chrome never offer it.
            "--hide-crash-restore-bubble",
            "--disable-session-crashed-bubble",
            "--no-first-run",
            "--no-default-browser-check",
            # Never replay a previous window's tabs — the pipeline drives
            # exactly one page and restored tabs only cost memory and
            # confuse the "which page is mine" lookup below.
            "--disable-features=InfiniteSessionRestore",
        ]
        # Env-configured paths PLUS anything installed by
        # install_extension(). Installed copies live beside the profile
        # (see _staging_root), so they survive profile resets and remain
        # valid even if the original download folder is moved or deleted.
        extensions = extension_paths()
        staged_paths = staged_extension_paths()
        # A staged copy SUPERSEDES the .env path it was made from: loading
        # both would register the same extension twice (two service
        # workers racing on the same CAPTCHA). Match on the extension's
        # own identity (manifest name) rather than the path, since the
        # staged copy deliberately lives somewhere else.
        staged_names = {_extension_identity(p) for p in staged_paths}
        deduped = []
        for path in extensions:
            if _extension_identity(path) in staged_names:
                _log(f"skipping .env path (installed in profile instead): {path}")
                continue
            deduped.append(path)
        extensions = deduped
        for staged in staged_paths:
            if staged not in extensions:
                extensions.append(staged)
                _log(f"loading staged extension from profile: {staged}")
        headless = ai_mode_headless()
        installed = _installed_extension_ids(profile)
        if installed:
            _log(
                f"profile has {len(installed)} extension(s) installed in it "
                "(these persist across runs)"
            )
        if extensions:
            joined = ",".join(extensions)
            args.append(f"--load-extension={joined}")
            if not allow_all_extensions():
                # Opt-in only. This flag disables everything the profile
                # itself installed, which is what made chrome://extensions
                # look empty and hand-adding an extension impossible.
                args.append(f"--disable-extensions-except={joined}")
                _log(
                    "NOTE: --disable-extensions-except is set, so extensions you "
                    "installed in the profile by hand will NOT load. Unset "
                    "GOOGLE_AI_MODE_ALLOW_ALL_EXTENSIONS=false to allow them."
                )
            for path in extensions:
                _log(f"loading extension: {path}")
            _log("verifying the extension actually registered...")
        if extensions or installed:
            if headless:
                # MV3 service workers don't run in old headless, and even
                # --headless=new is unreliable for extensions; a solver that
                # never starts is worse than no solver, so say so plainly.
                _log(
                    "WARNING: extensions are unreliable in headless Chrome — "
                    "set GOOGLE_AI_MODE_HEADLESS=false if the solver doesn't run"
                )

        if extensions and ai_mode_browser_channel() == "chrome" and channel is None:
            # Managed Chrome silently refuses command-line unpacked
            # extensions (see extensions_require_chromium). Switching build
            # is the only way to honour the request, so say so loudly
            # rather than pretending the solver is running.
            _log("=" * 68)
            _log("extensions requested -> switching to bundled Chromium.")
            _log("Reason: this Chrome is enterprise-managed and refuses")
            _log("command-line unpacked extensions (verified: not loaded).")
            _log("TRADE-OFF: Chromium is CAPTCHA'd more often than real Chrome.")
            _log("To keep real Chrome instead (extension will NOT load), set")
            _log("  GOOGLE_AI_MODE_EXTENSIONS_NEED_CHROMIUM=false")
            _log("=" * 68)
        if (
            extensions
            and channel == "chrome"
            and not first_run
            and not installed
            and not staged_paths  # a staged copy already persists; no warning needed
        ):
            # An unpacked extension passed via --load-extension only
            # registers on a profile CREATED with it present. Installing it
            # INTO the profile (manage_ai_mode_extensions.py --install, or
            # chrome://extensions by hand) avoids this entirely, which is
            # the recommended route now.
            _log(
                "NOTE: this profile already existed, so a --load-extension "
                "extension may not register. Prefer installing it into the "
                "profile: python manage_ai_mode_extensions.py --install <path>"
            )
        _playwright = sync_playwright().start()
        launch_kwargs = dict(
            headless=headless,
            viewport={"width": 1280, "height": 900},
            args=args,
        )
        if channel:  # None => Playwright's bundled Chromium
            launch_kwargs["channel"] = channel
        _context = _playwright.chromium.launch_persistent_context(
            str(profile), **launch_kwargs
        )

        if extensions or installed:
            # MV3 extensions register a service worker; if none appears the
            # extension is not running, whatever the flags said. Give the
            # worker a few seconds — registration is lazy, and a 2s check
            # produced false negatives.
            # Count only OUR extensions' workers. Chrome always runs its
            # own component extensions (pdf, hangout_services,
            # network_speech_synthesis), so a bare worker count reports
            # success even when the solver never loaded — that false
            # positive is exactly how a non-running solver went unnoticed.
            loaded = 0
            for _ in range(8):
                time.sleep(1)
                loaded = len(_own_extension_workers(_context, extensions))
                if loaded:
                    break

            if loaded:
                _log(f"extension(s) active: {loaded} background worker(s)")
            elif keep_profile():
                # Deliberately NOT rebuilding here any more. The old code
                # deleted the profile to force a --load-extension extension
                # to register, which also destroyed every extension the user
                # had installed by hand — making manual management
                # impossible. Preserving the profile is the whole point of
                # keep_profile(); tell the user how to fix it instead.
                _log(
                    "WARNING: no extension background worker detected. If you "
                    "expected one, install it INTO the profile so it persists:"
                )
                _log("  python manage_ai_mode_extensions.py --install <path>")
                _log("  python manage_ai_mode_extensions.py --open   # add by hand")
                _log(
                    "  (profile preserved — set GOOGLE_AI_MODE_KEEP_PROFILE=false "
                    "to allow the old destructive rebuild)"
                )
            elif not first_run and not _rebuilt_for_extension:
                # Legacy path, opt-in via GOOGLE_AI_MODE_KEEP_PROFILE=false.
                # Measured: an EXISTING profile refuses a --load-extension
                # unpacked extension (0 workers) while a FRESH profile with
                # identical flags loads it (1 worker).
                _log(
                    "extension did not load — rebuilding the automation profile "
                    "once (GOOGLE_AI_MODE_KEEP_PROFILE=false). This DELETES "
                    "manually installed extensions."
                )
                _rebuilt_for_extension = True
                try:
                    _context.close()
                except Exception:
                    pass
                try:
                    _playwright.stop()
                except Exception:
                    pass
                _context = None
                _playwright = None
                # Preserve staged extensions across the rebuild — they are
                # the user's installs, and losing them is the exact failure
                # this whole change set removes.
                _rebuild_profile_preserving_staged(profile)
                return _get_page()  # fresh profile => extension registers
            else:
                _log(
                    "WARNING: extension still not loaded after rebuilding the "
                    f"profile. Check the path really contains manifest.json: "
                    f"{extensions}. CAPTCHAs will need solving by hand."
                )

        return _single_page(_context)


def open_for_signin(timeout_seconds: int = 300) -> bool:
    """Open the automation profile's browser so you can sign in to Google by
    hand, then keep the session for every later run.

    OPTIONAL — not needed in normal operation. Anonymous AI Mode works
    fine on Patchright (verified live: ~5s answer, no CAPTCHA, no sign-in),
    so this exists only as an escape hatch if a particular IP/profile
    starts getting challenged despite the stealth patches. Note that
    signing in may not even be possible: Google refuses OAuth from an
    automation-controlled browser ("This browser or app may not be
    secure"), which is why anonymous-plus-Patchright is the supported path.

    This is deliberately a separate, interactive step rather than something
    a pipeline run does on its own: it needs a human at the keyboard.
    Returns True if the session looks signed in when you're done.
    """
    page = _get_page()
    page.goto("https://accounts.google.com/", wait_until="domcontentloaded", timeout=30_000)
    _log("=" * 68)
    _log("Sign in to Google in the browser window that just opened.")
    _log("The session is saved to the automation profile and reused by every")
    _log("later run, so this is a one-time step.")
    _log(f"Waiting up to {timeout_seconds}s — press Ctrl+C once you're done.")
    _log("=" * 68)

    deadline = time.monotonic() + timeout_seconds
    try:
        while time.monotonic() < deadline:
            time.sleep(5)
            try:
                page.goto(
                    "https://www.google.com/", wait_until="domcontentloaded", timeout=20_000
                )
                if "Sign in" not in page.inner_text("body"):
                    _log("signed in — the profile will reuse this session")
                    return True
            except Exception:
                continue
    except KeyboardInterrupt:
        pass
    _log("still anonymous — AI Mode will likely keep hitting CAPTCHAs")
    return False


def reset_session(reason: str = "") -> bool:
    """Close the browser, wipe the profile, and start clean. Returns True
    if a reset actually happened.

    This is the automatic recovery for a session that is alive but
    persistently failing — every query CAPTCHA'd, refused, timing out, or
    returning an empty answer pane. Restarting the process by hand used to
    be the only cure; this does the same thing in-process.

    What it discards, and why that is the point: Google's rate-limit flag
    travels with the session cookies in the profile, so simply relaunching
    the same profile inherits the block. Deleting the profile forces a
    fresh identity, and the next _get_page() re-seeds it from the real
    Chrome profile so the replacement is warm rather than a cold, even
    more CAPTCHA-prone blank profile.

    Installed extensions SURVIVE: they live in a sibling directory (see
    _staging_root), specifically so a reset cannot cost the user their
    solver extension.

    Capped by max_session_resets(): if the environment is genuinely broken
    (no network, Chrome missing), rebuilding profiles forever would just
    burn time and hide the real error.
    """
    global _session_resets, _consecutive_failures, _consecutive_captchas
    global _last_captcha_at, _last_refusal_at

    if _session_resets >= max_session_resets():
        _log(
            f"session already reset {_session_resets}x — not resetting again. "
            "The failure is probably environmental (network, Chrome, or a "
            "long IP-level block), not the profile."
        )
        return False

    _session_resets += 1
    _log("=" * 68)
    _log(f"AUTO-RECOVERY {_session_resets}/{max_session_resets()}: resetting session")
    if reason:
        _log(f"reason: {reason}")
    _log("closing browser, discarding cookies, rebuilding the profile...")

    shutdown()

    try:
        profile = Path(active_profile_dir())
        _rebuild_profile_preserving_staged(profile)
        _log(f"profile reset: {profile}")
    except Exception as exc:  # noqa: BLE001
        # A locked profile dir must not abort the run: a plain relaunch
        # (already done by shutdown) still clears in-memory state.
        _log(f"could not delete the profile ({type(exc).__name__}) — relaunching only")

    # Clear every failure signal so the fresh session is judged on its own
    # results, not the old one's.
    _consecutive_failures = 0
    _consecutive_captchas = 0
    globals()["_consecutive_transient"] = 0
    globals()["_last_captcha_at"] = 0.0
    globals()["_last_refusal_at"] = 0.0

    cool_off = float(os.getenv("GOOGLE_AI_MODE_RESET_COOLOFF") or 60)
    if cool_off > 0:
        _log(f"cooling off {cool_off:.0f}s before the first query on the new session")
        time.sleep(cool_off)
    _log("=" * 68)
    return True


def session_stats() -> dict:
    """Failure/recovery counters, for logging and tests."""
    return {
        "consecutive_failures": _consecutive_failures,
        "session_resets": _session_resets,
        "max_failures": max_consecutive_failures(),
        "max_resets": max_session_resets(),
    }


def _wait_for_profile_lock_release(profile: Path, timeout: float = 15.0) -> bool:
    """Block until Chrome has released the profile, or `timeout` elapses.

    context.close() returns before the Chrome process has actually exited.
    Relaunching in that window fails with "Opening in existing browser
    session ... the profile is already in use", and because that failure
    left the module with no context, the NEXT query tried the same thing —
    an unbreakable relaunch loop that ate the whole session-reset budget
    without a single successful query.

    Chrome holds `SingletonLock` (a symlink on POSIX, a plain file on
    Windows) for exactly as long as it owns the profile, so its
    disappearance is the signal. On Windows the file sometimes lingers
    after exit, so a still-present lock is not treated as fatal — we just
    stop waiting and let the launch try.
    """
    lock = profile / "SingletonLock"
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if not (lock.exists() or lock.is_symlink()):
            return True
        time.sleep(0.25)
    return False


def shutdown() -> None:
    global _context, _playwright
    with _browser_lock:
        # Capture the profile BEFORE clearing state — active_profile_dir()
        # reads the same env the launch used, so it stays correct, but the
        # wait must happen after close() and before the next launch.
        try:
            profile = Path(active_profile_dir())
        except Exception:  # noqa: BLE001
            profile = None

        try:
            if _context is not None:
                _context.close()
        except Exception:
            pass
        try:
            if _playwright is not None:
                _playwright.stop()
        except Exception:
            pass
        _context = None
        _playwright = None

        # Do not return until Chrome has let go of the profile, or the
        # very next _get_page() races it for the lock and fails.
        if profile is not None:
            _release_profile(profile)
        if profile is not None and profile.exists():
            if not _wait_for_profile_lock_release(profile):
                _log(
                    "profile lock still held after 15s — launching anyway "
                    "(Chrome may be shutting down slowly)"
                )
            # A killed Chrome leaves the crash flag; clear it now so the
            # relaunch does not boot into "Restore pages?".
            _mark_profile_exited_cleanly(profile)


# Google 400s a search URL beyond roughly 8 KB ("Your client has issued a
# malformed or illegal request"). The whole prompt rides in ?q=, and
# percent-encoding inflates it (spaces/newlines/punctuation become %XX), so
# the safe budget on the RAW prompt is far below 8 KB. Measured live: a
# 9,086-char prompt produced a 10,701-char URL and every request failed.
MAX_URL_CHARS = 7800


def ai_mode_url_length(prompt: str) -> int:
    """Length of the URL `prompt` WOULD produce, without truncating or
    warning.

    build_ai_mode_url() is the wrong tool for asking "does this fit?": it
    truncates and logs a warning, so using it to size a prompt spams the
    log with warnings for candidates that were never sent. Callers that
    shrink a prompt to fit (see market_research.discovery.build_draft_prompt)
    measure with this instead.
    """
    q = urllib.parse.quote_plus(prompt)
    return len(f"https://www.google.com/search?q={q}&udm=50&hl=en")


def build_ai_mode_url(prompt: str) -> str:
    """Build the udm=50 URL, truncating the prompt if it would exceed
    Google's URL limit.

    Truncating a prompt is bad, but it beats the alternative: without this
    guard an over-long prompt silently 400s on every single query and the
    caller just sees empty answers with no indication why. The warning makes
    the cause visible, and callers that build growing prompts (see
    market_research.discovery._format_excluded) should cap their own input
    so this never fires.
    """
    q = urllib.parse.quote_plus(prompt)
    url = f"https://www.google.com/search?q={q}&udm=50&hl=en"
    if len(url) <= MAX_URL_CHARS:
        return url

    # Shrink the raw prompt until the encoded URL fits, then warn loudly.
    text = prompt
    while text and len(f"https://www.google.com/search?q={urllib.parse.quote_plus(text)}&udm=50&hl=en") > MAX_URL_CHARS:
        text = text[: int(len(text) * 0.9)]
    _log(
        f"WARNING: prompt too long for a Google search URL "
        f"({len(prompt)} chars -> {len(url)}-char URL, limit {MAX_URL_CHARS}); "
        f"truncated to {len(text)} chars. Shorten the prompt at the call site."
    )
    return f"https://www.google.com/search?q={urllib.parse.quote_plus(text)}&udm=50&hl=en"


def _page_is_captcha(page) -> bool:
    try:
        if "/sorry/" in (page.url or ""):
            return True
    except Exception:
        pass
    try:
        text = page.inner_text("body").lower()
    except Exception:
        return False
    return any(m in text for m in CAPTCHA_MARKERS)


# Google's "Verifying your request ... Continue" interstitial. Not a
# reCAPTCHA, so a solver extension never touches it; it appears when many
# browsers share one IP and means the SESSION is flagged.
_VERIFY_MARKER = "verifying your request"
_VERIFY_ROUNDS_PER_QUERY = 3
_verify_total = 0


def _page_is_verify(page) -> bool:
    try:
        return _VERIFY_MARKER in (page.inner_text("body", timeout=3000) or "").lower()
    except Exception:
        return False


def _try_verify_continue(page, timeout_seconds: float = 20.0) -> bool:
    """Press the interstitial's Continue button once it enables. True when
    the page moved on. The query must still be re-issued afterwards."""
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        try:
            button = page.get_by_role("button", name=re.compile("continue", re.I)).first
            if button.is_visible(timeout=1000) and button.is_enabled(timeout=1000):
                button.click(timeout=3000)
                page.wait_for_timeout(4000)
                return not _page_is_verify(page)
        except Exception:
            pass
        page.wait_for_timeout(2000)
    return False


# Marker AI Mode renders once generation has finished. Far more reliable
# than guessing from answer length: measured, a short answer (280 chars)
# never tripped the old >400-char heuristic and burned the full 90s timeout
# on every single query.
_DONE_MARKERS = ("ai mode response is ready", "response is ready")


def _answer_text_len(page) -> int:
    try:
        return len(page.locator("main").first.inner_text(timeout=1000).strip())
    except Exception:
        return 0


def _answer_is_complete(page) -> bool:
    try:
        body = page.inner_text("body").lower()
    except Exception:
        return False
    return any(m in body for m in _DONE_MARKERS)


def _wait_for_response(page) -> None:
    """Wait for AI Mode to finish generating.

    Three independent completion signals, because no single one is reliable:

    1. The explicit "AI Mode response is ready" marker — the strongest
       signal, and present on normal completions.
    2. The "Thinking..." indicator appearing and then clearing. It often
       never appears at all (measured), so it cannot be relied on alone.
    3. Answer text that has stopped growing across consecutive polls —
       covers responses that render without either marker.

    The previous version required BOTH no-"Thinking" AND >400 characters of
    text, so any answer shorter than that spun until the 90s timeout — which
    was every short reply, costing ~94s per query instead of ~5s.
    """
    deadline = time.monotonic() + RESPONSE_TIMEOUT_MS / 1000
    seen_thinking = False
    last_len = -1
    stable_polls = 0

    while time.monotonic() < deadline:
        if _answer_is_complete(page):
            return

        try:
            thinking = page.get_by_text(re.compile(r"Thinking", re.I)).count()
        except Exception:
            thinking = 0

        if thinking > 0:
            seen_thinking = True
            stable_polls = 0
        else:
            if seen_thinking:
                return  # indicator appeared and has now cleared
            length = _answer_text_len(page)
            # Growing text means it's still streaming; two consecutive
            # unchanged polls with real content means it has settled.
            if length > 0 and length == last_len:
                stable_polls += 1
                if stable_polls >= 2:
                    return
            else:
                stable_polls = 0
            last_len = length

        page.wait_for_timeout(THINKING_POLL_MS)


# AI Mode's own "I won't answer this" page. NOT a block: Google served the
# page normally and the model simply declined THIS prompt. Distinguishing
# it matters because the correct responses are opposite — a refusal should
# be retried with a different/shorter prompt immediately, while a CAPTCHA
# needs a long cool-off. Treating a refusal as a block wastes 5 minutes
# per occurrence and never fixes it (observed live: a run cooling off 300s
# after AI Mode returned "no response available for this search").
_AI_MODE_REFUSAL_MARKERS = (
    "no response available for this search",
    "try asking something else",
    "i can't help with that",
    "i'm not able to help with that",
    "couldn't generate a response",
    "something went wrong, and an ai response wasn't generated",
)


# Google's per-account quota message. NOT a refusal and NOT a CAPTCHA:
# the prompt is fine and the session is fine, there is simply no answer
# budget left for a while.
#
# Undetected, this is silent and expensive. Observed live: after the
# quota hit, ~118 consecutive chunks returned this page, the caller read
# the prompt's own echoed JSON template as an empty-but-valid answer, and
# two whole markets reported "phone 85 -> 85" — no error, no retry, no
# data, and the run looked successful.
_AI_MODE_RATE_LIMIT_MARKERS = (
    "reached the request limit for ai responses",
    "request limit for ai responses",
    "try again in a little while",
)


def is_ai_mode_rate_limited(text: str) -> bool:
    """True when Google says the account is out of AI Mode responses.

    Separate from is_ai_mode_refusal because the remedies are opposite:
    a refusal should be retried immediately with a reworded or smaller
    prompt, whereas this can only be cured by waiting — retrying at once
    just burns queries against a quota that has not reset.
    """
    low = (text or "").strip().lower()
    if not low:
        return False
    # Scan the whole page, not a trailing window. The limit notice is the
    # ANSWER, so it sits between the echoed prompt and whatever UI chrome
    # Google appends ("Feedback", "Share", related searches) -- measured
    # live at 1,288 chars from the end, i.e. outside the old 800-char
    # tail. The miss was silent and expensive: no marker meant no
    # auto-recovery, so the run logged ordinary "empty answer" rounds and
    # kept querying a quota that had not reset.
    #
    # Whole-page scanning is safe here because these markers are Google's
    # own error text, not words a prompt would contain.
    return any(m in low for m in _AI_MODE_RATE_LIMIT_MARKERS)


def is_ai_mode_refusal(text: str) -> bool:
    """True when AI Mode served a page but declined to answer the prompt.

    Deliberately separate from the CAPTCHA path: retrying a refusal
    immediately (ideally with a smaller batch or reworded prompt) can
    succeed, whereas waiting cannot — the prompt, not the rate limit, is
    what Google objected to.
    """
    low = (text or "").strip().lower()
    if not low:
        return False
    # The rendered page echoes our own prompt before AI Mode's actual
    # answer, so a long INPUT prompt (common for chunked verification
    # calls, which send 10+ companies of JSON) makes the whole page long
    # even when the model's own answer is just a one-line refusal. Only
    # checking the last stretch of text — where the real answer/failure
    # message actually lives — catches that case; checking the whole page
    # with a flat length cap does not (verified live: a 4,628-char page
    # whose actual answer was "Something went wrong, and an AI response
    # wasn't generated." failed the old len()>600 gate and was returned as
    # a "real answer" that then failed JSON parsing every time).
    tail = low[-600:]
    return any(m in tail for m in _AI_MODE_REFUSAL_MARKERS)


# Google's own "the generation failed" message, as distinct from the model
# declining the prompt. Treated as transient (retry after a short backoff)
# rather than as a content refusal.
_TRANSIENT_ERROR_MARKER = "something went wrong"

# Consecutive Google-side generation failures. Tracked separately because
# a refusal loops INSIDE ask() via `continue`, so it rarely reaches the
# per-call failure counter — meaning a session stuck in a permanent
# "Something went wrong" state retried forever with a 20s backoff and the
# automatic session reset never fired (observed live: 8 in a row, zero
# AUTO-RECOVERY lines). Rebuilding the session is the one in-process cure,
# so this counter is what triggers it.
_consecutive_transient = 0
_DEFAULT_MAX_TRANSIENT = 3


def max_consecutive_transient() -> int:
    return _env_int("GOOGLE_AI_MODE_MAX_TRANSIENT", _DEFAULT_MAX_TRANSIENT)

_refusal_total = 0
_last_refusal_at = 0.0


def saw_refusal_recently(within_seconds: float = 600.0) -> bool:
    """True if AI Mode declined a prompt within the last `within_seconds`."""
    if not _last_refusal_at:
        return False
    return (time.monotonic() - _last_refusal_at) <= within_seconds


def saw_captcha_recently(within_seconds: float = 600.0) -> bool:
    """True if Google challenged us within the last `within_seconds`.

    This is the signal for "an empty result was probably caused by Google
    blocking us", and it deliberately survives a successful solve. The
    circuit-breaker counter (_consecutive_captchas) cannot be used for
    this: it resets to zero on any solve, so a CAPTCHA that WAS solved
    leaves no trace and blocking becomes indistinguishable from the market
    genuinely running out of companies.

    A solved CAPTCHA still means Google is rate-limiting this IP — the
    solve gets one query through, it does not lift the throttle — so the
    caller should back off rather than conclude there is nothing left to
    find.
    """
    if not _last_captcha_at:
        return False
    return (time.monotonic() - _last_captcha_at) <= within_seconds


def captcha_stats() -> dict:
    """Snapshot of CAPTCHA activity, for logging and diagnosis."""
    return {
        "total": _captcha_total,
        "consecutive": _consecutive_captchas,
        "seconds_since_last": (
            None if not _last_captcha_at else round(time.monotonic() - _last_captcha_at, 1)
        ),
    }


def _wait_for_captcha_clear(page, timeout_seconds: int = CAPTCHA_WAIT_SECONDS) -> bool:
    """Wait for a CAPTCHA to be cleared, polling instead of sleeping blind.

    The old code slept a flat 60s and then continued regardless. Two
    problems, both reported live ("after solving captcha only buffering,
    after 60 second again come captcha"):

      * a solve after 5s still cost the full 60s of waiting, and
      * a solve was never DETECTED, so the caller could not tell success
        from failure and simply re-read the same /sorry/ page.

    Polling fixes both: return True as soon as the page is no longer a
    CAPTCHA (the caller then re-issues the query), or False on timeout.
    Also detects the solver extension clearing it without a human.
    """
    deadline = time.monotonic() + timeout_seconds
    # Say who is expected to act. The old wording ("solve it in the
    # browser window") read as a request for manual intervention, which
    # was misleading: with a solver extension loaded these clear in ~3-7s
    # of a 60s window, far faster than a human could react. The window is
    # only a fallback for a challenge the extension cannot handle.
    solver = "solver extension active" if extension_paths() or staged_extension_paths()         else "no solver extension — solve it in the browser window"
    _log(
        f"CAPTCHA #{_consecutive_captchas} ({solver}) — waiting up to "
        f"{timeout_seconds}s; the query is re-sent automatically"
    )
    announced_solved = False
    while time.monotonic() < deadline:
        try:
            page.wait_for_timeout(2000)
        except Exception:
            return False  # window closed; the outer handler deals with it
        if not _page_is_captcha(page):
            if not announced_solved:
                remaining = max(0, int(deadline - time.monotonic()))
                elapsed = timeout_seconds - remaining
                # Report how long it took, not just the slack left. A
                # few seconds means the extension solved it; tens of
                # seconds means a human probably did.
                _log(
                    f"CAPTCHA cleared after {elapsed}s "
                    f"({'auto-solved' if elapsed <= 15 else 'likely solved by hand'})"
                    " — re-sending the query"
                )
                announced_solved = True
            return True
    _log(f"CAPTCHA still present after {timeout_seconds}s")
    return False


def _extract_answer_text(page) -> str:
    """Return the rendered answer, with link hrefs inlined as "text (url)".

    Link labels are often just "Official Website" with the domain only in
    the href, so plain innerText would silently drop every URL. Rewrite the
    links in place (not on a detached clone — an unattached node falls back
    to textContent and leaks <script> source), read innerText, then restore
    the page so it is left untouched.
    """
    try:
        return page.evaluate(
            """
            () => {
              const main = document.querySelector('main') || document.body;
              const links = Array.from(main.querySelectorAll('a[href^="http"]'))
                .filter(a => !(a.getAttribute('href') || '').includes('google.com'));
              const originals = links.map(a => a.textContent);
              for (const a of links) {
                a.textContent = `${a.textContent.trim()} (${a.getAttribute('href')})`;
              }
              const hidden = main.querySelectorAll('script, style');
              const display = Array.from(hidden).map(el => el.style.display);
              hidden.forEach(el => { el.style.display = 'none'; });
              const text = main.innerText || '';
              hidden.forEach((el, i) => { el.style.display = display[i]; });
              links.forEach((a, i) => { a.textContent = originals[i]; });
              return text;
            }
            """
        )
    except Exception:
        return ""


def generative_backend_enabled() -> bool:
    """True when a generative web backend is available.

    AI Mode (udm=50) is the preferred backend again now that it runs on
    Patchright: the "anonymous AI Mode is instantly CAPTCHA'd" problem that
    got it written off was a plain-Playwright detection issue, not
    something inherent to udm=50 (re-verified live — see module docstring).
    Falls back to the AI Overview extension relay when AI Mode is switched
    off via GOOGLE_AI_MODE_ENABLED.
    """
    if ai_mode_enabled():
        return True
    from vendor_intel.clients.google_ai_scraper import google_ai_enabled

    return bool(google_ai_enabled())


def ask_generative(prompt: str, *, timeout: int = 180) -> str:
    """Ask the active generative web backend one question; return the answer.

    Single entry point for every "ask a free web LLM" call site (market
    understanding, website lookup, field filling), so the backend is
    swappable in one place.

    Prefers Google AI Mode (a real generative model — it CAN follow a
    batched "reply with ONLY a JSON array" instruction, verified live),
    falling back to the AI Overview relay only when AI Mode is disabled.

    NOTE, and it only applies to the AI OVERVIEW fallback: that path is
    search-backed, so a batched "classify these 25 companies, reply as
    JSON" prompt matches no web page and comes back "About 0 results" —
    keep prompts to a single question when running on that backend. AI Mode
    has no such restriction. Returns "" on any failure, which every caller
    already treats as "no answer".
    """
    if ai_mode_enabled():
        return ask(prompt)

    import asyncio

    from vendor_intel.clients.google_ai_scraper import google_ai_ask, google_ai_enabled

    if not google_ai_enabled():
        return ""
    try:
        data = asyncio.run(google_ai_ask(prompt))
        return (data.get("markdown") or "").strip()
    except Exception:
        return ""


def ask(prompt: str, *, retries: int = 1) -> str:
    """Send one prompt to Google AI Mode and return the answer text.

    Returns "" on CAPTCHA, timeout, or an empty answer — callers treat that
    as "no result" rather than retrying forever, so one bad query can't
    stall a long run.
    """
    if not prompt.strip():
        return ""

    global _consecutive_captchas
    if _consecutive_captchas >= _MAX_CONSECUTIVE_CAPTCHAS:
        # Google has flagged the profile: every query gets challenged and
        # hand-solving does not clear it. Rebuilding the session discards
        # the flagged cookies, which is the one in-process cure — so try
        # that before giving up. Previously this returned "" immediately,
        # which meant a fully-blocked session could never trigger the
        # auto-recovery below (it short-circuits before the retry loop).
        if reset_session(
            f"{_consecutive_captchas} consecutive CAPTCHAs — profile is flagged"
        ):
            return ask(prompt, retries=0)
        return ""

    with _browser_lock:  # one query at a time per page
        # Space requests out so a long run doesn't trip Google's
        # "unusual traffic" IP throttle (see _pace_query).
        _pace_query()
        # A solved CAPTCHA must not consume the whole attempt budget: with
        # retries=1 the old loop spent attempt 0 on the CAPTCHA and attempt
        # 1 re-reading the same /sorry/ page, so a successful hand-solve
        # still returned "". Give CAPTCHA rounds their own allowance on top
        # of the caller's retries.
        attempts_left = retries + 1
        captcha_rounds_left = _MAX_CAPTCHA_ROUNDS_PER_QUERY
        verify_rounds_left = _VERIFY_ROUNDS_PER_QUERY
        attempt = 0
        while attempts_left > 0:
            attempts_left -= 1
            attempt += 1
            # Every step below is inside the try: once the browser window is
            # closed or crashes, ANY page call raises, and an escaping
            # exception kills the whole pipeline run. Observed live: a run
            # that had already screened 1,225 companies died on page.goto
            # because only goto's own timeout was being caught.
            try:
                page = _get_page()
                try:
                    page.goto(
                        build_ai_mode_url(prompt),
                        wait_until="domcontentloaded",
                        timeout=30_000,
                    )
                except Exception as exc:
                    # AI Mode client-side-redirects right after load (adds
                    # &sei=), which can interrupt Playwright's own navigation
                    # wait. The page still lands correctly, so that one is
                    # benign; anything else means retry.
                    if "interrupted by another navigation" not in str(exc):
                        raise

                page.wait_for_timeout(1500)
                if _page_is_verify(page):
                    globals()["_verify_total"] = _verify_total + 1
                    globals()["_last_captcha_at"] = time.monotonic()
                    if verify_rounds_left <= 0:
                        _log("'Verifying your request' persists — giving up on this query")
                        attempts_left = 0
                        break
                    verify_rounds_left -= 1
                    # Each recovery re-issues the query without costing an attempt.
                    attempts_left += 1
                    if _try_verify_continue(page):
                        _log("'Verifying your request' cleared with Continue — re-sending query")
                        continue
                    if verify_rounds_left >= 1:
                        # The flag rides on the cookies: drop them and
                        # relaunch the browser on the same profile.
                        _log("'Verifying your request' — clearing cookies and relaunching the browser")
                        try:
                            if _context is not None:
                                _context.clear_cookies()
                        except Exception:
                            pass
                        shutdown()
                        time.sleep(float(os.getenv("GOOGLE_AI_MODE_VERIFY_COOLOFF") or 20))
                        continue
                    if not reset_session("Google 'Verifying your request' page persists"):
                        attempts_left = 0
                        break
                    continue
                if _page_is_captcha(page):
                    _consecutive_captchas += 1
                    # Record the CAPTCHA in a way a later solve won't erase.
                    globals()["_captcha_total"] = _captcha_total + 1
                    globals()["_last_captcha_at"] = time.monotonic()
                    if captcha_rounds_left <= 0:
                        _log("CAPTCHA persists after several solves — giving up on this query")
                        # break, not return: fall through to the shared
                        # failure accounting below so this counts toward a
                        # session reset like any other failed query.
                        attempts_left = 0
                        break
                    captcha_rounds_left -= 1
                    if _wait_for_captcha_clear(page):
                        # Solved. Re-issue the query: after solving, Google
                        # leaves the browser on /sorry/ (or bounces to a
                        # plain search), so without navigating again the
                        # next pass just re-reads the CAPTCHA page and the
                        # solve is wasted. That was the actual cause of
                        # "solve it and nothing happens, then 60s later
                        # another CAPTCHA".
                        _consecutive_captchas = 0
                        attempts_left += 1  # the solve shouldn't cost an attempt
                    continue

                _wait_for_response(page)
                if _page_is_captcha(page):
                    continue

                text = (_extract_answer_text(page) or "").strip()
                if text and is_ai_mode_refusal(text):
                    # Google answered fine, the model just declined THIS
                    # prompt. Record it as a refusal (not a block) so the
                    # caller retries instead of cooling off for minutes,
                    # and don't return the refusal text — it would be
                    # parsed as a failed JSON answer downstream.
                    _consecutive_captchas = 0  # the page loaded; we're not blocked
                    globals()["_refusal_total"] = _refusal_total + 1
                    globals()["_last_refusal_at"] = time.monotonic()
                    globals()["_last_query_at"] = time.monotonic()
                    # Log the TAIL, not the head: the page begins with
                    # Google's nav chrome ("Skip to main content /
                    # Accessibility help / AI Mode / All / Images ..."), so
                    # printing the first 60 chars showed boilerplate and
                    # hid the actual failure message — which made a real
                    # Google-side error look like an unexplained empty
                    # answer in the log.
                    reason = " ".join(text.split())[-90:]
                    _log(
                        f"AI Mode returned no answer (...{reason}) — "
                        "not a block; retrying with a fresh request"
                    )
                    # "Something went wrong, and an AI response wasn't
                    # generated" is a Google-SIDE transient, not a content
                    # refusal: the same prompt usually succeeds shortly
                    # after, but an immediate retry tends to return the
                    # identical error. A short backoff converts a wasted
                    # attempt into a likely success.
                    if _TRANSIENT_ERROR_MARKER in text.lower():
                        globals()["_consecutive_transient"] = _consecutive_transient + 1
                        if _consecutive_transient >= max_consecutive_transient():
                            # Backing off is not clearing it. The session
                            # itself is wedged, so rebuild it (new profile,
                            # fresh cookies) and retry on the clean one.
                            if reset_session(
                                f"{_consecutive_transient} consecutive Google-side "
                                "generation failures — backoff is not helping"
                            ):
                                globals()["_consecutive_transient"] = 0
                                return ask(prompt, retries=0)
                            return ""
                        backoff = float(os.getenv("GOOGLE_AI_MODE_TRANSIENT_BACKOFF") or 20)
                        if backoff > 0:
                            _log(
                                f"transient Google error "
                                f"({_consecutive_transient}/{max_consecutive_transient()}) "
                                f"— waiting {backoff:.0f}s before retry"
                            )
                            time.sleep(backoff)
                            globals()["_last_query_at"] = time.monotonic()
                    continue

                if text:
                    _consecutive_captchas = 0  # a real answer means we're unblocked
                    globals()["_consecutive_failures"] = 0  # session is healthy
                    globals()["_consecutive_transient"] = 0
                    # Restart the pacing clock from the END of this query, so
                    # generation time counts toward the next gap instead of
                    # being added on top of it.
                    globals()["_last_query_at"] = time.monotonic()
                    return text
            except Exception as exc:  # noqa: BLE001
                _log(f"query failed ({type(exc).__name__}: {str(exc)[:120]})")
                if "asyncio loop" in str(exc).lower():
                    # Diagnostic for the greenlet failure: name who
                    # installed a loop on this thread, since sync
                    # Patchright cannot run once one exists.
                    _log(_asyncio_diagnostic())
                if _looks_like_dead_browser(exc):
                    # NOT gated on `attempt < retries`: a broken context
                    # outlives this ask() call, so skipping teardown on the
                    # final attempt handed the same dead context to the
                    # next query. That is how one greenlet error became an
                    # unbroken run of identical failures.
                    _log("browser/driver unusable — tearing down and relaunching")
                    shutdown()  # next _get_page() starts a fresh window

        # Every attempt for this prompt failed. Count it against the
        # session, and once the streak is long enough treat the SESSION as
        # the problem rather than the prompt: rebuild it and retry this
        # same prompt once on the fresh session.
        globals()["_consecutive_failures"] = _consecutive_failures + 1
        _log(
            f"no answer for this query "
            f"({_consecutive_failures}/{max_consecutive_failures()} consecutive failures)"
        )
        if _consecutive_failures >= max_consecutive_failures():
            did_reset = reset_session(
                f"{_consecutive_failures} consecutive queries returned no answer"
            )
            if not did_reset:
                # Cap reached. Zero the streak so the next query does not
                # immediately re-attempt a reset that will be refused
                # again — otherwise the counter climbs forever (6/5, 7/5,
                # ...) and every remaining query pays for a pointless
                # reset attempt and its log noise.
                globals()["_consecutive_failures"] = 0
            if did_reset:
                # Retry the prompt that tripped the reset, so the caller
                # sees a real answer rather than an empty result from the
                # query that happened to be last in the bad streak.
                # retries=0: the fresh session gets one clean shot, and a
                # failure here cannot recurse into another reset.
                return ask(prompt, retries=0)
        return ""


_DEAD_BROWSER_MARKERS = (
    "target page, context or browser has been closed",
    "browser has been closed",
    "target closed",
    "connection closed",
    "browser closed",
    "page crashed",
    # The driver's stdio pipe died (observed as a Node EPIPE traceback).
    # The Python objects still exist but every call on them fails.
    "epipe",
    "broken pipe",
    "pipe closed",
    # The sync greenlet was left inside an asyncio loop, so this context
    # can never be driven again from this thread — reusing it produced an
    # infinite identical-failure loop, one relaunch attempt per query,
    # because nothing marked the context unusable.
    "sync api inside the asyncio loop",
    "please use the async api",
    # A launch that lost the race for the profile lock: the previous
    # Chrome has not exited yet. The remedy is the same (tear down and
    # relaunch), but it needs the lock released first — see shutdown().
    "opening in existing browser session",
    "profile is already in use",
    # A launch that failed outright leaves _context None but _playwright
    # started; treating it as live wedges every later query.
    "browsertype.launch_persistent_context",
)


def _asyncio_diagnostic() -> str:
    """Describe the thread's event-loop state, for the greenlet error.

    Patchright's sync API raises "using Playwright Sync API inside the
    asyncio loop" purely because a loop is installed on the calling
    thread. Reporting WHICH loop and whether it is closed turns that
    otherwise-unactionable error into something traceable.
    """
    import asyncio
    import threading

    bits = [f"thread={threading.current_thread().name}"]
    try:
        running = asyncio.get_running_loop()
        bits.append(f"RUNNING loop={running!r}")
    except RuntimeError:
        bits.append("no running loop")
    try:
        import warnings

        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            current = asyncio.get_event_loop_policy().get_event_loop()
        bits.append(f"installed loop={current!r} closed={current.is_closed()}")
    except Exception as exc:  # noqa: BLE001
        bits.append(f"no installed loop ({type(exc).__name__})")
    return "  asyncio state: " + ", ".join(bits)


def _looks_like_dead_browser(exc: Exception) -> bool:
    """True when the error means the browser/driver is gone rather than the
    page merely misbehaving — those are recoverable only by relaunching.

    Deliberately broad. A false positive costs one browser relaunch; a
    false NEGATIVE keeps a permanently broken context in place and every
    subsequent query fails on it, which is what turned a single greenlet
    error into 5/5 consecutive failures and a burnt session-reset budget.
    """
    return any(m in str(exc).lower() for m in _DEAD_BROWSER_MARKERS)
