#!/usr/bin/env python3
"""Launch discovery for the 25-market Fire Safety batch, 10 markets at a
time, running EACH wave of 10 to completion (250 verified, or fully
exhausted) before moving to the next 10 -- per the user's explicit
instruction, not a round-robin across all 25 at once.

Generalized from run_10_browsers.py (which drove 10 workers across 4
drone markets). This drives 10 CONCURRENT MARKETS at once, one worker
each. A market whose current state/city has run dry moves to the next
unswept one and keeps going, rather than stopping the way the first
version of the drone launcher did.

DISCOVER-THEN-VERIFY, INTERLEAVED: discovery alone can never reach a
VERIFIED target -- verified only moves when a separate verification pass
judges the discovered backlog (this project's standing lesson, learned
repeatedly on the drone batches: "6 of 26 unchanged" for hours because
nothing had judged the new companies). So each round here is: every
worker in the wave does ONE region, then the wave stops, exports, and
verifies together (each market still gets its own Chromium profile for
verification -- one profile per market, never shared with discovery,
since a shared profile produces about:blank tabs). Then the wave resumes
discovery on whichever markets are still short. This repeats until every
market in the wave hits target or exhausts its search space, and ONLY
THEN does the next wave of 10 begin.

Each fire market gets its OWN Chromium profile (data/fire_profile_<slug>)
with its OWN staged captcha-raptor copy.

Usage:
    python run_fire_10_browsers.py
    python run_fire_10_browsers.py --target 250 --wave 10
"""
from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
import threading
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

from india_states import INDIA_STATES  # noqa: E402
from india_cities import INDIA_CITIES  # noqa: E402
import supervisor_all as S  # noqa: E402
import market_batches  # noqa: E402
from fire_markets import MARKETS_FIRE  # noqa: E402
from vendor_intel.pipeline.market_research.state import market_slug  # noqa: E402

LOGS = ROOT / "logs_fire"
LOGS.mkdir(exist_ok=True)

FIRE_SLUGS: tuple[str, ...] = tuple(slug for slug, *_ in MARKETS_FIRE)

BATCH_SIZE = 20
MAX_ROUNDS_PER_WAVE = 200  # safety ceiling, not expected to be hit
_EXTENSION_MASTER = ROOT / "data" / "ai_mode_chrome_profile_extensions"


def log(msg: str) -> None:
    print(msg, flush=True)


def _slugify_for_match(slug: str) -> str:
    """The regex-safe fragment of the --file path verify_all_drones.py
    passes to each verifier, e.g. gas_flame_detection ->
    'gas-flame-detection-market'. Matching on the file path (derived
    the same way verify_all_drones.py derives it, via market_slug) is
    reliable where matching on --market-name is not: market names like
    "Gas & Flame Detection Market - India" contain regex-special
    characters ('&') that silently broke the match, which is what let
    a verifier be mistaken for already-drained while still running.
    """
    market = market_batches.by_slug(slug)[1]
    return re.escape(market_slug(market))


def stage_extension(slug: str, profile_suffix: str | None = None) -> None:
    """Give this market's Chromium profile its own captcha-raptor copy.

    Fire markets have never run before, so there is no existing
    fire_profile_<slug>_extensions sibling to copy from. Stage directly
    from the confirmed-working master copy instead. A suffixed worker
    (fire_profile_<slug>_w2, for a second parallel worker on the same
    market) needs its OWN staged copy too -- run_fire_market.py never
    stages extensions itself, only this launcher does, and the profile
    directory name (which is what needs a matching _extensions sibling)
    includes the suffix.
    """
    name = f"fire_profile_{slug}"
    if profile_suffix:
        name += f"_{profile_suffix}"
    dst = ROOT / "data" / f"{name}_extensions"
    if dst.is_dir():
        return
    if _EXTENSION_MASTER.is_dir():
        shutil.copytree(_EXTENSION_MASTER, dst)


def covered_regions(slug: str) -> set[str]:
    log_path = LOGS / f"{slug}.log"
    if not log_path.is_file():
        return set()
    try:
        text = log_path.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return set()
    return {m.split("(")[0].strip()
            for m in re.findall(r"coverage: ([^\n]+)", text)}


def next_region(slug: str) -> str | None:
    used = covered_regions(slug)
    for state in INDIA_STATES:
        if state.split("(")[0].strip() not in used:
            return state
    for city in INDIA_CITIES:
        if city not in used:
            return city
    return None


def next_regions(slug: str, n: int) -> list[str]:
    """Up to N distinct unswept regions for this market, for running
    multiple workers on the same market at once. Reading covered_regions
    once and picking N from it (rather than calling next_region() N
    times) matters here: next_region() only knows a region is taken once
    its own "coverage: ..." log line has been written, which happens
    AFTER that region's worker finishes -- two workers started together
    would both read the same empty log and pick the identical region."""
    used = covered_regions(slug)
    out: list[str] = []
    for state in INDIA_STATES:
        if len(out) >= n:
            return out
        name = state.split("(")[0].strip()
        if name not in used and state not in out:
            out.append(state)
    for city in INDIA_CITIES:
        if len(out) >= n:
            return out
        if city not in used and city not in out:
            out.append(city)
    return out


def discover_one_region(slug: str, region: str, goal_padding: int,
                        profile_suffix: str | None = None) -> int:
    """ONE discovery call for this market's current region. Returns
    companies gained. Deliberately does not loop internally -- the wave
    controls rotation so it can interleave a verify pass between rounds
    instead of a worker holding its slot until its region list runs out.

    profile_suffix lets a SECOND (or Nth) worker cover this same market
    at once, each in its own isolated Chromium profile via
    run_fire_market.py's own --profile-suffix (never shared -- a shared
    profile produces about:blank tabs). Both workers write to the same
    on-disk state file for this market; that is safe because state.py's
    save() merges by dedupe key under a file lock rather than
    overwriting, which was built and stress-tested specifically for
    this kind of concurrent same-market write.
    """
    market = market_batches.by_slug(slug)[1]
    before = S.discovered(market)
    goal = before + goal_padding

    # Always the SAME log file regardless of profile_suffix: coverage_
    # regions() reads this file to know which regions are already swept,
    # and that has to stay one shared ledger across every worker on this
    # market or two suffixed workers would each think the other's region
    # was still untried.
    log_path = LOGS / f"{slug}.log"
    handle = log_path.open("a", encoding="utf-8", errors="replace")
    cmd = [
        sys.executable, "-u", str(ROOT / "run_fire_market.py"),
        "--slug", slug,
        "--target", str(goal),
        "--batch-size", str(BATCH_SIZE),
        "--region", region,
    ]
    if profile_suffix:
        cmd += ["--profile-suffix", profile_suffix]
    try:
        subprocess.run(cmd, cwd=str(ROOT), stdout=handle,
                       stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL,
                       timeout=1800)
    except subprocess.TimeoutExpired:
        pass
    except Exception as exc:  # noqa: BLE001
        handle.write(f"\n[launcher] worker error: {exc}\n")
    finally:
        handle.close()

    return S.discovered(market) - before


def _count_matching(pattern: str) -> int:
    ps = (
        "(Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" "
        "-ErrorAction SilentlyContinue | Where-Object { $_.CommandLine "
        f"-match '{pattern}' }} | Measure-Object).Count"
    )
    try:
        out = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                             capture_output=True, text=True,
                             timeout=60).stdout.strip()
        return int(out or 0)
    except (ValueError, subprocess.SubprocessError):
        return -1  # unknown -- caller should not assume drained


def stop_browsers_for(slugs: list[str]) -> None:
    """Kill any Chromium still open for these slugs, so verification's
    OWN launch of the same profiles never collides with a discovery
    process that has not exited yet. Polls until the kill is actually
    confirmed rather than a fixed sleep -- a Stop-Process -Force can
    still take longer than a few seconds to fully release a Chromium
    profile lock, and starting a verifier before that happens is what
    produces about:blank tabs sharing the profile with a live worker."""
    if not slugs:
        return
    pattern = "|".join(slugs)
    kill_ps = (
        "Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" "
        "-ErrorAction SilentlyContinue | Where-Object { "
        f"$_.CommandLine -match 'run_fire_market.*--slug ({pattern})\\b' "
        "} | ForEach-Object { Stop-Process -Id $_.ProcessId -Force "
        "-ErrorAction SilentlyContinue }"
    )
    match_pattern = f"run_fire_market.*--slug ({pattern})\\b"
    for attempt in range(10):
        try:
            subprocess.run(["powershell", "-NoProfile", "-Command", kill_ps],
                           capture_output=True, timeout=60)
        except subprocess.SubprocessError:
            pass
        time.sleep(2)
        remaining = _count_matching(match_pattern)
        if remaining == 0:
            break
    else:
        log(f"  !! stop_browsers_for: {slugs} still alive after "
            f"{attempt + 1} kill attempts")
    time.sleep(2)  # let the OS finish releasing the profile lock


def export_and_verify(slugs: list[str]) -> None:
    """Export the wave's markets, then verify them -- each in its own
    profile, so this can run in parallel across the wave the same way
    discovery does.
    """
    log(f"  -- exporting + verifying {slugs}")
    subprocess.run([sys.executable, str(ROOT / "market_batches.py")],
                   cwd=str(ROOT), capture_output=True, timeout=900)

    cmd = [sys.executable, str(ROOT / "verify_all_drones.py")]
    for slug in slugs:
        cmd += ["--only", slug]
    subprocess.run(cmd, cwd=str(ROOT), capture_output=True, timeout=300)

    # verify_all_drones.py launches detached children (one Popen per
    # slug, each spawned by its own subprocess.Popen call, which is not
    # instant); wait for the verifiers scoped to THIS wave's slugs to
    # both APPEAR and then drain before resuming discovery on the same
    # profiles. Checking only for "0 running" without first confirming
    # they ever started was a false-positive trap: a slow-to-launch
    # child (or one whose command line didn't match the regex yet on
    # the very first poll) reads as "already drained", and the next
    # round's discovery then launches straight into a profile a
    # verifier is still using -- the about:blank collision.
    match_pattern = "verify_distributors.*(" + \
        "|".join(_slugify_for_match(s) for s in slugs) + ")"
    seen = False
    waited = 0
    while waited < 1800:
        remaining = _count_matching(match_pattern)
        if remaining > 0:
            seen = True
        elif seen:
            break
        elif waited >= 60:
            # never saw them appear after a full minute -- something
            # about the launch or the match pattern is wrong; stop
            # waiting on a check that may never turn positive.
            log(f"  !! export_and_verify: verifiers for {slugs} never "
                f"observed running after {waited}s -- proceeding anyway")
            break
        time.sleep(10)
        waited += 10
    log(f"  -- verification drained after {waited}s (seen={seen})")


def run_wave(slugs: list[str], target: int, workers_per_market: int = 1) -> None:
    """Run this wave of markets to completion: alternate one discovery
    round per market with a verify pass, until every market in the wave
    hits target or has no region left to search.

    workers_per_market > 1 runs that many CONCURRENT discovery workers
    on the SAME market, each on a different unswept region and its own
    isolated Chromium profile (via --profile-suffix w1, w2, ...). Useful
    for a small wave (e.g. wave 3's 5 markets) where running only 5
    browsers leaves capacity unused; this fills it by covering more
    ground per round instead of leaving 5 of the "10 browsers" idle.
    """
    for slug in slugs:
        stage_extension(slug)
        for i in range(1, workers_per_market):
            stage_extension(slug, profile_suffix=f"w{i + 1}")

    for round_num in range(1, MAX_ROUNDS_PER_WAVE + 1):
        market_of = market_batches.by_slug
        active = [s for s in slugs if S.verified(market_of(s)[1]) < target]
        active = [s for s in active if next_region(s) is not None]
        if not active:
            break

        log(f"  round {round_num}: discovering {active} "
            f"({workers_per_market} worker(s)/market)")
        threads = []
        for slug in active:
            regions = next_regions(slug, workers_per_market)
            for i, region in enumerate(regions):
                # worker 0 uses the plain (unsuffixed) profile -- the
                # one every single-worker round already used, with its
                # extension already staged by the plain stage_extension
                # call above. Only workers 1+ get "w2", "w3", ... to
                # match the suffixes stage_extension() staged for them.
                suffix = f"w{i + 1}" if i > 0 else None
                t = threading.Thread(
                    target=lambda s=slug, r=region, sf=suffix:
                        discover_one_region(s, r, 60, profile_suffix=sf),
                    daemon=True,
                )
                threads.append(t)
                t.start()
                time.sleep(2)  # stagger Chromium launches
        for t in threads:
            t.join()

        stop_browsers_for(active)
        export_and_verify(active)

    for slug in slugs:
        market = market_batches.by_slug(slug)[1]
        v = S.verified(market)
        reason = "target reached" if v >= target else "region search exhausted"
        log(f"  {slug:38} verified={v:>4}/{target}  ({reason})")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--target", type=int, default=250)
    ap.add_argument("--wave", type=int, default=10,
                    help="how many markets run per wave")
    ap.add_argument("--skip-waves", type=int, default=0,
                    help="number of leading waves to skip (already done)")
    ap.add_argument("--workers-per-market", type=int, default=1,
                    help="concurrent discovery workers per market, for "
                    "using spare browser capacity on a small wave")
    args = ap.parse_args()

    by_slug = market_batches.by_slug
    remaining = list(FIRE_SLUGS)
    if args.skip_waves:
        skip_count = args.skip_waves * args.wave
        skipped, remaining = remaining[:skip_count], remaining[skip_count:]
        log(f"skipping {len(skipped)} already-done market(s): {skipped}")

    while remaining:
        wave = remaining[: args.wave]
        remaining = remaining[args.wave:]
        log(f"===== wave: {wave} =====")
        run_wave(wave, args.target, workers_per_market=args.workers_per_market)

    log("")
    log("all 25 fire markets processed:")
    for slug in FIRE_SLUGS:
        market = by_slug(slug)[1]
        log(f"  {slug:38} verified={S.verified(market):>4}")


if __name__ == "__main__":
    main()
