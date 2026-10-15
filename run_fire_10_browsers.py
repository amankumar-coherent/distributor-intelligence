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
import json
import os
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
from food_markets import MARKETS_FOOD  # noqa: E402
from wound_markets import GLOBAL_REGIONS, MARKETS_WOUND  # noqa: E402
from water_markets import BRAZIL_REGIONS, MARKETS_BRAZIL_RO  # noqa: E402
from chemical_markets import GLOBAL_REGIONS as CHEM_REGIONS, MARKETS_CHEMICAL  # noqa: E402
from ingredient_markets import GLOBAL_REGIONS as INGR_REGIONS, MARKETS_INGREDIENT  # noqa: E402
from geo_markets import MARKETS_GEO, REGIONS_BY_SLUG as GEO_REGIONS  # noqa: E402
from h2_markets import MARKETS_H2  # noqa: E402
from dc_markets import DC_REGIONS, MARKETS_DC  # noqa: E402
from vendor_intel.pipeline.market_research.state import market_slug  # noqa: E402

_INDIA_REGIONS = tuple(INDIA_STATES) + tuple(INDIA_CITIES)

# --batch picks which batch this launcher drives. Same wave / rolling logic
# for all; the runner script, profile prefix, log folder and the regions
# discovery rotates through (Indian states+cities, or countries) differ.
_BATCHES = {
    "fire": ("run_fire_market.py", "fire_profile_", "logs_fire", MARKETS_FIRE, _INDIA_REGIONS),
    "food": ("run_food_market.py", "food_profile_", "logs_food", MARKETS_FOOD, _INDIA_REGIONS),
    "wound": ("run_wound_market.py", "wound_profile_", "logs_wound", MARKETS_WOUND, GLOBAL_REGIONS),
    "brazil": ("run_brazil_market.py", "brazil_profile_", "logs_brazil", MARKETS_BRAZIL_RO, BRAZIL_REGIONS),
    "chem": ("run_chemical_market.py", "chem_profile_", "logs_chem", MARKETS_CHEMICAL, CHEM_REGIONS),
    "ingr": ("run_ingredient_market.py", "ingr_profile_", "logs_ingr", MARKETS_INGREDIENT, INGR_REGIONS),
    # India-only since 2026-10-14 (user): Indian states then cities, like
    # fire/food. (GEO_REGIONS = the earlier worldwide + focus plan.)
    # States only, no city sweeps (user, 2026-10-14).
    "geo": ("run_geo_market.py", "geo_profile_", "logs_geo", MARKETS_GEO, tuple(INDIA_STATES)),
    # PEM & AEM electrolyzers, India-HQ only, states only (user, 2026-10-14).
    "h2": ("run_h2_market.py", "h2_profile_", "logs_h2", MARKETS_H2, tuple(INDIA_STATES)),
    # Global Data Center Market, worldwide country by country (user, 2026-10-15).
    "dc": ("run_dc_market.py", "dc_profile_", "logs_dc", MARKETS_DC, DC_REGIONS),
}
RUNNER, PROFILE_PREFIX = "", ""
LOGS = ROOT
FIRE_SLUGS: tuple[str, ...] = ()
REGIONS: tuple[str, ...] | dict[str, tuple[str, ...]] = ()  # shared list, or per-market plan


def select_batch(name: str) -> None:
    global RUNNER, PROFILE_PREFIX, LOGS, FIRE_SLUGS, REGIONS
    RUNNER, PROFILE_PREFIX, logs, markets, REGIONS = _BATCHES[name]
    LOGS = ROOT / logs
    LOGS.mkdir(exist_ok=True)
    FIRE_SLUGS = tuple(slug for slug, *_ in markets)


select_batch("fire")

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
    name = f"{PROFILE_PREFIX}{slug}"
    if profile_suffix:
        name += f"_{profile_suffix}"
    dst = ROOT / "data" / f"{name}_extensions"
    if dst.is_dir():
        return
    if _EXTENSION_MASTER.is_dir():
        shutil.copytree(_EXTENSION_MASTER, dst)


def _coverage_file(slug: str) -> Path:
    return ROOT / "data" / "region_coverage" / f"{PROFILE_PREFIX}{slug}.json"


def finished_regions(slug: str) -> set[str]:
    """Regions whose search FINISHED, recorded by the launcher itself.

    Several browsers of one market write the same log file at once, and on
    Windows they overwrite each other's lines, so the "coverage:" lines in
    that log lose most of the regions (Brazil run: 3 of 15 survived, and
    round 2 re-searched four regions instead of moving on). This file has one
    writer (the market's own loop), so nothing is lost."""
    try:
        return set(json.loads(_coverage_file(slug).read_text(encoding="utf-8")))
    except (OSError, ValueError):
        return set()


def mark_finished(slug: str, regions: list[str]) -> None:
    done = finished_regions(slug) | {r.split("(")[0].strip() for r in regions}
    f = _coverage_file(slug)
    f.parent.mkdir(parents=True, exist_ok=True)
    tmp = f.with_suffix(".tmp")
    tmp.write_text(json.dumps(sorted(done), ensure_ascii=False), encoding="utf-8")
    os.replace(tmp, f)


def covered_regions(slug: str) -> set[str]:
    used = finished_regions(slug)
    log_path = LOGS / f"{slug}.log"
    if not log_path.is_file():
        return used
    try:
        text = log_path.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return used
    return used | {m.split("(")[0].strip()
                   for m in re.findall(r"coverage: ([^\n]+)", text)}


def next_region(slug: str) -> str | None:
    found = next_regions(slug, 1)
    return found[0] if found else None


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
    # A batch may give each market its own search plan ({slug: regions},
    # e.g. geo: worldwide + that market's focus countries) instead of one
    # shared country list.
    pool = REGIONS.get(slug, ()) if isinstance(REGIONS, dict) else REGIONS
    for region in pool:
        if len(out) >= n:
            break
        # "Karnataka (Bengaluru, Mysuru)" is logged and matched as "Karnataka".
        if region.split("(")[0].strip() not in used and region not in out:
            out.append(region)
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
        sys.executable, "-u", str(ROOT / RUNNER),
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
        f"$_.CommandLine -match '{Path(RUNNER).stem}.*--slug ({pattern})\\b' "
        "} | ForEach-Object { Stop-Process -Id $_.ProcessId -Force "
        "-ErrorAction SilentlyContinue }"
    )
    match_pattern = f"{Path(RUNNER).stem}.*--slug ({pattern})\\b"
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


# -- rolling mode: N slots, a finished market hands its slot to the next ------

VERIFY_TIMEOUT = 2400          # a verifier still running after this is stuck
WORKERS_PER_MARKET = 1         # browsers per market per round (set by main)
ZERO_ROUNDS_EXHAUSTED = 12     # consecutive regions with nothing new = exhausted
_print_lock = threading.Lock()


def rlog(msg: str) -> None:
    with _print_lock:
        log(f"[{time.strftime('%H:%M:%S')}] {msg}")


def _kill(ps_filter: str) -> None:
    ps = ("Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | "
          f"Where-Object {{ {ps_filter} }} | ForEach-Object {{ "
          "Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }")
    try:
        subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                       capture_output=True, timeout=60)
    except subprocess.SubprocessError:
        pass


def reset_profile(slug: str, reason: str) -> None:
    """Close this market's browsers and archive its profile (cookies and
    all), so the next round starts on a brand-new identity. Archived, not
    deleted, like every other profile reset in this project."""
    base = f"{PROFILE_PREFIX}{slug}"
    # The market's own profile plus its parallel workers' (_w2, _w3, ...),
    # each with or without the _chromium suffix.
    own = re.compile(rf"^{re.escape(base)}(_w\d+)?(_chromium)?$")
    _kill(f"$_.Name -match '^(chrome|chromium)\\.exe$' -and "
          f"$_.CommandLine -match '{re.escape(base)}(_w\\d+)?(_chromium)?([\\s\"]|$)'")
    time.sleep(3)
    archive = ROOT / "data" / "_archived_profiles_rolling"
    archive.mkdir(exist_ok=True)
    stamp = time.strftime("%Y%m%d_%H%M%S")
    moved = 0
    for p in sorted((ROOT / "data").iterdir()):
        if p.is_dir() and own.match(p.name):
            try:
                p.rename(archive / f"{p.name}_{stamp}")
                moved += 1
            except OSError:
                shutil.rmtree(p, ignore_errors=True)
    rlog(f"  {slug}: profile reset ({reason}), {moved} folder(s) archived")


def verify_one(slug: str) -> None:
    """Export and verify ONE market, waiting for its verifier to finish.
    A verifier still running after VERIFY_TIMEOUT is killed and its profile
    reset: it is stuck on a block page, and the next round must not share
    the profile with it."""
    subprocess.run([sys.executable, str(ROOT / "market_batches.py"), "--only", slug],
                   cwd=str(ROOT), capture_output=True, timeout=900)
    subprocess.run([sys.executable, str(ROOT / "verify_all_drones.py"), "--only", slug],
                   cwd=str(ROOT), capture_output=True, timeout=300)
    pattern = "verify_distributors.*" + _slugify_for_match(slug)
    seen, waited = False, 0
    while waited < VERIFY_TIMEOUT:
        n = _count_matching(pattern)
        if n > 0:
            seen = True
        elif seen or waited >= 90:
            return
        time.sleep(15)
        waited += 15
    _kill(f"$_.Name -eq 'python.exe' -and $_.CommandLine -match '{pattern}'")
    reset_profile(slug, f"verifier stuck {VERIFY_TIMEOUT // 60} min")


def market_loop(slug: str, target: int, start_delay: float) -> None:
    """Drive one market to target (or exhaustion) on its own: discover one
    region, verify, repeat. A round that finds nothing new resets the
    profile, since that is what a browser stuck on CAPTCHA / rate-limit
    pages looks like from outside."""
    time.sleep(start_delay)
    market = market_batches.by_slug(slug)[1]
    zero_rounds = 0
    rlog(f"START {slug} (verified {S.verified(market)}/{target}, "
         f"{WORKERS_PER_MARKET} browser(s))")
    while True:
        v = S.verified(market)
        if v >= target:
            rlog(f"DONE {slug}: target reached ({v})")
            return
        regions = next_regions(slug, WORKERS_PER_MARKET)
        if not regions:
            rlog(f"DONE {slug}: every region searched ({v} verified)")
            return
        before = S.discovered(market)
        threads = []
        lasted: dict[str, float] = {}   # region -> seconds its worker ran
        for i, region in enumerate(regions):
            # Worker 0 keeps the plain profile; others get _w2, _w3, ...
            suffix = f"w{i + 1}" if i else None
            stage_extension(slug, suffix)
            def _search(region=region, suffix=suffix):
                t0 = time.time()
                discover_one_region(slug, region, 60, profile_suffix=suffix)
                lasted[region] = time.time() - t0

            t = threading.Thread(target=_search, daemon=True)
            threads.append(t)
            t.start()
            time.sleep(3)  # stagger Chromium launches
        for t in threads:
            t.join()
        # A worker that died in under 90 s (browser failed to start) did not
        # search its region, so it is not marked and will be retried.
        mark_finished(slug, [r for r in regions if lasted.get(r, 0) >= 90])
        gained = S.discovered(market) - before
        stop_browsers_for([slug])
        verify_one(slug)
        v = S.verified(market)
        names = ", ".join(r.split("(")[0].strip() for r in regions)
        rlog(f"  {slug}: {names} +{gained} found, verified {v}/{target}")
        if gained > 0:
            zero_rounds = 0
            continue
        zero_rounds += 1
        reset_profile(slug, "round found nothing new")
        if zero_rounds >= ZERO_ROUNDS_EXHAUSTED:
            rlog(f"DONE {slug}: {zero_rounds} regions in a row found nothing new ({v} verified)")
            return


def finalize(batch: str, slugs: list[str], workers: int) -> Path:
    """Last pipeline step: active-status + ownership check on every verified
    company, via scripts/status_check_folder.py. Closed companies (with
    evidence) are removed, acquired ones renamed "ABC (acquired by XYZ)";
    output is the checked xlsx + a JSON per market and a report (no zip).

    The verified files are copied into output/final/<batch>/<batch>_verified
    under slug names first: the checker reads a folder, and the folder name
    keys its checkpoints, so each batch gets its own and re-running resumes.
    """
    final = ROOT / "output" / "final" / batch
    stage = final / f"{batch}_verified"
    stage.mkdir(parents=True, exist_ok=True)
    copied = 0
    for slug in slugs:
        market = market_batches.by_slug(slug)[1]
        src = ROOT / "output" / "market_research" / f"{market_slug(market)}_verified_distributors.xlsx"
        if src.is_file():
            shutil.copy2(src, stage / f"{slug}_verified_distributors.xlsx")
            copied += 1
    out = final / f"{batch} distributors (status checked)"
    rlog(f"FINALIZE {batch}: status + ownership check on {copied} market file(s) -> {out}")
    env = dict(os.environ, STATUS_WALL_WAITS=os.environ.get("STATUS_WALL_WAITS", "6"))
    # No --zip: the user wants the Excel and JSON files only (2026-10-13).
    subprocess.run([sys.executable, "-u", str(ROOT / "scripts" / "status_check_folder.py"),
                    "--input", str(stage), "--output", str(out),
                    "--workers", str(workers)], cwd=str(ROOT), env=env)
    report = out / "status_check_report.xlsx"
    # The checker builds output only once EVERY company is checked, and writes
    # the report last; a run cut short by Google blocks leaves no fresh report,
    # and re-running resumes it.
    if report.is_file() and report.stat().st_mtime > time.time() - 600:
        rlog(f"FINALIZE {batch}: done -> {out}")
    else:
        rlog(f"FINALIZE {batch}: status check incomplete (blocked) — re-run to resume")
    return out


def run_rolling(slugs: list[str], target: int, slots: int) -> None:
    from concurrent.futures import ThreadPoolExecutor

    todo = [s for s in slugs if S.verified(market_batches.by_slug(s)[1]) < target]
    rlog(f"rolling: {len(todo)} market(s) to run, {slots} at a time, target {target}")
    with ThreadPoolExecutor(max_workers=slots) as pool:
        # Only the first `slots` markets start together; stagger their
        # browser launches. Later markets start the moment a slot frees up.
        for i, slug in enumerate(todo):
            pool.submit(_safe_market_loop, slug, target, 5.0 * i if i < slots else 0.0)


def _safe_market_loop(slug: str, target: int, start_delay: float) -> None:
    """A pool thread's exception is otherwise silent. Log it and retry the
    market (its saved work is kept) up to 3 times before giving the slot up."""
    import traceback

    for attempt in range(1, 4):
        try:
            market_loop(slug, target, start_delay if attempt == 1 else 30.0)
            return
        except Exception:  # noqa: BLE001
            rlog(f"ERROR {slug} (attempt {attempt}/3):\n{traceback.format_exc()}")
    rlog(f"GAVE UP {slug} after 3 errors")


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
    ap.add_argument("--batch", choices=sorted(_BATCHES), default="fire",
                    help="which India batch to run (default fire)")
    ap.add_argument("--rolling", action="store_true",
                    help="keep --wave markets running at once; a finished "
                    "market's slot starts the next one immediately")
    ap.add_argument("--finalize-only", action="store_true",
                    help="skip discovery; run only the final status + "
                    "ownership check and packaging")
    ap.add_argument("--status-workers", type=int, default=10,
                    help="browsers for the final status + ownership check")
    args = ap.parse_args()
    select_batch(args.batch)

    by_slug = market_batches.by_slug
    if args.finalize_only:
        finalize(args.batch, list(FIRE_SLUGS), args.status_workers)
        return
    if args.rolling:
        global WORKERS_PER_MARKET
        WORKERS_PER_MARKET = max(1, args.workers_per_market)
        run_rolling(list(FIRE_SLUGS), args.target, args.wave)
        log("")
        log(f"all {len(FIRE_SLUGS)} {args.batch} markets processed:")
        for slug in FIRE_SLUGS:
            log(f"  {slug:38} verified={S.verified(by_slug(slug)[1]):>4}")
        finalize(args.batch, list(FIRE_SLUGS), args.status_workers)
        return
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
    log(f"all {len(FIRE_SLUGS)} {args.batch} markets processed:")
    for slug in FIRE_SLUGS:
        market = by_slug(slug)[1]
        log(f"  {slug:38} verified={S.verified(market):>4}")


if __name__ == "__main__":
    main()
