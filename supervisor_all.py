#!/usr/bin/env python3
"""Keep all 24 markets working toward 400, unattended.

Supersedes overnight_supervisor.py, which only knew the first ten drone
markets. This covers all three groups:

    10 batch-one drone markets   (India, state sweeps)
    15 batch-two drone markets   (India, state sweeps)
     1 Dangjo pepper market      (16 countries, its own sweep)

Handles the three failure modes seen repeatedly in this project, in the
order they cost the most:

  1. A MARKET STOPS. Discovery exits deciding the market is exhausted.
     That verdict has been wrong every time it mattered: Polyurethane
     stopped at 309 and reached 401; defense_drones stopped at 12 and
     reached 208. A stopped market is always restarted.

  2. A MARKET STALLS. It runs but stops accepting companies, because the
     exclusion list has saturated and a broad query has no new names to
     give. The fix is a NARROWER question -- state by state for the
     India markets, which took anti_drone past a Bengaluru-only list and
     drone_software from 145 to 570.

  3. THE BACKEND REFUSES. "AI cannot generate a response", a rate limit,
     or a CAPTCHA wall. These live in the session cookies, so a browser
     restart never clears them -- only a profile reset does. A market
     stuck in that state gets its profile rebuilt and restarts clean.
     Staged captcha-raptor sits OUTSIDE the profile and survives.

Everything underneath checkpoints per batch, so a restart costs nothing.

Usage:
    python supervisor_all.py
    python supervisor_all.py --interval 300 --status-every 6
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

from india_states import INDIA_STATES  # noqa: E402
import run_dangjo  # noqa: E402
import market_batches  # noqa: E402

TARGET = 250
LOGS = ROOT / "logs_drones"
DATA = ROOT / "data"

# (slug, market, runner script, profile prefix) across every registered
# batch -- drone batch 1, drone batch 2, fire batch. Previously built
# here directly from drone_markets/drone_markets2, which meant every new
# batch (fire_markets.py, 25 markets) required editing this file by hand.
# market_batches.py holds the merge once; this now just aliases it.
GROUPS: list[tuple[str, str, str, str]] = list(market_batches.GROUPS)

REFUSAL_MARKERS = (
    "ai cannot generate",
    "can't generate a response",
    "couldn't generate a response",
    "no response available for this search",
    "request limit for ai responses",
    "try again in a little while",
    "unusual traffic",
)


def _ps_count(like: str) -> int:
    ps = (
        "Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" "
        "-ErrorAction SilentlyContinue | "
        f"Where-Object {{ $_.CommandLine -like '{like}' }} | "
        "Measure-Object | Select-Object -ExpandProperty Count"
    )
    try:
        out = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                             capture_output=True, text=True, timeout=45).stdout
        return int(out.strip() or 0)
    except (ValueError, subprocess.SubprocessError):
        return 0


def _kill(like: str) -> None:
    ps = (
        "Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" "
        "-ErrorAction SilentlyContinue | "
        f"Where-Object {{ $_.CommandLine -like '{like}' }} | "
        "ForEach-Object { Stop-Process -Id $_.ProcessId -Force "
        "-ErrorAction SilentlyContinue }"
    )
    try:
        subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                       capture_output=True, timeout=60)
    except subprocess.SubprocessError:
        pass


def _parked() -> set[str]:
    """Slugs the user has taken out of the rotation.

    Read fresh each cycle so a market can be parked or released while
    the supervisor is running.
    """
    f = DATA / "parked_markets.txt"
    if not f.is_file():
        return set()
    try:
        lines = f.read_text(encoding="utf-8").splitlines()
    except OSError:
        return set()
    return {ln.strip() for ln in lines
            if ln.strip() and not ln.strip().startswith("#")}


def _next_unswept(slug: str) -> str | None:
    """A state this market has never been queried on.

    Falls through to individual CITIES once every state is used -- a
    state bundles 2-4 cities into one question and saturates the same
    way a country-level query does, so the fix is the same narrowing
    that took Dangjo's China from 1 company to 78.
    """
    import re
    from india_cities import INDIA_CITIES
    log = LOGS / f"{slug}.log"
    used: set[str] = set()
    if log.is_file():
        try:
            text = log.read_text(encoding="utf-8", errors="replace")
            used = {m.split("(")[0].strip()
                    for m in re.findall(r"coverage: ([^\n]+)", text)}
        except OSError:
            pass
    for state in INDIA_STATES:
        if state.split("(")[0].strip() not in used:
            return state
    for city in INDIA_CITIES:
        if city not in used:
            return city
    return None


def _states_left(slug: str) -> int:
    """How many Indian states this market has never been asked about."""
    import re
    try:
        from india_states import INDIA_STATES
    except ImportError:
        return 0
    every = {s.split("(")[0].strip() for s in INDIA_STATES}
    log = LOGS / f"{slug}.log"
    if not log.is_file():
        return len(every)
    try:
        text = log.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return len(every)
    used = {m.split("(")[0].strip()
            for m in re.findall(r"coverage: ([^\n]+)", text)}
    return len(every - used)


def discovered(market: str) -> int:
    from vendor_intel.pipeline.market_research.state import (
        STATE_DIR, market_slug,
    )
    p = STATE_DIR / f"{market_slug(market)}.json"
    if not p.is_file():
        return 0
    try:
        return len(json.loads(p.read_text(encoding="utf-8"))
                   .get("companies") or [])
    except (json.JSONDecodeError, OSError):
        return 0


def verified(market: str) -> int:
    """Companies that PASSED verification -- the actual deliverable.

    Discovery routinely collects two companies for every one that
    survives the distributor rule, so a discovered-count target stops a
    market at roughly half the companies the user asked for.
    """
    from vendor_intel.pipeline.market_research.state import market_slug
    p = (ROOT / "output" / "market_research"
         / f"{market_slug(market)}_verify_checkpoint.json")
    if not p.is_file():
        return 0
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return 0
    return sum(1 for v in (data.get("verdicts") or []) if v.get("keep"))


def poisoned(log: Path) -> bool:
    """True when the log tail shows a refusal that has not recovered."""
    if not log.is_file():
        return False
    try:
        text = log.read_text(encoding="utf-8", errors="replace")[-4000:].lower()
    except OSError:
        return False
    hits = [m for m in REFUSAL_MARKERS if m in text]
    if not hits:
        return False
    # A marker followed by an accept means it already recovered.
    return text.rfind("accepted #") < max(text.rfind(m) for m in hits)


def reset_profile(prefix: str, slug: str) -> None:
    """Rebuild a profile; staged extensions are a sibling and survive."""
    base = DATA / f"{prefix}{slug}"
    for suffix in ("_chromium", ""):
        p = Path(str(base) + suffix)
        if p.is_dir():
            shutil.rmtree(p, ignore_errors=True)


def launch(slug: str, script: str, state: str | None,
           market: str = "") -> None:
    # Ask for more than the market already HAS, scaled to what it still
    # needs. A flat target (TARGET*2) becomes a ceiling the moment
    # discovered passes it: run_drone_market then reports "already have
    # 606" and exits instantly, silently dropping the market from its
    # wave even though it is far short on verified companies.
    have = discovered(market) if market else 0
    missing = max(TARGET - (verified(market) if market else 0), 0)
    goal = have + max(missing * 2, 40)
    cmd = [sys.executable, "-u", str(ROOT / script),
           "--slug", slug, "--target", str(goal),
           "--batch-size", "20"]
    if state:
        cmd += ["--region", state]
    env = dict(os.environ)
    # 14s per browser: with ~24 running that is still under two queries a
    # second overall, which is what keeps the shared IP alive.
    env["GOOGLE_AI_MODE_QUERY_DELAY"] = "14"
    env["MARKET_RESEARCH_BLOCK_COOLOFF"] = "60"
    env["GOOGLE_AI_MODE_RESET_COOLOFF"] = "120"
    LOGS.mkdir(exist_ok=True)
    handle = (LOGS / f"{slug}.log").open("a", encoding="utf-8",
                                         errors="replace")
    subprocess.Popen(cmd, cwd=str(ROOT), env=env, stdout=handle,
                     stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL)


def launch_dangjo() -> None:
    env = dict(os.environ)
    env["GOOGLE_AI_MODE_QUERY_DELAY"] = "14"
    env["MARKET_RESEARCH_BLOCK_COOLOFF"] = "60"
    env["GOOGLE_AI_MODE_RESET_COOLOFF"] = "120"
    handle = (ROOT / "logs_dangjo_overnight2.txt").open(
        "a", encoding="utf-8", errors="replace")
    subprocess.Popen(
        [sys.executable, "-u", str(ROOT / "run_dangjo_by_country.py"),
         "--passes", "2", "--batch-size", "10"],
        cwd=str(ROOT), env=env, stdout=handle,
        stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--interval", type=int, default=300)
    ap.add_argument("--status-every", type=int, default=6)
    ap.add_argument("--wave", type=int, default=0,
                    help="run only N markets at a time (0 = all at once)")
    ap.add_argument("--wave-minutes", type=int, default=30,
                    help="how long each wave runs before rotating")
    ap.add_argument("--target", type=int, default=None,
                    help="override the module TARGET (default 250)")
    args = ap.parse_args()

    if args.target is not None:
        global TARGET
        TARGET = args.target

    state_idx: dict[str, int] = {}
    last: dict[str, int] = {}
    stalls: dict[str, int] = {}
    cycle = 0

    print(f"supervisor_all started {datetime.now():%Y-%m-%d %H:%M:%S}",
          flush=True)
    print(f"{len(GROUPS)} drone markets + dangjo, target {TARGET}, "
          f"checking every {args.interval}s", flush=True)
    if args.wave:
        print(f"wave mode: {args.wave} markets at a time, rotating every "
              f"{args.wave_minutes} min", flush=True)

    wave_start = time.time()
    wave_pos = 0

    while True:
        cycle += 1
        now = datetime.now().strftime("%H:%M:%S")

        # Which markets may run this cycle. Without a wave, all of them.
        parked = _parked()
        pending = [g for g in GROUPS
                   if verified(g[1]) < TARGET and g[0] not in parked]
        # Highest-yield first: a state never asked about returns new
        # companies, a state already swept mostly returns the exclusion
        # list back. Without this the scheduler runs batch one (all 26
        # states already covered) while batch two sits on ~20 unexplored
        # states each.
        # Closest to target first: a market 1 company short can be
        # finished this cycle, a market 240 short cannot, and a completed
        # market leaves the queue and frees its slot for everyone else.
        # Unexplored states are used only to break ties, so a market with
        # fresh ground still outranks an equally-placed exhausted one.
        pending.sort(key=lambda g: (TARGET - verified(g[1]),
                                    -_states_left(g[0])))
        if args.wave and pending:
            elapsed = (time.time() - wave_start) / 60.0
            if elapsed >= args.wave_minutes:
                wave_pos += args.wave
                wave_start = time.time()
                print(f"[{now}] --- wave rotate -> position {wave_pos} ---",
                      flush=True)
            if wave_pos >= len(pending):
                wave_pos = 0
            active = pending[wave_pos:wave_pos + args.wave]
            if len(active) < args.wave:          # wrap around the queue
                active += pending[:args.wave - len(active)]
            active_slugs = {g[0] for g in active}
            # Anything outside the wave must not keep querying: it is
            # competing for the same IP and profile budget.
            for slug, market, _s, _p in pending:
                if slug not in active_slugs and _ps_count(f"*slug {slug}*"):
                    _kill(f"*slug {slug}*")
        else:
            active = pending

        for slug, market, script, prefix in active:
            have = discovered(market)
            # Target is 300 VERIFIED. Discovery must keep going past 300
            # discovered because verification drops about half.
            if verified(market) >= TARGET:
                continue

            alive = _ps_count(f"*slug {slug}*") > 0
            log = LOGS / f"{slug}.log"

            # 3. Refusal / rate limit: the profile carries it, so rebuild.
            if poisoned(log):
                print(f"[{now}] {slug}: refusal in log -> profile reset "
                      f"(at {have})", flush=True)
                _kill(f"*slug {slug}*")
                time.sleep(3)
                reset_profile(prefix, slug)
                idx = state_idx.get(slug, 0)
                fresh = _next_unswept(slug)
                launch(slug, script,
                      fresh or INDIA_STATES[idx % len(INDIA_STATES)], market)
                state_idx[slug] = idx + 1
                last[slug] = have
                time.sleep(4)
                continue

            # 1. Stopped: always restart.
            if not alive:
                idx = state_idx.get(slug, -1)
                if idx < 0 and not (LOGS / f"{slug}.log").is_file():
                    # Genuinely first run: the broad query is the right
                    # opening move and seeds the exclusion list.
                    state, why = None, "first run (all India)"
                    state_idx[slug] = 0
                elif idx < 0:
                    # Already has history. Skip straight to a state it has
                    # never been asked about -- re-running all-India here
                    # just returns companies it already holds.
                    state = _next_unswept(slug)  # None only if every city is used too
                    why = f"fresh state: {state.split('(')[0].strip()}"
                    state_idx[slug] = 0
                else:
                    state = (_next_unswept(slug)
                            or INDIA_STATES[idx % len(INDIA_STATES)])
                    why = f"state: {state.split('(')[0].strip()}"
                    state_idx[slug] = idx + 1
                print(f"[{now}] {slug}: stopped at {have} -> {why}",
                      flush=True)
                launch(slug, script, state, market)
                last[slug] = have
                time.sleep(4)
                continue

            # 2. Running but not gaining: narrow the question.
            if last.get(slug) == have:
                stalls[slug] = stalls.get(slug, 0) + 1
                if stalls[slug] >= 3:
                    idx = state_idx.get(slug, 0)
                    state = (_next_unswept(slug)
                            or INDIA_STATES[idx % len(INDIA_STATES)])
                    print(f"[{now}] {slug}: stalled at {have} -> state: "
                          f"{state.split('(')[0].strip()}", flush=True)
                    _kill(f"*slug {slug}*")
                    time.sleep(3)
                    launch(slug, script, state, market)
                    state_idx[slug] = idx + 1
                    stalls[slug] = 0
                    time.sleep(4)
            else:
                stalls[slug] = 0
            last[slug] = have

        # Verified, not discovered: every other market reports its
        # verified count, and mixing the two made dangjo read
        # "359/300" when it is 140 verified and already delivered.
        dj = verified(run_dangjo.MARKET)
        # A finished market must stay finished. Without this the next
        # cycle sees it under target and relaunches what was just
        # stopped; "data/dangjo.stopped" is how the user calls it done.
        dj_stopped = (DATA / "dangjo.stopped").exists()
        if dj < TARGET and not dj_stopped and _ps_count("*run_dangjo*") == 0:
            print(f"[{now}] dangjo: stopped at {dj} -> restart", flush=True)
            launch_dangjo()

        if cycle % args.status_every == 0:
            total = 0
            done = 0
            print(f"\n--- {now} ---", flush=True)
            for slug, market, _s, _p in GROUPS:
                d = discovered(market)
                k = verified(market)
                total += d
                if k >= TARGET:
                    done += 1
                print(f"  {slug:22}{k:>6}/{TARGET} verified "
                      f"({d} discovered)", flush=True)
            print(f"  {'dangjo':22}{dj:>6}/{TARGET}", flush=True)
            print(f"  {'TOTAL':22}{total + dj:>6}   "
                  f"({done} at target)\n", flush=True)

        time.sleep(args.interval)


if __name__ == "__main__":
    main()
