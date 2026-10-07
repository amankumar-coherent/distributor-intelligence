#!/usr/bin/env python3
"""Alternate discovery and verification automatically until every market hits 300.

A market's verified count -- the number that actually counts -- only
moves during a verification pass. Discovery alone builds a backlog of
unjudged companies and the completion count sits still, which is exactly
what kept happening: hours of wave discovery with "6 of 26" unchanged
because nothing had judged the new companies yet.

So run the two phases in a loop:

    DISCOVER   supervisor_all in waves of 5, for --discover-minutes
    VERIFY     all markets, resuming from checkpoints (only the backlog)
    repeat

Discovery and verification must never overlap: each market drives one
Chromium profile and a profile takes an exclusive lock, so running both
at once produces about:blank tabs and no results. Every phase change here
fully stops the previous phase and clears stale locks before starting the
next.

Markets at or past 300 verified are skipped by the supervisor itself, so
the loop naturally winds down as markets complete.

Progress is appended to logs_auto_cycle.txt, and a market crossing 300 is
announced there the cycle it happens.

Usage:
    python auto_cycle.py
    python auto_cycle.py --discover-minutes 120 --wave 5
"""
from __future__ import annotations

import argparse
import os
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

import supervisor_all as S  # noqa: E402
import market_batches  # noqa: E402

TARGET = 300
# (slug, market_name) across every registered batch -- drone batch 1,
# drone batch 2, fire batch. Was built here directly from drone_markets/
# drone_markets2; now aliases the shared registry so a new batch does
# not need this file edited too.
ALL = [(slug, market) for slug, market, _scope, _kind in market_batches.ALL_MARKETS]


def log(msg: str) -> None:
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    print(line, flush=True)


def counts() -> dict[str, int]:
    return {slug: S.verified(market) for slug, market in ALL}


def stop_everything() -> None:
    """Kill both phases and clear profile locks.

    An unclean stop leaves SingletonLock behind, and the next phase then
    launches into a profile Chromium still considers owned.
    """
    ps = (
        "Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" "
        "-ErrorAction SilentlyContinue | Where-Object { "
        "$_.CommandLine -like '*supervisor_all*' -or "
        "$_.CommandLine -like '*run_drone*' -or "
        "$_.CommandLine -like '*verify_distributors*' -or "
        "$_.CommandLine -like '*sweep_states*' } | ForEach-Object { "
        "Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"
    )
    subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                   capture_output=True, timeout=120)
    time.sleep(4)
    subprocess.run(["powershell", "-NoProfile", "-Command",
                    "Get-Process chrome,chromium -ErrorAction SilentlyContinue "
                    "| Where-Object { $_.Path -like '*vendor-intel*' } "
                    "| Stop-Process -Force -ErrorAction SilentlyContinue"],
                   capture_output=True, timeout=120)
    time.sleep(2)
    for name in ("SingletonLock", "SingletonCookie", "SingletonSocket"):
        for p in (ROOT / "data").glob(f"*/{name}"):
            try:
                p.unlink()
            except OSError:
                pass


def export_all() -> None:
    """Verifiers read the XLSX, so it must include the newest companies.

    Runs market_batches.py as a subprocess (not imported and called
    in-process) deliberately: discovery.py's product-scope/definition
    registry is per-process, so a fresh subprocess is the simplest way
    to get a clean registration pass, and it also means this can never
    hold a browser profile lock a discovery worker needs.
    """
    r = subprocess.run([sys.executable, str(ROOT / "market_batches.py")],
                       cwd=str(ROOT), capture_output=True, text=True,
                       timeout=900)
    log(f"  export: {(r.stdout or '').strip().splitlines()[-1:] or ['failed']}")


def discover(minutes: int, wave: int) -> None:
    log(f"DISCOVER phase: {minutes} min, waves of {wave}")
    handle = (ROOT / "logs_supervisor_all.txt").open("a", encoding="utf-8",
                                                     errors="replace")
    p = subprocess.Popen(
        [sys.executable, "-u", str(ROOT / "supervisor_all.py"),
         "--interval", "180", "--status-every", "10",
         "--wave", str(wave), "--wave-minutes", "30",
         "--target", str(TARGET)],
        cwd=str(ROOT), stdout=handle, stderr=subprocess.STDOUT,
        stdin=subprocess.DEVNULL)
    time.sleep(minutes * 60)
    p.terminate()
    stop_everything()


def verify() -> None:
    log("VERIFY phase: judging the backlog")
    export_all()
    env = dict(os.environ)
    subprocess.run([sys.executable, "-u", str(ROOT / "verify_all_drones.py")],
                   cwd=str(ROOT), env=env, capture_output=True, timeout=600)
    # verify_all_drones launches detached children; wait for them to drain.
    ps = ("(Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" "
          "-ErrorAction SilentlyContinue | Where-Object { $_.CommandLine "
          "-like '*verify_distributors*' } | Measure-Object).Count")
    waited = 0
    while waited < 7200:
        time.sleep(60)
        waited += 60
        out = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                             capture_output=True, text=True,
                             timeout=90).stdout.strip()
        if out in ("0", ""):
            break
    log(f"  verification drained after {waited // 60} min")
    stop_everything()


def main() -> None:
    global TARGET
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--discover-minutes", type=int, default=120)
    ap.add_argument("--wave", type=int, default=5)
    ap.add_argument("--max-cycles", type=int, default=0,
                    help="0 = until every market reaches target")
    ap.add_argument("--target", type=int, default=None,
                    help="override the module TARGET (default 300)")
    args = ap.parse_args()

    if args.target is not None:
        TARGET = args.target

    log("auto_cycle started - discovery and verification will alternate")
    before = counts()
    done0 = sum(1 for v in before.values() if v >= TARGET)
    log(f"start: {sum(before.values())} verified, {done0} of 25 at {TARGET}")

    cycle = 0
    while True:
        cycle += 1
        if args.max_cycles and cycle > args.max_cycles:
            log("max cycles reached - stopping")
            break
        if all(v >= TARGET for v in counts().values()):
            log("EVERY MARKET HAS REACHED 300 - done")
            break

        log(f"===== cycle {cycle} =====")
        stop_everything()
        discover(args.discover_minutes, args.wave)
        verify()

        after = counts()
        crossed = [s for s, v in after.items()
                   if v >= TARGET > before.get(s, 0)]
        gained = sum(after.values()) - sum(before.values())
        done = sum(1 for v in after.values() if v >= TARGET)
        log(f"cycle {cycle} done: +{gained} verified, "
            f"{done} of 25 markets at {TARGET}")
        for s in crossed:
            log(f"  *** {s} REACHED {TARGET} ({after[s]} verified) ***")
        before = after


if __name__ == "__main__":
    main()
