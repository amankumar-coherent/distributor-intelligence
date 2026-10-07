#!/usr/bin/env python3
"""Keep auto_cycle alive overnight, restarting it if it dies.

auto_cycle alternates discovery and verification correctly, but nothing
watches auto_cycle itself. It died twice today and everything sat idle
until noticed -- once for ~50 minutes, once for ~30. Unattended overnight
that could cost the whole night.

This does one job: check every few minutes that auto_cycle is running,
and relaunch it if not. Because auto_cycle resumes from the same state
files and checkpoints, a restart costs nothing beyond the interrupted
phase.

It also records each market that reaches target, so the morning report
does not depend on reading back through cycle logs.

Usage:
    python overnight_keepalive.py
    python overnight_keepalive.py --check-every 300
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

# (slug, market_name) across every registered batch. Was built here
# directly from drone_markets/drone_markets2; now aliases the shared
# registry so a new batch (e.g. fire_markets.py) needs no edit here.
ALL = [(slug, market) for slug, market, _scope, _kind in market_batches.ALL_MARKETS]
LOG = ROOT / "logs_overnight.txt"


def log(msg: str) -> None:
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}"
    print(line, flush=True)
    with LOG.open("a", encoding="utf-8", errors="replace") as fh:
        fh.write(line + "\n")


def alive(pattern: str) -> bool:
    ps = ("(Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" "
          "-ErrorAction SilentlyContinue | Where-Object { $_.CommandLine "
          f"-like '{pattern}' }} | Measure-Object).Count")
    try:
        out = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                             capture_output=True, text=True, timeout=60).stdout
        return int(out.strip() or 0) > 0
    except (ValueError, subprocess.SubprocessError):
        return False


def start_cycle(minutes: int, wave: int, target: int) -> None:
    handle = (ROOT / "logs_auto_cycle.txt").open("a", encoding="utf-8",
                                                 errors="replace")
    subprocess.Popen(
        [sys.executable, "-u", str(ROOT / "auto_cycle.py"),
         "--discover-minutes", str(minutes), "--wave", str(wave),
         "--target", str(target)],
        cwd=str(ROOT), env=dict(os.environ), stdout=handle,
        stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL)


def snapshot() -> dict[str, int]:
    return {slug: S.verified(market) for slug, market in ALL}


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--check-every", type=int, default=300)
    ap.add_argument("--discover-minutes", type=int, default=45)
    ap.add_argument("--wave", type=int, default=5)
    ap.add_argument("--target", type=int, default=None,
                    help="override the completion target (default %d)"
                    % S.TARGET)
    ap.add_argument("--only", action="append", default=[], metavar="SLUG",
                    help="restrict to this market slug (repeatable)")
    args = ap.parse_args()

    target = args.target if args.target is not None else S.TARGET
    scope = ALL
    if args.only:
        wanted = set(args.only)
        scope = [(s, m) for s, m in ALL if s in wanted]
        missing = wanted - {s for s, _ in scope}
        if missing:
            raise SystemExit(f"unknown slug(s): {sorted(missing)}")

    def snap() -> dict[str, int]:
        return {slug: S.verified(market) for slug, market in scope}

    done = {s for s, v in snap().items() if v >= target}
    log(f"keepalive started - target {target}, scope {len(scope)} "
        f"market(s), {len(done)} already at target, "
        f"{len(scope) - len(done)} to go")

    restarts = 0
    while True:
        if not alive("*auto_cycle*"):
            restarts += 1
            log(f"auto_cycle not running -> restarting (#{restarts})")
            start_cycle(args.discover_minutes, args.wave, target)
            time.sleep(30)

        now = snap()
        for slug, count in now.items():
            if count >= target and slug not in done:
                done.add(slug)
                log(f"*** {slug} REACHED {target} ({count} verified) - "
                    f"{len(done)} of {len(scope)} in scope complete ***")

        if all(v >= target for v in now.values()):
            log("EVERY MARKET IN SCOPE HAS REACHED TARGET - stopping")
            break

        time.sleep(args.check_every)


if __name__ == "__main__":
    main()
