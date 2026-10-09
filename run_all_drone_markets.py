#!/usr/bin/env python3
"""Launch all ten drone markets in parallel, each in its own Chromium.

Ten separate processes, ten profile directories, ten browsers, each with
its own staged captcha-raptor. Separate profiles are not a nicety: a
Chromium profile takes an exclusive lock, and two processes sharing one
produce about:blank tabs and zero results.

RATE LIMITING. Ten browsers query Google from one IP. A single browser
hit "You've reached the request limit" twice in this project's earlier
runs, so ten will hit it harder. Each process handles that itself -- it
detects the limit, rebuilds its own profile and resumes -- and because
the profiles are independent, one market being throttled does not stop
the other nine. Everything checkpoints per batch, so a killed process
resumes rather than restarts.

Usage:
    python run_all_drone_markets.py                  # all ten, target 400
    python run_all_drone_markets.py --target 200
    python run_all_drone_markets.py --only drone_ai --only uav_systems
    python run_all_drone_markets.py --status         # report, launch nothing
"""
from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

from drone_markets import MARKETS  # noqa: E402

LOGS = ROOT / "logs_drones"


def status() -> None:
    """Discovered and verified counts per market."""
    from vendor_intel.pipeline.market_research.state import (
        STATE_DIR,
        market_slug,
    )

    out = ROOT / "output" / "market_research"
    print(f"  {'market':20}{'discovered':>12}{'verified':>10}")
    total_d = total_v = 0
    for slug, market, _ in MARKETS:
        ms = market_slug(market)
        state = STATE_DIR / f"{ms}.json"
        disc = 0
        if state.is_file():
            try:
                disc = len(json.loads(state.read_text(encoding="utf-8"))
                           .get("companies") or [])
            except (json.JSONDecodeError, OSError):
                pass
        cp = out / f"{ms}_verify_checkpoint.json"
        ver = 0
        if cp.is_file():
            try:
                ver = sum(1 for v in json.loads(cp.read_text(encoding="utf-8"))
                          .get("verdicts") or [] if v.get("keep"))
            except (json.JSONDecodeError, OSError):
                pass
        total_d += disc
        total_v += ver
        print(f"  {slug:20}{disc:>12}{ver:>10}")
    print(f"  {'TOTAL':20}{total_d:>12}{total_v:>10}")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--target", type=int, default=400)
    ap.add_argument("--batch-size", type=int, default=20)
    ap.add_argument("--only", action="append", default=[], metavar="SLUG")
    ap.add_argument("--status", action="store_true")
    ap.add_argument("--query-delay", default="10",
                    help="seconds between queries PER BROWSER (default 10)")
    args = ap.parse_args()

    if args.status:
        status()
        return

    queue = [m for m in MARKETS if not args.only or m[0] in args.only]
    if not queue:
        print(f"no match. Valid slugs: {', '.join(m[0] for m in MARKETS)}")
        raise SystemExit(2)

    LOGS.mkdir(exist_ok=True)

    print(f"launching {len(queue)} market(s) in parallel")
    print(f"target {args.target} verified each, GLOBAL coverage")
    print(f"per-browser query delay: {args.query_delay}s")
    print("")

    procs = []
    for slug, market, _ in queue:
        log = LOGS / f"{slug}.log"
        env_cmd = [
            sys.executable, "-u", str(ROOT / "run_drone_market.py"),
            "--slug", slug,
            "--target", str(args.target),
            "--batch-size", str(args.batch_size),
        ]
        import os

        env = dict(os.environ)
        # Each browser paces itself. Ten browsers at 10s each is one query
        # per second overall, which is what makes the shared IP survivable.
        env["GOOGLE_AI_MODE_QUERY_DELAY"] = str(args.query_delay)
        env["MARKET_RESEARCH_BLOCK_COOLOFF"] = "60"
        env["GOOGLE_AI_MODE_RESET_COOLOFF"] = "120"

        handle = log.open("w", encoding="utf-8", errors="replace")
        p = subprocess.Popen(env_cmd, cwd=str(ROOT), env=env,
                             stdout=handle, stderr=subprocess.STDOUT,
                             stdin=subprocess.DEVNULL)
        procs.append((slug, p))
        print(f"  {slug:20} pid {p.pid}  -> logs_drones/{slug}.log")

    print("")
    print("All launched. They run independently; a rate-limited market")
    print("rebuilds its own profile without affecting the others.")
    print("")
    print("  python run_all_drone_markets.py --status   # counts")


if __name__ == "__main__":
    main()
