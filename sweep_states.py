#!/usr/bin/env python3
"""Sweep one or more drone markets state by state across India.

A market that stops yielding on "India" is almost never out of
companies -- it is out of companies the model will volunteer to a broad
question. anti_drone shows this exactly: 96 companies, 55 of them in
Bengaluru, because only six cities had ever been queried and 22 states
had never been asked at all.

Asking "counter-drone distributors in Gujarat (Ahmedabad, Surat,
Vadodara, Rajkot)" is a different question from "counter-drone
distributors in India, excluding these 96". The same technique took
Dangjo from 38 to 112, and South Korea alone from 1 to 41.

Each state gets its own short run, and the exclusion list carries across
them, so no state re-offers what another already found.

Usage:
    python sweep_states.py --slug anti_drone
    python sweep_states.py --slug anti_drone --slug drone_software
    python sweep_states.py --slug anti_drone --limit 8   # top 8 states
"""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

from drone_markets import by_slug  # noqa: E402
from drone_markets2 import by_slug2  # noqa: E402


def resolve(slug: str) -> tuple[str, str]:
    """(market_name, runner_script) for a slug in either batch."""
    try:
        return by_slug(slug)[1], "run_drone_market.py"
    except KeyError:
        return by_slug2(slug)[1], "run_drone2_market.py"
from india_states import states  # noqa: E402

LOGS = ROOT / "logs_drones"


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


def run_state(slug: str, state: str, *, batch: int) -> int:
    """One state pass. Returns companies gained."""
    market, script = resolve(slug)
    before = discovered(market)

    env = dict(os.environ)
    env["GOOGLE_AI_MODE_QUERY_DELAY"] = "12"
    env["MARKET_RESEARCH_BLOCK_COOLOFF"] = "30"
    env["GOOGLE_AI_MODE_RESET_COOLOFF"] = "60"

    LOGS.mkdir(exist_ok=True)
    handle = (LOGS / f"{slug}.log").open("a", encoding="utf-8",
                                         errors="replace")
    # target = before + batch, so each state run stops once it has taken
    # its share rather than trying to reach 400 from one state.
    cmd = [sys.executable, "-u", str(ROOT / script),
           "--slug", slug, "--target", str(before + batch),
           "--batch-size", str(batch), "--region", state]
    try:
        subprocess.run(cmd, cwd=str(ROOT), env=env, stdout=handle,
                       stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL,
                       timeout=900)
    except subprocess.TimeoutExpired:
        pass
    except Exception as exc:  # noqa: BLE001
        print(f"    ! {type(exc).__name__}: {exc}", flush=True)

    after = discovered(market)
    return after - before


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--slug", action="append", required=True)
    ap.add_argument("--batch", type=int, default=10)
    ap.add_argument("--limit", type=int, default=None,
                    help="only the top N states")
    args = ap.parse_args()

    sweep = states(args.limit)
    for slug in args.slug:
        market, _script = resolve(slug)
        start = discovered(market)
        print("")
        print("=" * 66)
        print(f"{slug}: state sweep across {len(sweep)} states "
              f"(from {start})")
        print("=" * 66, flush=True)

        total = 0
        for state in sweep:
            gained = run_state(slug, state, batch=args.batch)
            total += gained
            mark = "+" if gained else " "
            print(f"  {mark}{gained:<3} {state}", flush=True)
            time.sleep(2)

        end = discovered(market)
        print(f"\n  {slug}: {start} -> {end} (+{total})", flush=True)

    from vendor_intel.pipeline.market_research.export_xlsx import export_xlsx

    for slug in args.slug:
        export_xlsx(resolve(slug)[0])
    print("\nexported")


if __name__ == "__main__":
    main()
