#!/usr/bin/env python3
"""Verify every drone market in parallel, each in its own browser profile.

Discovery collected 1,107 Indian companies across ten markets and none
have been judged. Verification is what removes manufacturers, retailers
and companies that merely look like distributors -- it dropped 34 of 92
in Dangjo and 472 across the six original markets, so the discovered
count is not the deliverable.

Parallel for the same reason discovery is: each market has its own
Chromium profile, and a profile takes an exclusive lock, so ten
verifications can run at once without contending. Running them in
sequence would take most of a day.

Each verifier resumes from its own checkpoint, so re-running costs
nothing and only unjudged companies are asked about.

Usage:
    python verify_all_drones.py
    python verify_all_drones.py --only drone_ai --only uav_systems
    python verify_all_drones.py --status
"""
from __future__ import annotations

import argparse
import json
import os
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

import market_batches  # noqa: E402

# Every registered batch (drone batch 1, drone batch 2, fire batch).
# Each batch's profiles use a different prefix, so the verifier has to
# be told which one a slug belongs to -- pointing a verifier at the
# wrong profile would collide with a live browser. Was built here
# directly from drone_markets/drone_markets2; now aliases the shared
# registry so a new batch needs no edit here.
MARKETS = [(slug, market, scope)
           for slug, market, scope, _kind in market_batches.ALL_MARKETS]
_PREFIX = {slug: market_batches.profile_prefix_for(slug)
           for slug, _market, _scope, _kind in market_batches.ALL_MARKETS}
from vendor_intel.pipeline.market_research.state import (  # noqa: E402
    STATE_DIR,
    market_slug,
)

OUT = ROOT / "output" / "market_research"
LOGS = ROOT / "logs_drones"


def status() -> None:
    print(f"  {'market':24}{'discovered':>11}{'judged':>8}{'kept':>7}"
          f"{'pending':>9}")
    td = tj = tk = 0
    for slug, market, _ in MARKETS:
        ms = market_slug(market)
        st = STATE_DIR / f"{ms}.json"
        d = 0
        if st.is_file():
            try:
                d = len(json.loads(st.read_text(encoding="utf-8"))
                        .get("companies") or [])
            except (json.JSONDecodeError, OSError):
                pass
        cp = OUT / f"{ms}_verify_checkpoint.json"
        v = []
        if cp.is_file():
            try:
                v = json.loads(cp.read_text(encoding="utf-8")).get(
                    "verdicts") or []
            except (json.JSONDecodeError, OSError):
                pass
        k = sum(1 for x in v if x.get("keep"))
        td += d
        tj += len(v)
        tk += k
        print(f"  {slug:24}{d:>11}{len(v):>8}{k:>7}{d - len(v):>9}")
    print(f"  {'-'*59}")
    print(f"  {'TOTAL':24}{td:>11}{tj:>8}{tk:>7}{td - tj:>9}")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--only", action="append", default=[], metavar="SLUG")
    ap.add_argument("--status", action="store_true")
    args = ap.parse_args()

    if args.status:
        status()
        return

    queue = [m for m in MARKETS if not args.only or m[0] in args.only]
    if not queue:
        print(f"no match. Valid: {', '.join(m[0] for m in MARKETS)}")
        raise SystemExit(2)

    LOGS.mkdir(exist_ok=True)
    print(f"verifying {len(queue)} market(s) in parallel")
    print("")

    for slug, market, _ in queue:
        ms = market_slug(market)
        xlsx = OUT / f"{ms}.xlsx"
        if not xlsx.is_file():
            print(f"  {slug:24} no export - skipped")
            continue

        env = dict(os.environ)
        # Each market drives its OWN profile, which is what makes ten
        # concurrent verifications possible.
        env["GOOGLE_AI_MODE_PROFILE_DIR"] = str(
            ROOT / "data" / f"{_PREFIX[slug]}{slug}")
        env["GOOGLE_AI_MODE_QUERY_DELAY"] = "12"
        env["GOOGLE_AI_MODE_RESET_COOLOFF"] = "120"

        cmd = [sys.executable, "-u",
               str(ROOT / "verify_distributors_ai_mode.py"),
               "--file", str(xlsx),
               "--market-name", market.replace(" - Distributors Only", "")]

        handle = (LOGS / f"{slug}_verify.log").open(
            "w", encoding="utf-8", errors="replace")
        p = subprocess.Popen(cmd, cwd=str(ROOT), env=env, stdout=handle,
                             stderr=subprocess.STDOUT,
                             stdin=subprocess.DEVNULL)
        print(f"  {slug:24} pid {p.pid} -> logs_drones/{slug}_verify.log")

    print("")
    print("  python verify_all_drones.py --status    # progress")


if __name__ == "__main__":
    main()
