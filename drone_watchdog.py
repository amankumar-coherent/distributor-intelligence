#!/usr/bin/env python3
"""Keep every drone market working toward 400, restarting the ones that stop.

A market that reports "3 consecutive empty batches ... genuinely
exhausted" is usually NOT exhausted. Measured twice in this project:

  * Polyurethane declared exhaustion at 309 and reached 401 once the real
    cause (a prompt bug) was fixed;
  * Dangjo sat at 38 on a global query and reached 112 when the same
    question was asked country by country.

Anti-Drone stopped at 86 with a clean yield curve -- 45 productive rounds
then flat zero, no truncation and no prose. That signature is the
EXCLUSION LIST saturating: with 86 names embedded, a broad global query
has no new famous names left to offer, so it returns nothing even though
the market has more companies in it.

Two strategies, applied in order:

  1. RESTART. A fresh session sometimes resumes on its own, and it costs
     one pass to find out.
  2. GEOGRAPHIC SPLIT. Ask region by region. "Counter-drone distributors
     in Latin America" is a different question from "counter-drone
     distributors worldwide, excluding these 86", and it reaches regional
     firms a global query never surfaces. This is what took Dangjo from
     38 to 112 and South Korea from 1 to 41.

Usage:
    python drone_watchdog.py                 # watch all, restart stalled
    python drone_watchdog.py --once          # one pass, then exit
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

from drone_markets import MARKETS  # noqa: E402

LOGS = ROOT / "logs_drones"
TARGET = 400

# Regions to sweep when a market stalls on the global question. Broad
# enough that each is a real market, narrow enough to be a genuinely
# different question from "worldwide".
REGIONS = (
    # India's drone industry by concentration. A market that saturates on
    # "India" still has plenty left city by city -- the same effect that
    # took Dangjo from 38 to 112 when asked country by country.
    "Bengaluru, Karnataka",
    "Hyderabad, Telangana",
    "Delhi NCR (Delhi, Noida, Gurugram)",
    "Pune, Maharashtra",
    "Chennai, Tamil Nadu",
    "Mumbai, Maharashtra",
    "Ahmedabad and Gujarat",
    "Coimbatore, Tamil Nadu",
    "Kolkata and West Bengal",
    "Jaipur and Rajasthan",
    "Kochi and Kerala",
    "Indore and Madhya Pradesh",
    "Chandigarh, Punjab and Haryana",
    "Lucknow and Uttar Pradesh",
    "Bhubaneswar and eastern India",
    "Nagpur and central India",
)


def discovered(market: str) -> int:
    from vendor_intel.pipeline.market_research.state import (
        STATE_DIR,
        market_slug,
    )

    path = STATE_DIR / f"{market_slug(market)}.json"
    if not path.is_file():
        return 0
    try:
        return len(json.loads(path.read_text(encoding="utf-8"))
                   .get("companies") or [])
    except (json.JSONDecodeError, OSError):
        return 0


def running(slug: str) -> bool:
    """True when this market has a live discovery process."""
    ps = (
        "Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" "
        "-ErrorAction SilentlyContinue | "
        f"Where-Object {{ $_.CommandLine -like '*slug {slug}*' }} | "
        "Measure-Object | Select-Object -ExpandProperty Count"
    )
    try:
        out = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                             capture_output=True, text=True, timeout=30).stdout
        return int(out.strip() or 0) > 0
    except (ValueError, subprocess.SubprocessError):
        return False


def launch(slug: str, *, region: str | None) -> subprocess.Popen:
    """Start one market, optionally scoped to a region."""
    cmd = [sys.executable, "-u", str(ROOT / "run_drone_market.py"),
           "--slug", slug, "--target", str(TARGET), "--batch-size", "20"]
    if region:
        cmd += ["--region", region]

    env = dict(os.environ)
    env["GOOGLE_AI_MODE_QUERY_DELAY"] = "12"
    env["MARKET_RESEARCH_BLOCK_COOLOFF"] = "60"
    env["GOOGLE_AI_MODE_RESET_COOLOFF"] = "120"

    LOGS.mkdir(exist_ok=True)
    log = (LOGS / f"{slug}.log").open("a", encoding="utf-8", errors="replace")
    return subprocess.Popen(cmd, cwd=str(ROOT), env=env, stdout=log,
                            stderr=subprocess.STDOUT,
                            stdin=subprocess.DEVNULL)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--once", action="store_true")
    ap.add_argument("--interval", type=int, default=300)
    args = ap.parse_args()

    # Which region each market should try next. A market that stalls
    # repeatedly works through the list rather than repeating one region.
    next_region: dict[str, int] = {}

    while True:
        for slug, market, _ in MARKETS:
            have = discovered(market)
            if have >= TARGET:
                continue
            if running(slug):
                continue

            idx = next_region.get(slug, -1)
            if idx < 0:
                # First stall: a plain restart, which is enough when the
                # stop was a transient failure rather than saturation.
                region = None
                next_region[slug] = 0
                why = "restart"
            else:
                region = REGIONS[idx % len(REGIONS)]
                next_region[slug] = idx + 1
                why = f"region: {region}"

            print(f"[{time.strftime('%H:%M:%S')}] {slug} stalled at {have}"
                  f"/{TARGET} -> {why}", flush=True)
            launch(slug, region=region)
            time.sleep(5)      # stagger launches

        if args.once:
            return
        time.sleep(args.interval)


if __name__ == "__main__":
    main()
