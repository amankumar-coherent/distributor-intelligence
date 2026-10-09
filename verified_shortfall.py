#!/usr/bin/env python3
"""How many markets are still short of the verified target.

Prints a single integer to stdout so a shell can compare it directly, or
a per-market table with --report.

Counts VERIFIED companies (keep verdicts), not discovered ones: the
discovery target and the verified count differ by the rejection rate,
which ranges from 17% to 65% across these markets.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

from vendor_intel.pipeline.market_research.state import (  # noqa: E402
    STATE_DIR,
    market_slug,
)

OUT = ROOT / "output" / "market_research"

MARKETS = (
    ("Global Silicon Carbide Market - Distributors Only", "Silicon Carbide"),
    ("Global Wearable Glucometer Market - Distributors Only", "Wearable Glucometer"),
    ("Global Dry Eyes Treatment Market - Distributors Only", "Dry Eyes Treatment"),
    ("Global Solar Rooftop Market - Distributors Only", "Solar Rooftop"),
    ("Global Polyurethane Market - Distributors Only", "Polyurethane"),
    ("Global CPaaS (Communication Platform as a Service) Market - Distributors Only", "CPaaS"),
)


def counts(slug: str) -> tuple[int, int]:
    """(discovered, verified) for one market."""
    state = STATE_DIR / f"{slug}.json"
    discovered = 0
    if state.is_file():
        try:
            discovered = len(
                json.loads(state.read_text(encoding="utf-8")).get("companies") or []
            )
        except (json.JSONDecodeError, OSError):
            pass

    checkpoint = OUT / f"{slug}_verify_checkpoint.json"
    verified = 0
    if checkpoint.is_file():
        try:
            verdicts = json.loads(
                checkpoint.read_text(encoding="utf-8")).get("verdicts") or []
            verified = sum(1 for v in verdicts if v.get("keep"))
        except (json.JSONDecodeError, OSError):
            pass
    return discovered, verified


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--target", type=int, default=400)
    ap.add_argument("--report", action="store_true")
    args = ap.parse_args()

    short = 0
    rows = []
    for market, label in MARKETS:
        discovered, verified = counts(market_slug(market))
        if verified < args.target:
            short += 1
        rows.append((label, discovered, verified, args.target - verified))

    if args.report:
        total = sum(r[2] for r in rows)
        print(f"  {'market':21}{'discovered':>12}{'verified':>10}{'short':>8}")
        for label, discovered, verified, gap in rows:
            print(f"  {label:21}{discovered:>12}{verified:>10}"
                  f"{max(0, gap):>8}")
        print(f"  {'TOTAL':21}{'':>12}{total:>10}")
        return 0

    print(short)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
