#!/usr/bin/env python3
"""Report how far each of the four remaining markets has got.

Exists as a file rather than inline in the PowerShell watchdog because
`@"..."@` here-strings containing Python are parsed by PowerShell itself
("The 'from' keyword is not supported in this version of the language"),
so the watchdog silently never ran.

Prints one line per market to stderr for the log, and the COUNT of markets
still short of target to stdout, so the caller can test it with a plain
string comparison.

Exit code is 0 when every market has reached target, 1 otherwise, so it
also works as a bare `if` in a shell.
"""
from __future__ import annotations

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

MARKETS = (
    "Global Dry Eyes Treatment Market - Distributors Only",
    "Global Solar Rooftop Market - Distributors Only",
    "Global Polyurethane Market - Distributors Only",
    "Global CPaaS (Communication Platform as a Service) Market - Distributors Only",
)

TARGET = 400
OUT_DIR = ROOT / "output" / "market_research"


def saved(market: str) -> int:
    path = STATE_DIR / f"{market_slug(market)}.json"
    if not path.is_file():
        return 0
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return 0
    return len(data.get("companies") or [])


def main() -> int:
    short = 0
    for market in MARKETS:
        n = saved(market)
        verified = (OUT_DIR / f"{market_slug(market)}_verified_distributors.xlsx")
        flag = "yes" if verified.is_file() else "no"
        print(f"  {n:4}/{TARGET}  verified={flag:3}  {market}", file=sys.stderr)
        if n < TARGET:
            short += 1

    # stdout carries ONLY the count, so the caller can compare it directly.
    print(short)
    return 0 if short == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
