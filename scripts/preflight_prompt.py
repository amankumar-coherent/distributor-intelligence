#!/usr/bin/env python3
"""Pre-launch check: does every region's discovery prompt fit Google's URL?

discovery.build_draft_prompt shrinks the exclusion list (and then examples /
definition / scope) until the percent-ENCODED URL fits MAX_URL_CHARS minus a
safety margin -- but when the FIXED parts alone are too long (long scope plus
a long country-restriction list) nothing fits, the URL is cut from the end,
the JSON reply template is lost and AI Mode answers in prose: 0 companies
(vip 2026-10-25: 7,819 > 7,800 with a 34-country gate). Character counts are
a poor proxy (a space encodes to 3 chars), so this measures the real URL.

Called by run_fire_10_browsers at start-up (refuses to launch on failure);
also usable by hand:
    python scripts/preflight_prompt.py --batch vip
Exit 0 = every market/region fits, 1 = at least one does not.
"""
from __future__ import annotations

import argparse
import importlib
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "src"))


def check(batch: str, worst_ok: int | None = None) -> list[str]:
    """Problem lines for a batch (empty = all regions fit)."""
    saved = dict(os.environ)
    sys.argv = [sys.argv[0]]
    import run_fire_10_browsers as L
    from vendor_intel.pipeline.market_research import discovery as D
    from vendor_intel.scraping.google_ai_mode import MAX_URL_CHARS, ai_mode_url_length

    # Refuse above the HARD limit (the URL is truncated -> 0 companies); warn
    # inside the safety margin (soft 7,682 and caps 7,606 ran fine there).
    limit = worst_ok or MAX_URL_CHARS
    warn_at = MAX_URL_CHARS - D._URL_SAFETY_MARGIN
    L.select_batch(batch)
    runner = importlib.import_module(Path(L.RUNNER).stem)
    problems = []
    try:
        for slug in L.FIRE_SLUGS:
            pool = L.REGIONS.get(slug, ()) if isinstance(L.REGIONS, dict) else L.REGIONS
            worst, worst_region = 0, ""
            for region in pool or (None,):
                market = runner.configure(slug, region)
                prompt = D.build_draft_prompt(market, excluded=[], count=10, collected=[])
                n = ai_mode_url_length(f"{D._DRAFT_SYSTEM.strip()}\n\n{prompt.strip()}")
                if n > worst:
                    worst, worst_region = n, region
            status = ("TOO LONG" if worst > limit
                      else f"OK (inside the {D._URL_SAFETY_MARGIN}-char safety margin)" if worst > warn_at
                      else "OK")
            print(f"  preflight {slug}: worst URL {worst} (limit {limit}) at {worst_region!r} -> {status}")
            if worst > limit:
                problems.append(f"{slug}: discovery URL {worst} > {limit} at region {worst_region!r} "
                                "-- shorten the scope in its market module")
    finally:
        os.environ.clear()
        os.environ.update(saved)
    return problems


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--batch", required=True)
    args = ap.parse_args()
    probs = check(args.batch)
    for p in probs:
        print("  !!", p)
    sys.exit(1 if probs else 0)


if __name__ == "__main__":
    main()
