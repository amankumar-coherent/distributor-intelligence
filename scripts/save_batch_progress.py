#!/usr/bin/env python3
"""Make a batch's progress portable (clone on another PC, then resume).

The rolling launcher knows which regions a market has finished from two
places: data/region_coverage/<prefix><slug>.json (portable) and the
"coverage: <region>" lines of logs_<batch>/<slug>.log (logs are gitignored,
so a clone loses them and would re-search every region). This copies the
log-only knowledge into the region_coverage JSON and lists the files a clone
needs for the batch to resume:

  data/market_research/<market>.json          discovery state (companies found)
  data/region_coverage/<prefix><slug>.json    regions already searched
  output/market_research/<market>*.xlsx/json  exports, verify checkpoints, verified lists

Usage:
    python scripts/save_batch_progress.py --batch food            # write coverage, list files
    python scripts/save_batch_progress.py --batch food --git-add  # also stage them
"""
from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "src"))


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--batch", required=True)
    ap.add_argument("--git-add", action="store_true")
    args = ap.parse_args()

    sys.argv = [sys.argv[0]]
    import run_fire_10_browsers as L
    import market_batches as MB
    from vendor_intel.pipeline.market_research.state import market_slug, state_path

    L.select_batch(args.batch)
    files: list[Path] = []
    started = 0
    for slug in L.FIRE_SLUGS:
        market = MB.by_slug(slug)[1]
        covered = L.covered_regions(slug)          # JSON + log lines
        if covered - L.finished_regions(slug):
            L.mark_finished(slug, sorted(covered))
        cov = L._coverage_file(slug)
        st = Path(state_path(market))
        if st.is_file():
            started += 1
            files.append(st)
        if cov.is_file():
            files.append(cov)
        stem = market_slug(market)
        files += [p for p in (ROOT / "output" / "market_research").glob(f"{stem}*")
                  if p.is_file() and p.suffix in (".xlsx", ".json")]
        print(f"{slug:44} regions done {len(covered):3} | found {L.S.discovered(market):4} | "
              f"verified {L.S.verified(market):4}")
    print(f"\n{started}/{len(L.FIRE_SLUGS)} markets started; {len(files)} progress files")
    if args.git_add and files:
        rel = [str(p.relative_to(ROOT)) for p in files]
        for i in range(0, len(rel), 100):
            subprocess.run(["git", "add", "-f", "--"] + rel[i:i + 100], cwd=str(ROOT), check=True)
        print("staged.")


if __name__ == "__main__":
    main()
