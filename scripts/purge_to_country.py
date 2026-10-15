#!/usr/bin/env python3
"""Reduce one batch's markets to companies headquartered in ONE country.

Generalises purge_drones_to_india.py (user, 2026-10-14: the 9 geo markets were
searched worldwide, then "I want only India companies, not global").

For every market in the batch:
  * state file: keep rows whose Headquarters ends in the country; renumber "#";
    rebuild seen_keys from the kept names;
  * verify checkpoint: keep only the verdicts of kept companies (the verifier
    resumes by WHICH companies are judged, so already-judged India companies
    are not asked again);
  * everything removed is archived first (<file>.<country>_purge_<time>), so
    the purge is recoverable - this is hours of collected data.

Usage:
    python scripts/purge_to_country.py --batch geo --country India --dry-run
    python scripts/purge_to_country.py --batch geo --country India
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "src"))

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

import run_fire_10_browsers as L  # noqa: E402
import market_batches  # noqa: E402
from vendor_intel.pipeline.market_research.state import STATE_DIR, market_slug, normalise_country  # noqa: E402
from vendor_intel.utils.dedupe import company_dedupe_key  # noqa: E402

OUT = ROOT / "output" / "market_research"


def _norm(name: object) -> str:
    return re.sub(r"[^a-z0-9]", "", str(name or "").lower())


def country_of(row: dict) -> str:
    hq = str(row.get("Headquarters") or "").strip()
    return hq.split(",")[-1].strip() if "," in hq else ""


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--batch", required=True)
    ap.add_argument("--country", required=True)
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    want = normalise_country(args.country)
    stamp = time.strftime("%Y%m%d_%H%M%S")
    L.select_batch(args.batch)
    print(f"  {'market':34}{'found':>7}{'keep':>7}{'remove':>8}{'verdicts kept':>15}{'verified kept':>15}")
    totals = [0, 0, 0, 0]
    for slug in L.FIRE_SLUGS:
        market = market_batches.by_slug(slug)[1]
        st = STATE_DIR / f"{market_slug(market)}.json"
        if not st.is_file():
            print(f"  {slug:34}  (no state yet)")
            continue
        data = json.loads(st.read_text(encoding="utf-8"))
        rows = data.get("companies") or []
        keep = [r for r in rows if normalise_country(country_of(r)) == want]
        drop = [r for r in rows if normalise_country(country_of(r)) != want]
        kept_names = {_norm(r.get("Company")) for r in keep}

        cp = OUT / f"{market_slug(market)}_verify_checkpoint.json"
        ck = json.loads(cp.read_text(encoding="utf-8")) if cp.is_file() else {"offset": 0, "verdicts": []}
        verdicts = ck.get("verdicts") or []
        v_keep = [v for v in verdicts if _norm(v.get("Company")) in kept_names]
        kept_ok = sum(1 for v in v_keep if v.get("keep"))

        totals[0] += len(rows); totals[1] += len(keep); totals[2] += len(drop); totals[3] += kept_ok
        print(f"  {slug:34}{len(rows):>7}{len(keep):>7}{len(drop):>8}{len(v_keep):>9}/{len(verdicts):<5}{kept_ok:>15}")
        if args.dry_run:
            continue

        tag = f"{want}_purge_{stamp}"
        st.with_name(st.name + f".{tag}").write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        if cp.is_file():
            cp.with_name(cp.name + f".{tag}").write_text(
                json.dumps(ck, ensure_ascii=False, indent=2), encoding="utf-8")

        for i, row in enumerate(keep, 1):
            row["#"] = str(i)
        data["companies"] = keep
        data["last_index"] = len(keep)
        data["seen_keys"] = [company_dedupe_key(str(r.get("Company") or "")) for r in keep]
        st.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        if cp.is_file():
            cp.write_text(json.dumps({"offset": len(v_keep), "verdicts": v_keep},
                                     ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"  {'-' * 86}")
    print(f"  {'TOTAL':34}{totals[0]:>7}{totals[1]:>7}{totals[2]:>8}{'':>15}{totals[3]:>15}")
    print("\n(dry run - nothing written)" if args.dry_run else f"\nbackups written next to each file (*.{want}_purge_{stamp})")


if __name__ == "__main__":
    main()
