#!/usr/bin/env python3
"""Seed the 120 EUROPE adhesives & foam markets (eadh) with every Europe-HQ
company of the delivered GLOBAL final files (user 2026-11-03: "add Europe
companies from my global file as seed; those seeds don't need any step").

Per market:
  1. discovery state  data/market_research/<europe market>.json  <- the seeds
     (so discovery never re-finds them and they count toward the cap 250);
  2. verify checkpoint  output/market_research/<..>_verify_checkpoint.json
     <- a keep=True verdict per seed (the verifier skips them);
  3. status answers  data/status_check_v2/eadh_verified/shard_seed.json
     <- "active" per seed (the final status check skips them; names that
     already read "ABC (acquired by XYZ)" keep that name);
  4. export + verifier pass with nothing to judge -> the verified workbook
     exists even for markets that are already full from seeds alone.
Re-running is safe (merges, never duplicates).

Usage: python scripts/seed_eadh.py [--only eadh001_...]
"""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "src"))
os.chdir(ROOT)

import eadh_markets as E  # noqa: E402
from scripts.status_check_folder import company_key  # noqa: E402
from vendor_intel.pipeline.market_research.state import MarketResearchState, market_slug  # noqa: E402

GLOBAL_FINAL = ROOT / "output" / "final" / "adh" / "adh distributors (status checked)"
SEED_SHARD = ROOT / "data" / "status_check_v2" / "eadh_verified" / "shard_seed.json"
PY = sys.executable
SEED_REASON = "seed: Europe company from the global final list (already verified and status-checked)"


def atomic_write(path: Path, obj) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, path)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--only", default=None)
    ap.add_argument("--no-export", action="store_true")
    args = ap.parse_args()

    shard = json.loads(SEED_SHARD.read_text(encoding="utf-8")) if SEED_SHARD.exists() else {}
    total = 0
    for slug, name, _scope, _kind in E.MARKETS_EADH:
        if args.only and slug != args.only:
            continue
        gslug, _gname = E.GLOBAL_OF[slug]
        src = GLOBAL_FINAL / f"{gslug}_verified_distributors.json"
        rows = json.loads(src.read_text(encoding="utf-8"))["companies"]
        seeds = [dict(r) for r in rows if E.is_europe(r.get("Headquarters"))]

        st = MarketResearchState.load(name)
        before = len(st.companies)
        for i, r in enumerate(seeds, before + 1):
            r["#"] = i
        st.add_batch(seeds)
        st.save()
        added = [c for c in st.companies[before:]]

        ms = market_slug(name)
        ck_path = ROOT / "output" / "market_research" / f"{ms}_verify_checkpoint.json"
        ck = json.loads(ck_path.read_text(encoding="utf-8")) if ck_path.exists() else {"offset": 0, "verdicts": []}
        have = {v.get("Company") for v in ck["verdicts"]}
        for c in st.companies:
            if c.get("Company") in have or c.get("Company") not in {s.get("Company") for s in seeds}:
                continue
            ck["verdicts"].append({"Company": c["Company"], "keep": True, "reason": SEED_REASON,
                                   "confidence": 100, "in_market": True, "is_distributor": True,
                                   "fits_criteria": True, "status": "active", "acquired_by": "", "seed": True})
            shard[company_key(c["Company"], c.get("Website"))] = {
                "Company": c["Company"], "status": "active", "acquired_by": "", "year": "",
                "evidence": "seed from the global final list - status already checked there"}
        ck["offset"] = len(ck["verdicts"])
        atomic_write(ck_path, ck)
        total += len(st.companies)
        print(f"{slug:44} seeds {len(seeds):4} | in state {len(st.companies):4} (added {len(added)})", flush=True)

        if not args.no_export:
            subprocess.run([PY, "market_batches.py", "--only", slug], cwd=str(ROOT), capture_output=True)
            xlsx = ROOT / "output" / "market_research" / f"{ms}.xlsx"
            subprocess.run([PY, "verify_distributors_ai_mode.py", "--file", str(xlsx), "--market-name", name],
                           cwd=str(ROOT), capture_output=True)
    atomic_write(SEED_SHARD, shard)
    print(f"DONE: {total} seeded companies in state, {len(shard)} status answers in {SEED_SHARD}", flush=True)


if __name__ == "__main__":
    main()
