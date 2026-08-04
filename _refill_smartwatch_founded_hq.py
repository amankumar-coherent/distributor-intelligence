#!/usr/bin/env python3
"""Re-fill Founded + HQ for the Global Smartwatch market, then re-export.

Two passes (both merge-only — existing values are never overwritten):
  1) LinkedIn MCP firmographics (Founded / HQ / cities / about / key brands)
  2) Wikidata fallback (free, no key) for any company still missing Founded/HQ
     — recovers well-known brands whose LinkedIn page omits the founding year
     (Fitbit=2007, Nokia=1865, Sony=1946, LG=1958, Casio=1957, ...).

Safe to Ctrl+C — the JSON is saved after every company, so re-running only
processes the remaining companies.
"""
from __future__ import annotations

import asyncio
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

from vendor_intel.placeholders.load_keys import apply_env_overrides  # noqa: E402
from vendor_intel.enrichment.gap_fill.gaps import has_founded, has_headquarters  # noqa: E402
from vendor_intel.enrichment.linkedin_firmographic_verify import (  # noqa: E402
    verify_firmographics_for_market,
)
from vendor_intel.placeholders import wikidata  # noqa: E402
from export_founded_hq_v2 import export_one, find_main_json  # noqa: E402

MARKET_DIR = ROOT / "output" / "demo" / "global_smartwatch_market_global"


def _domain_of(row: dict) -> str:
    return str(row.get("website") or row.get("domain") or "").strip()


def _clean_company_name(name: str) -> str:
    """Drop common legal suffixes so Wikidata label match works (e.g. 'Fitbit, Inc.' -> 'Fitbit')."""
    n = re.sub(
        r"[,\s]+(inc\.?|incorporated|corp\.?|corporation|co\.?|company|ltd\.?|limited|"
        r"llc|plc|gmbh|ag|sa|s\.a\.|group|holdings?)\b\.?$",
        "",
        name.strip(),
        flags=re.I,
    ).strip(" .,")
    # Repeat once for names with two suffixes ("Casio Computer Co., Ltd").
    n2 = re.sub(
        r"[,\s]+(inc\.?|incorporated|corp\.?|corporation|co\.?|company|ltd\.?|limited|"
        r"llc|plc|gmbh|ag|sa|s\.a\.|group|holdings?)\b\.?$",
        "",
        n,
        flags=re.I,
    ).strip(" .,")
    return n2 or n or name


async def _wikidata_lookup(name: str, domain: str) -> dict | None:
    """Wikidata lookup with a cleaned-name retry."""
    info = await wikidata.lookup_company(name, domain=domain)
    if info and (info.get("founded_year") or info.get("headquarters")):
        return info
    cleaned = _clean_company_name(name)
    if cleaned and cleaned.lower() != name.lower():
        alt = await wikidata.lookup_company(cleaned, domain=domain)
        if alt and (alt.get("founded_year") or alt.get("headquarters")):
            return alt
    return info


async def _wikidata_fallback(rows: list, persist) -> dict:
    """Fill Founded/HQ from Wikidata for rows LinkedIn could not complete."""
    todo = [
        r for r in rows
        if r.get("is_relevant", True) and (not has_founded(r) or not has_headquarters(r))
    ]
    print(f"\n[wikidata] fallback queue={len(todo)} (missing Founded/HQ after LinkedIn)", flush=True)
    filled_f = filled_h = 0
    for i, r in enumerate(todo, 1):
        name = str(r.get("company") or r.get("brand") or "").strip()
        if not name:
            continue
        try:
            info = await _wikidata_lookup(name, _domain_of(r))
        except Exception as exc:
            print(f"  [wikidata] [{i}/{len(todo)}] {name[:40]} ERROR: {exc}", flush=True)
            info = None
        if info:
            crawl = dict(r.get("_crawl") or {})
            added = []
            if info.get("founded_year") and not has_founded(r):
                crawl["founded_year"] = str(info["founded_year"])
                added.append(f"founded={info['founded_year']}")
                filled_f += 1
            if info.get("headquarters") and not has_headquarters(r):
                crawl["headquarters"] = str(info["headquarters"])
                added.append(f"hq={info['headquarters']}")
                filled_h += 1
            if added:
                r["_crawl"] = crawl
                print(f"  [wikidata] [{i}/{len(todo)}] {name[:40]} +{', '.join(added)}", flush=True)
            else:
                print(f"  [wikidata] [{i}/{len(todo)}] {name[:40]} (no new fields)", flush=True)
        else:
            print(f"  [wikidata] [{i}/{len(todo)}] {name[:40]} (not found)", flush=True)
        persist()
    return {"founded_filled": filled_f, "hq_filled": filled_h, "queued": len(todo)}


async def main() -> None:
    apply_env_overrides()

    jp = find_main_json(MARKET_DIR)
    assert jp, f"no main json in {MARKET_DIR}"
    data = json.loads(jp.read_text(encoding="utf-8"))
    rows = data.get("relevant_companies") or []

    def _persist() -> None:
        jp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    # --- Pass 1: LinkedIn MCP (reset attempted flag on rows still missing F/HQ) ---
    reset = 0
    for r in rows:
        if not r.get("is_relevant", True) or (has_founded(r) and has_headquarters(r)):
            continue
        crawl = dict(r.get("_crawl") or {})
        if crawl.pop("linkedin_mcp_attempted", None) is not None:
            reset += 1
        crawl.pop("linkedin_mcp_timeout", None)
        r["_crawl"] = crawl
    print(f"[refill] reset linkedin_mcp_attempted on {reset} rows missing Founded/HQ", flush=True)

    li_stats = await verify_firmographics_for_market(
        rows, concurrent=1, verify_all=True, persist_fn=_persist,
    )
    _persist()
    print("[refill] linkedin stats:", json.dumps(li_stats, indent=2), flush=True)

    # --- Pass 2: Wikidata fallback for anything still missing ---
    wd_stats = await _wikidata_fallback(rows, _persist)
    _persist()
    print("[refill] wikidata stats:", json.dumps(wd_stats, indent=2), flush=True)

    both = sum(
        1 for r in rows
        if r.get("is_relevant", True) and has_founded(r) and has_headquarters(r)
    )
    print(f"\n[refill] now Founded+HQ = {both} / {len(rows)}", flush=True)

    exp = export_one(MARKET_DIR, suffix="founded_and_hq_v2")
    print("[refill] export:", json.dumps(exp, indent=2), flush=True)


if __name__ == "__main__":
    asyncio.run(main())
