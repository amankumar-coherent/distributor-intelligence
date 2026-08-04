#!/usr/bin/env python3
"""Audit contact_person / email / linkedin: JSON vs apollo cache vs snapshot."""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT))

from export_founded_hq_v2 import find_main_json, has_founded, has_headquarters
from vendor_intel.enrichment.gap_fill.gaps import norm


def main() -> None:
    slug = sys.argv[1] if len(sys.argv) > 1 else "cyber_physical_systems_security"
    market = ROOT / "output" / slug
    json_path = find_main_json(market)
    data = json.loads(json_path.read_text(encoding="utf-8"))
    rows = [
        r
        for r in data.get("relevant_companies") or []
        if r.get("is_relevant", True) and has_founded(r) and has_headquarters(r)
    ]
    cache = json.loads((market / "apollo_contact_cache.json").read_text(encoding="utf-8"))
    snapshot = {}
    snap_path = market / "apollo_enriched_snapshot.json"
    if snap_path.is_file():
        snapshot = json.loads(snap_path.read_text(encoding="utf-8"))

    partial = 0
    cache_name_gap = 0
    cache_email_gap = 0
    cache_li_gap = 0
    snap_name_gap = 0
    empty_name = 0

    print(f"\n=== AUDIT: {slug} ({len(rows)} companies) ===\n")
    for r in rows:
        co = str(r.get("company") or "")
        k = norm(co).lower()
        c = r.get("_crawl") or {}
        jp = norm(c.get("contact_person"))
        je = norm(c.get("email"))
        jl = norm(c.get("linkedin"))
        ch = cache.get(k) or {}
        cp = norm(ch.get("contact_person"))
        ce = norm(ch.get("email"))
        cl = norm(ch.get("linkedin"))
        sp = snapshot.get(k) or {}
        sp_p = norm(sp.get("contact_person"))
        sp_e = norm(sp.get("email"))
        sp_l = norm(sp.get("linkedin"))

        flags: list[str] = []
        if not jp:
            empty_name += 1
            flags.append("EMPTY_NAME")
        if "*" in jp:
            partial += 1
            flags.append("PARTIAL_JSON")
        if cp and cp != jp and ("*" in jp or not jp):
            cache_name_gap += 1
            flags.append("CACHE_HAS_NAME")
        if ce and not je:
            cache_email_gap += 1
            flags.append("CACHE_HAS_EMAIL")
        if cl and not jl:
            cache_li_gap += 1
            flags.append("CACHE_HAS_LI")
        if sp_p and sp_p != jp and ("*" in jp or not jp):
            snap_name_gap += 1
            flags.append("SNAP_HAS_NAME")
        if sp_e and not je:
            flags.append("SNAP_HAS_EMAIL")
        if sp_l and not jl:
            flags.append("SNAP_HAS_LI")

        if flags:
            print(f"{co[:40]:40} {', '.join(flags)}")
            print(f"  name  json={jp or '-'}")
            if cp:
                print(f"        cache={cp}")
            if sp_p:
                print(f"        snap={sp_p}")
            print(f"  email json={je or '-'} | cache={ce or '-'} | snap={sp_e or '-'}")
            print(f"  li    json={'yes' if jl else '-'} | cache={'yes' if cl else '-'} | snap={'yes' if sp_l else '-'}")

    print("\n=== SUMMARY ===")
    print(f"  Empty contact name in JSON:     {empty_name}")
    print(f"  Partial names (* in JSON):      {partial}")
    print(f"  Cache has better name:          {cache_name_gap}")
    print(f"  Cache has email, JSON missing:  {cache_email_gap}")
    print(f"  Cache has linkedin, JSON missing:{cache_li_gap}")
    print(f"  Snapshot has better name:       {snap_name_gap}")


if __name__ == "__main__":
    main()
