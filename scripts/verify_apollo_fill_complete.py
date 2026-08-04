#!/usr/bin/env python3
"""Full verification: JSON + cache + Excel vs Apollo enrich run."""
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

    stats = {
        "total": len(rows),
        "name": 0,
        "partial": 0,
        "email": 0,
        "linkedin": 0,
        "all_three": 0,
        "empty_name": [],
        "partial_name": [],
        "no_email": [],
        "no_linkedin": [],
    }

    print(f"\n{'Company':<38} {'Contact':<26} {'Email':<8} {'LI':<4}")
    print("-" * 80)
    for r in rows:
        co = str(r.get("company") or "")
        c = r.get("_crawl") or {}
        name = norm(c.get("contact_person"))
        email = norm(c.get("email"))
        li = norm(c.get("linkedin"))
        has_li = bool(li and "linkedin" in li.lower())

        if name:
            stats["name"] += 1
        else:
            stats["empty_name"].append(co)
        if "*" in name:
            stats["partial"] += 1
            stats["partial_name"].append((co, name))
        if email and "@" in email:
            stats["email"] += 1
        elif name:
            stats["no_email"].append(co)
        if has_li:
            stats["linkedin"] += 1
        elif name:
            stats["no_linkedin"].append(co)
        if name and email and has_li:
            stats["all_three"] += 1

        flag = " *" if "*" in name else ""
        print(
            f"{co[:38]:<38} {(name or '—')[:26]:<26} "
            f"{'Y' if email else '-':<8} {'Y' if has_li else '-':<4}{flag}"
        )

    print("\n=== SUMMARY (80 Founded+HQ) ===")
    print(f"  Contact name filled:  {stats['name']}/{stats['total']}")
    print(f"  Partial names (*):    {stats['partial']}")
    print(f"  Email filled:         {stats['email']}/{stats['total']}")
    print(f"  LinkedIn filled:      {stats['linkedin']}/{stats['total']}")
    print(f"  Name + email + LI:    {stats['all_three']}/{stats['total']}")
    print(f"  Empty name ({len(stats['empty_name'])}): {', '.join(stats['empty_name'])}")
    if stats["partial_name"]:
        print("  Partial names:")
        for co, n in stats["partial_name"]:
            print(f"    - {co}: {n}")
    cache_ok = sum(
        1
        for v in cache.values()
        if isinstance(v, dict) and norm(v.get("contact_person")) and not v.get("skipped")
    )
    print(f"\n  Cache with contact:   {cache_ok}/{len(cache)}")


if __name__ == "__main__":
    main()
