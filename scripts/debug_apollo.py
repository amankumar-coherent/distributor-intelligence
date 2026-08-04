#!/usr/bin/env python3
"""Quick Apollo debug for specific companies."""
from __future__ import annotations

import asyncio
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT))

from export_founded_hq_v2 import find_main_json, has_founded, has_headquarters
from vendor_intel.enrichment.apollo_contacts import (
    ApolloContactClient,
    _person_display_name,
    _score_title,
    company_domain,
)
from vendor_intel.placeholders.load_keys import apply_env_overrides


async def main() -> None:
    apply_env_overrides()
    market_dir = ROOT / "output" / "cyber_physical_systems_security"
    data = json.loads(find_main_json(market_dir).read_text(encoding="utf-8"))
    rows = [
        r
        for r in data["relevant_companies"]
        if r.get("is_relevant", True) and has_founded(r) and has_headquarters(r)
    ]
    targets = sys.argv[1:] or [
        "Deloitte",
        "Fortinet",
        "Microsoft Corporation",
        "Arctiq",
        "Cisa",
        "Palo Alto Networks",
    ]

    async with ApolloContactClient() as client:
        for t in targets:
            row = next((r for r in rows if (r.get("company") or "").strip() == t), None)
            if not row:
                print(f"NO ROW: {t}")
                continue
            dom = company_domain(row)
            print(f"\n=== {t} domain={dom!r} website={row.get('website')!r} ===")
            try:
                people = await client.search_people(dom, company_name=t)
                print(f"  search returned {len(people)} people")
                for i, p in enumerate(people[:5]):
                    title = str(p.get("title") or "")
                    print(
                        f"    [{i}] score={_score_title(title)} title={title!r} "
                        f"name={_person_display_name(p)!r} id={p.get('id')}"
                    )
                best = max(people, key=lambda p: _score_title(str(p.get("title") or "")))
                pid = str(best.get("id") or "")
                if pid:
                    enriched = await client.enrich_person(pid, domain=dom)
                    print(
                        f"  enrich id={pid} => name={enriched.get('name') if enriched else None!r} "
                        f"email={bool(enriched and enriched.get('email')) if enriched else False}"
                    )
                result = await client.find_contact(row, enrich_match=True)
                print(f"  find_contact => {result}")
                result2 = await client.find_contact(row, enrich_match=False)
                print(f"  find_contact (no enrich) => {result2}")
            except Exception as exc:
                print(f"  EXCEPTION: {type(exc).__name__}: {exc}")


if __name__ == "__main__":
    asyncio.run(main())
