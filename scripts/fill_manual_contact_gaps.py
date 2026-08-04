#!/usr/bin/env python3
"""Fill remaining contact gaps from verified public leadership sources."""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT))

from export_founded_hq_v2 import export_one, find_main_json, has_founded, has_headquarters
from vendor_intel.enrichment.gap_fill.gaps import merge_apollo_contact, norm

# Public leadership pages / LinkedIn — not in Apollo database.
MANUAL_CONTACTS: dict[str, dict[str, str]] = {
    "level 5 cyber": {
        "contact_person": "PJ Maloney",
        "contact_role": "CEO",
        "linkedin": "https://www.linkedin.com/in/maloneypj/",
    },
    "ics defense": {
        "contact_person": "Fevzi Acar",
        "contact_role": "Founder & Owner",
        "linkedin": "https://www.linkedin.com/in/fevzi-acar-01580648",
        "email": "info@icsdefense.net",
    },
    "scadahacker": {
        "contact_person": "Joel Langill",
        "contact_role": "Founder",
        "linkedin": "https://www.linkedin.com/in/joel-langill-scadahacker",
    },
}


def _sync_neon(market_dir: Path, data: dict, filtered: list[dict]) -> Path:
    ctx = data.get("query_context") or {}
    market = str(ctx.get("industry") or data.get("query") or market_dir.name)
    out = {
        "query": market,
        "query_context": ctx,
        "relevant_companies": filtered,
        "completed_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "filter_note": "Founded + HQ — manual leadership gap fill",
    }
    path = market_dir / "_neon_filtered_export.json"
    path.write_text(json.dumps(out, indent=2, ensure_ascii=False), encoding="utf-8")
    return path


def main() -> None:
    slug = sys.argv[1] if len(sys.argv) > 1 else "cyber_physical_systems_security"
    market_dir = ROOT / "output" / slug
    json_path = find_main_json(market_dir)
    data = json.loads(json_path.read_text(encoding="utf-8"))
    rows = [
        r
        for r in data.get("relevant_companies") or []
        if r.get("is_relevant", True) and has_founded(r) and has_headquarters(r)
    ]

    filled = 0
    for row in rows:
        key = norm(row.get("company") or row.get("brand")).lower()
        payload = MANUAL_CONTACTS.get(key)
        if not payload:
            continue
        before = norm((row.get("_crawl") or {}).get("contact_person"))
        if before and "*" not in before:
            continue
        merge_apollo_contact(
            row,
            payload,
            source="leadership_public",
            fill_empty_only=True,
            fields=frozenset({"contact_person", "contact_role", "email", "linkedin"}),
        )
        after = norm((row.get("_crawl") or {}).get("contact_person"))
        if after and after != before:
            filled += 1
            print(f"  filled {row.get('company')}: {after} ({payload.get('contact_role')})")

    data["manual_contact_gap_fill"] = {
        "filled": filled,
        "companies": list(MANUAL_CONTACTS.keys()),
        "completed_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
    }
    json_path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    neon = _sync_neon(market_dir, data, rows)
    excel = export_one(market_dir, suffix="founded_and_hq_v2")
    print(f"\nFilled {filled} manual gaps.")
    print(f"Excel: {excel.get('file')}")
    print(f"Neon JSON: {neon}")


if __name__ == "__main__":
    main()
