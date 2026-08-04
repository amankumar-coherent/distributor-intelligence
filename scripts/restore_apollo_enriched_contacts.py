#!/usr/bin/env python3
"""Restore contact name + email + LinkedIn from the first Apollo enrich run."""
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

# First successful Apollo enrich run (2026-07-09) — names, roles, email, LinkedIn.
ENRICHED: dict[str, dict[str, str]] = {
    "abb": {
        "contact_person": "Remi D",
        "contact_role": "CISO (Chief Information Security Officer)",
        "email": "rd@abb.com",
        "linkedin": "http://www.linkedin.com/in/r%c3%a9mi-d-38b996144",
    },
    "belden": {
        "contact_person": "Alexandre Peixoto",
        "contact_role": "Vice President Security",
        "email": "alexandre.peixoto@belden.com",
        "linkedin": "http://www.linkedin.com/in/alexandre-peixoto",
    },
    "cisco systems": {
        "contact_person": "Randa Al-Ashqar",
        "contact_role": "CISO",
        "email": "ral-ashqar@cisco.com",
        "linkedin": "http://www.linkedin.com/in/randa-al-ashqar-657789231",
    },
    "cyolo": {
        "contact_person": "Almog Apirion",
        "contact_role": "CEO & Co-Founder",
        "email": "almog@cyolo.io",
        "linkedin": "http://www.linkedin.com/in/almog-apirion",
    },
    "emerson": {
        "contact_person": "Scott K",
        "contact_role": "Vice President, Chief Information Security Officer",
        "email": "scott.kampwerth@emerson.com",
    },
    "honeywell international inc.": {
        "contact_person": "Davis Arora",
        "contact_role": "VP, Chief Information Security Officer",
        "email": "davinder.arora@honeywell.com",
        "linkedin": "http://www.linkedin.com/in/iamdavisarora",
    },
    "immersive labs": {
        "contact_person": "Phillip Neale",
        "contact_role": "VP of Risk and Security",
        "email": "phillip.neale@immersivelabs.com",
        "linkedin": "http://www.linkedin.com/in/phillip-neale-6483b557",
    },
    "johnson controls": {
        "contact_person": "David Ginn",
        "contact_role": "CISO",
        "email": "david.ginn@jci.com",
        "linkedin": "http://www.linkedin.com/in/david-ginn-0240355",
    },
    "orientindia": {
        "contact_person": "Vaibhav Patkar",
        "contact_role": "VIce President / CISO - CyberSecurity Strateg",
        "email": "vaibhavpatkar@orientindia.net",
        "linkedin": "http://www.linkedin.com/in/vaibhav-patkar-5046515",
    },
    "robert bosch gmbh": {
        "contact_person": "Rakesh Tripathi",
        "contact_role": "Head of Security",
        "email": "tripathi.rakesh@bosch.com",
        "linkedin": "http://www.linkedin.com/in/rakeshtripps",
    },
    "rockwell automation": {
        "contact_person": "Israel Goldberg",
        "contact_role": "CISO",
        "email": "israelg@rockwellautomation.com",
        "linkedin": "http://www.linkedin.com/in/israel-goldberg-6a580a23",
    },
    "schneider electric": {
        "contact_person": "Olivier Blum",
        "contact_role": "Chief Executive Officer",
        "email": "olivier.blum@se.com",
        "linkedin": "http://www.linkedin.com/in/olivier-blum",
    },
    "siemens": {
        "contact_person": "Nick Leghorn",
        "contact_role": "Chief Information Security Officer (CISO)",
        "email": "nick.leghorn@siemens.com",
        "linkedin": "http://www.linkedin.com/in/nickleghorn",
    },
    "accenture": {
        "contact_person": "Kris B",
        "contact_role": "Chief Information Security Officer",
    },
}


def _sync_neon_filtered(market_dir: Path, data: dict, filtered: list[dict]) -> Path:
    ctx = data.get("query_context") or {}
    market = str(ctx.get("industry") or data.get("query") or market_dir.name)
    out = {
        "query": market,
        "query_context": ctx,
        "relevant_companies": filtered,
        "completed_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "filter_note": "Founded + HQ — Apollo enriched restore",
    }
    path = market_dir / "_neon_filtered_export.json"
    path.write_text(json.dumps(out, indent=2, ensure_ascii=False), encoding="utf-8")
    return path


def main() -> None:
    slug = sys.argv[1] if len(sys.argv) > 1 else "cyber_physical_systems_security"
    market_dir = ROOT / "output" / slug
    json_path = find_main_json(market_dir)
    if not json_path:
        raise SystemExit(f"No export JSON in {market_dir}")

    data = json.loads(json_path.read_text(encoding="utf-8"))
    rows = [
        r
        for r in data.get("relevant_companies") or []
        if r.get("is_relevant", True) and has_founded(r) and has_headquarters(r)
    ]

    restored = 0
    for row in rows:
        key = norm(row.get("company") or row.get("brand")).lower()
        snapshot = ENRICHED.get(key)
        if not snapshot:
            continue
        crawl_before = dict(row.get("_crawl") or {})
        merge_apollo_contact(
            row,
            snapshot,
            source="apollo_enriched_restore",
            name_only=False,
            fill_empty_only=False,
        )
        crawl_after = dict(row.get("_crawl") or {})
        changed = any(
            norm(crawl_before.get(f)) != norm(crawl_after.get(f))
            for f in ("contact_person", "contact_role", "email", "linkedin")
        )
        if changed:
            restored += 1
            parts = [f"name={crawl_after.get('contact_person')!r}"]
            if crawl_after.get("email"):
                parts.append(f"email={crawl_after.get('email')}")
            if crawl_after.get("linkedin"):
                parts.append("linkedin=yes")
            print(f"  restored {row.get('company')}: {', '.join(parts)}")

    snapshot_path = market_dir / "apollo_enriched_snapshot.json"
    snapshot_path.write_text(json.dumps(ENRICHED, indent=2, ensure_ascii=False), encoding="utf-8")

    data["apollo_enriched_restore"] = {
        "restored": restored,
        "fields": ["contact_person", "contact_role", "email", "linkedin"],
        "completed_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
    }
    json_path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    filtered = [r for r in rows if has_founded(r) and has_headquarters(r)]
    neon_path = _sync_neon_filtered(market_dir, data, filtered)
    export_stats = export_one(market_dir, suffix="founded_and_hq_v2")
    print(f"Restored {restored} companies.")
    print(f"Excel: {export_stats.get('file')}")
    print(f"Neon JSON: {neon_path}")


if __name__ == "__main__":
    main()
