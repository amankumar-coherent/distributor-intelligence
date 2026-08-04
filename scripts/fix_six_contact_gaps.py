#!/usr/bin/env python3
"""Fix bad domains and fill the 6 remaining empty contact_person rows via Apollo."""
from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT))

from export_founded_hq_v2 import find_main_json, has_founded, has_headquarters
from vendor_intel.enrichment.gap_fill.gaps import norm

GAP_COMPANIES = [
    "Kaspersky Lab",
    "Microsoft Corporation",
    "Secomea",
    "Level 5 Cyber",
    "ICS Defense",
    "SCADAhacker",
]

DOMAIN_FIXES: dict[str, tuple[str, str]] = {
    "kaspersky lab": ("kaspersky.com", "https://www.kaspersky.com"),
    "microsoft corporation": ("microsoft.com", "https://www.microsoft.com"),
}


def main() -> None:
    slug = sys.argv[1] if len(sys.argv) > 1 else "cyber_physical_systems_security"
    market_dir = ROOT / "output" / slug
    json_path = find_main_json(market_dir)
    data = json.loads(json_path.read_text(encoding="utf-8"))
    rows = data.get("relevant_companies") or []

    fixed = 0
    for row in rows:
        key = norm(row.get("company") or row.get("brand")).lower()
        fix = DOMAIN_FIXES.get(key)
        if not fix:
            continue
        domain, website = fix
        if row.get("domain") != domain:
            row["domain"] = domain
            row["website"] = website
            crawl = dict(row.get("_crawl") or {})
            crawl["website"] = website
            row["_crawl"] = crawl
            fixed += 1
            print(f"  fixed domain: {row.get('company')} -> {domain}")

    if fixed:
        json_path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
        print(f"Updated {fixed} domain(s) in {json_path.name}")

    names_arg = ",".join(GAP_COMPANIES)
    cmd = [
        str(ROOT / ".venv" / "Scripts" / "python.exe"),
        str(ROOT / "apollo_contact_fill.py"),
        slug,
        "--enrich",
        "--contact-bundle",
        "--include-apollo-filled",
        "--companies",
        names_arg,
        "--no-neon",
        "--delay",
        "3",
    ]
    print("\nRunning Apollo fill for 6 gap companies...")
    print(" ", " ".join(cmd[-8:]))
    subprocess.run(cmd, cwd=ROOT, check=False)


if __name__ == "__main__":
    main()
