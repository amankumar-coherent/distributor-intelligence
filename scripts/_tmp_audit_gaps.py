"""Audit partial contact names and blank phones in run4 JSON."""
from __future__ import annotations

import json
import re
from pathlib import Path

ROOT = Path(r"C:\Users\aman.kumar\Desktop\coherent\output\run4")

PARTIAL_RE = re.compile(
    r"\*\*\*|has named|He is the|Board by|cybersecurity industry|"
    r"gokhan gokhan|^new$|\b[A-Z][a-z]+\s+[A-Z]\.?$",
    re.I,
)


def crawl(row: dict) -> dict:
    c = row.get("_crawl")
    return c if isinstance(c, dict) else {}


def main() -> None:
    for mid in [
        "global_smartwatch_market_global",
        "global_green_chemical_market_global",
        "global_firewall_market_global",
    ]:
        data = json.loads((ROOT / mid / f"{mid}.json").read_text(encoding="utf-8"))
        rows = data.get("relevant_companies") or []
        blank_name = []
        partial = []
        blank_phone = []
        for r in rows:
            c = crawl(r)
            company = str(r.get("company") or r.get("brand") or "").strip()
            name = str(c.get("contact_person") or "").strip()
            title = str(c.get("contact_title") or c.get("title") or "").strip()
            phone = str(c.get("phone") or c.get("whatsapp") or "").strip()
            if not name:
                blank_name.append(company)
            elif PARTIAL_RE.search(name) or "***" in name:
                partial.append((company, name, title, phone))
            elif len(name.split()) == 1 and len(name) < 15:
                partial.append((company, name, title, phone))
            if not phone:
                blank_phone.append((company, name, title))

        print(f"\n=== {mid} n={len(rows)} ===")
        print(f"blank_name={len(blank_name)} partial={len(partial)} blank_phone={len(blank_phone)}")
        print("-- partial/junk names --")
        for company, name, title, phone in partial:
            print(f"  {company[:42]:42} | {name[:48]:48} | {title[:28]:28} | ph={phone[:20]}")
        print("-- blank phones --")
        for company, name, title in blank_phone:
            print(f"  {company[:42]:42} | {name[:48]:48} | {title[:28]}")
        if rows:
            print("crawl keys:", sorted(crawl(rows[0]).keys()))


if __name__ == "__main__":
    main()
