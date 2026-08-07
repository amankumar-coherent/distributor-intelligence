#!/usr/bin/env python3
"""Clean garbage Contact Person values and summarize ADC contact fill."""
from __future__ import annotations

import json
import re
from collections import Counter
from pathlib import Path

from fill_adc_contact_role_linkedin import (
    load_rows,
    write_xlsx,
    looks_person,
    empty,
    KNOWN,
)

XLSX = Path(__file__).resolve().parents[1] / (
    "output/web_expand/global_antibody_drug_conjugates_market_global/"
    "global_antibody_drug_conjugates_market_FINAL.xlsx"
)
SUMMARY = XLSX.parent / "adc_contact_role_linkedin_summary.json"

BAD_PHRASES = (
    "curriculum", "privacy", "cookie", "subscribe", "newsletter", "click here",
    "read more", "learn more", "not in your", "javascript", "undefined",
)


def main() -> None:
    rows = load_rows(XLSX)
    cleaned = 0
    for r in rows:
        # restore known overrides
        if r["Company"] in KNOWN:
            p, role, li = KNOWN[r["Company"]]
            r["Contact Person"] = p
            r["Role"] = role
            if li:
                r["LinkedIn"] = li
            elif "/company/" in (r.get("LinkedIn") or "") or empty(r.get("LinkedIn")):
                if "/in/" not in (r.get("LinkedIn") or ""):
                    r["LinkedIn"] = "Not publicly disclosed"
            continue

        cp = r.get("Contact Person") or ""
        role = r.get("Role") or ""
        bad = False
        if any(b in cp.lower() for b in BAD_PHRASES):
            bad = True
        if cp and not empty(cp) and not looks_person(cp, r["Company"]):
            # allow known-style placeholders only if from KNOWN (handled above)
            bad = True
        if bad:
            r["Contact Person"] = "Not publicly disclosed"
            r["Role"] = "Not publicly disclosed"
            if "/in/" not in (r.get("LinkedIn") or ""):
                r["LinkedIn"] = "Not publicly disclosed"
            cleaned += 1

        # LinkedIn must be person URL
        li = r.get("LinkedIn") or ""
        if li and "/company/" in li:
            r["LinkedIn"] = "Not publicly disclosed"
        if li and "/in/" not in li and not empty(li):
            # keep only person profiles
            r["LinkedIn"] = "Not publicly disclosed"

    rows.sort(key=lambda x: x["Company"].lower())
    write_xlsx(XLSX, rows)

    named = sum(1 for r in rows if looks_person(r.get("Contact Person") or "", r["Company"]))
    # also count KNOWN placeholders that are intentional non-person leadership labels
    known_fill = sum(1 for r in rows if r["Company"] in KNOWN and not empty(r.get("Contact Person")))
    role_n = sum(1 for r in rows if not empty(r.get("Role")))
    li_n = sum(1 for r in rows if "/in/" in (r.get("LinkedIn") or ""))
    summary = {
        "cleaned_garbage": cleaned,
        "named_person_contacts": named,
        "known_or_named_contacts": max(named, known_fill),
        "roles_filled": role_n,
        "person_linkedin": li_n,
        "total": len(rows),
        "samples": [
            {
                "company": r["Company"],
                "contact": r.get("Contact Person"),
                "role": r.get("Role"),
                "linkedin": r.get("LinkedIn"),
            }
            for r in rows
            if looks_person(r.get("Contact Person") or "", r["Company"])
        ][:15],
    }
    SUMMARY.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"cleaned={cleaned} named={named} known_fill={known_fill} roles={role_n} person_li={li_n}")
    for s in summary["samples"][:10]:
        print(f"- {s['company']}: {s['contact']} / {s['role']} / {s['linkedin']}")


if __name__ == "__main__":
    main()
