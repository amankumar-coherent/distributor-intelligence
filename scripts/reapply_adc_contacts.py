#!/usr/bin/env python3
"""Re-apply known contacts + clean garbage; write summary."""
from __future__ import annotations

import importlib
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import fill_adc_contact_role_linkedin as m

importlib.reload(m)

XLSX = Path(__file__).resolve().parents[1] / (
    "output/web_expand/global_antibody_drug_conjugates_market_global/"
    "global_antibody_drug_conjugates_market_FINAL.xlsx"
)
SUMMARY = XLSX.parent / "adc_contact_role_linkedin_summary.json"

rows = m.load_rows(XLSX)
for r in rows:
    if r["Company"] in m.KNOWN:
        p, role, li = m.KNOWN[r["Company"]]
        r["Contact Person"] = p
        r["Role"] = role
        if li:
            r["LinkedIn"] = li
        elif "/in/" not in (r.get("LinkedIn") or ""):
            r["LinkedIn"] = "Not publicly disclosed"
        continue

    cp = r.get("Contact Person") or ""
    if cp and not m.empty(cp) and not m.looks_person(cp, r["Company"]):
        r["Contact Person"] = "Not publicly disclosed"
        r["Role"] = "Not publicly disclosed"
        if "/in/" not in (r.get("LinkedIn") or ""):
            r["LinkedIn"] = "Not publicly disclosed"

    li = r.get("LinkedIn") or ""
    if "/company/" in li or (li and "/in/" not in li and not m.empty(li)):
        r["LinkedIn"] = "Not publicly disclosed"

# Allow Dr. names: soften by accepting if starts with Dr.
for r in rows:
    cp = r.get("Contact Person") or ""
    if cp.startswith("Dr.") and r["Company"] in m.KNOWN:
        pass

rows.sort(key=lambda x: x["Company"].lower())
m.write_xlsx(XLSX, rows)

named = sum(1 for r in rows if m.looks_person(r.get("Contact Person") or "", r["Company"]) or (
    r["Company"] in m.KNOWN and not m.empty(r.get("Contact Person")) and "not publicly" not in (r.get("Contact Person") or "").lower()
))
# recount named properly
named = 0
for r in rows:
    cp = r.get("Contact Person") or ""
    if m.empty(cp) or "not publicly" in cp.lower():
        continue
    if m.looks_person(cp, r["Company"]) or (cp.startswith("Dr.") and len(cp.split()) >= 2):
        named += 1
    elif r["Company"] in m.KNOWN and all(x not in cp.lower() for x in ("leadership", "joint md", "not publicly")):
        # known named even if unusual formatting
        if " " in cp:
            named += 1

role_n = sum(1 for r in rows if not m.empty(r.get("Role")) and "not publicly" not in (r.get("Role") or "").lower())
li_n = sum(1 for r in rows if "/in/" in (r.get("LinkedIn") or ""))

samples = []
for r in rows:
    cp = r.get("Contact Person") or ""
    if m.empty(cp) or "not publicly" in cp.lower():
        continue
    samples.append({
        "company": r["Company"],
        "contact": cp,
        "role": r.get("Role"),
        "linkedin": r.get("LinkedIn"),
    })
    if len(samples) >= 20:
        break

summary = {
    "total": len(rows),
    "contact_person_filled": named,
    "role_filled": role_n,
    "person_linkedin_filled": li_n,
    "samples": samples,
}
SUMMARY.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
print(json.dumps(summary, indent=2)[:2500])
print(f"\ncontact={named} role={role_n} person_li={li_n} / {len(rows)}")
