#!/usr/bin/env python3
"""Audit Core Categories / Specialty Focus / Key Brands on avocado OFFICE_SLIM."""
from __future__ import annotations

import json
import re
from pathlib import Path

import openpyxl

p = Path(
    r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global"
    r"\global_avocado_oil_market_OFFICE_SLIM.xlsx"
)
wb = openpyxl.load_workbook(p, data_only=True)
ws = wb.active
rows = list(ws.iter_rows(values_only=True))
hdr = [str(c or "") for c in rows[1]]
ic = {h: i for i, h in enumerate(hdr) if h}

NPD = {
    "",
    "n/a",
    "na",
    "none",
    "-",
    "—",
    "not publicly disclosed",
    "unknown",
    "tbd",
    "null",
}
COUNTRY = re.compile(
    r"\b(United States|USA|UK|United Kingdom|Germany|France|Canada|Mexico|"
    r"Australia|China|India|Japan|Spain|Italy|Netherlands|Belgium|Switzerland|"
    r"Brazil|Chile|Peru|UAE|Dubai|Singapore|Malaysia|Thailand|Indonesia|"
    r"Vietnam|South Africa|Ireland|Poland|Sweden|Norway|Denmark|Scotland|"
    r"England|Wales|Austria|Portugal|Greece|Turkey|Saudi Arabia|Qatar|"
    r"Kuwait|Oman|Bahrain|New Zealand|Korea|Taiwan|Hong Kong|Philippines|"
    r"Colombia|Argentina|Egypt|Morocco|Kenya|Nigeria)\b",
    re.I,
)

out = []
for r in rows[2:]:
    out.append(
        {
            "company": str(r[ic["Company"]] or "").strip(),
            "website": str(r[ic.get("Website", "")] or "").strip() if "Website" in ic else "",
            "core": str(r[ic["Core Categories"]] or "").strip(),
            "focus": str(r[ic["Specialty Focus"]] or "").strip(),
            "brands": str(r[ic["Key Brands Represented"]] or "").strip(),
        }
    )

for label, key in [
    ("Core Categories", "core"),
    ("Specialty Focus", "focus"),
    ("Key Brands Represented", "brands"),
]:
    empty = sum(1 for x in out if x[key].lower() in NPD)
    print(f"{label}: empty/NPD={empty}/{len(out)}")

bad_focus_country = []
bad_focus_ending = []
brands_npd = []
suspect_brands = []
generic_templates = []

for x in out:
    if x["focus"].lower() not in NPD:
        if COUNTRY.search(x["focus"]):
            bad_focus_country.append((x["company"], x["focus"][:90]))
        if not re.search(r",\s*others\s*$", x["focus"], re.I):
            bad_focus_ending.append((x["company"], x["focus"][-50:]))
    if x["brands"].lower() in NPD:
        brands_npd.append(x["company"])
    # template-y cores
    if x["core"].startswith("Edible oils; avocado oil; foodservice/retail oils"):
        generic_templates.append(x["company"])
    # brands that look invented / too long / junk
    b = x["brands"]
    if re.search(r"w3schools|example\.com|lorem|chatgpt|reddit|github\.com|todo", b, re.I):
        suspect_brands.append((x["company"], b[:80], "junk"))
    elif len(b) > 220:
        suspect_brands.append((x["company"], b[:80] + "...", "very long"))

print(f"\nfocus has country tokens: {len(bad_focus_country)}")
for t in bad_focus_country[:25]:
    print(" ", t[0], "|", t[1])

print(f"\nfocus missing ', others' ending: {len(bad_focus_ending)}")
for t in bad_focus_ending[:20]:
    print(" ", t)

print(f"\nbrands NPD: {len(brands_npd)}")
print("brands NPD sample:", brands_npd[:30])

print(f"\ngeneric edible-oils template core: {len(generic_templates)}")
print(f"suspect brands: {len(suspect_brands)}")
for t in suspect_brands[:20]:
    print(" ", t)

print("\n--- first 50 rows ---")
for x in out[:50]:
    print(f"{x['company'][:32]:32}|{x['core'][:42]:42}|{x['focus'][:36]:36}|{x['brands'][:36]}")

Path(
    r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global\_avocado_cats_audit.json"
).write_text(json.dumps(out, indent=2, ensure_ascii=False), encoding="utf-8")
print(f"\naudit saved n={len(out)}")
wb.close()
