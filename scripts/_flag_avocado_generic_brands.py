#!/usr/bin/env python3
"""Flag generic/placeholder Key Brands and mismatched Core/Focus."""
from __future__ import annotations

import json
import re
from pathlib import Path

rows = json.loads(
    Path(
        r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global"
        r"\_avocado_cats_audit.json"
    ).read_text(encoding="utf-8")
)

PLACEHOLDER = re.compile(
    r"(?i)("
    r"principal .+ brands|"
    r"private label & brand|"
    r"private[- ]label .+ oils|"
    r"member banner private label|"
    r"contract-approved|"
    r"traded principal|"
    r"group ingredient|"
    r"house/?private-label oils|"
    r"house natural oil|"
    r"specialty avocado .+ brands|"
    r"warehouse private label|"
    r"cash & carry private label|"
    r"retailer private label|"
    r"national oil|"
    r"brand oils?, others|"
    r"brand edible oils|"
    r"brand avocado|"
    r"regional & private-label|"
    r"imported specialty oil brands|"
    r"produce house brands|"
    r"grower avocado|"
    r"cosmetic ingredient / oil|"
    r"portfolio principals|"
    r"supplier brands \(oils|"
    r"azelis principal brands|"
    r"barentz portfolio|"
    r"bidfood private label|"
    r"brakes private label|"
    r"banner private label"
    r")"
)

generic_brands = []
realish_brands = []
for r in rows:
    b = r["brands"]
    # Real-ish if contains a capitalized proper noun brand-like token that isn't only Private/Principal
    if PLACEHOLDER.search(b) or b.lower() in {
        "private label & regional avocado oil brands, others",
        "private label & brand edible oils, others",
    }:
        generic_brands.append(r)
    else:
        realish_brands.append(r)

print(f"generic/placeholder brands: {len(generic_brands)}/{len(rows)}")
print(f"realish brands: {len(realish_brands)}/{len(rows)}")
print("\n--- realish sample ---")
for r in realish_brands[:40]:
    print(f"{r['company'][:34]:34}|{r['brands'][:70]}")

print("\n--- generic sample ---")
for r in generic_brands[:25]:
    print(f"{r['company'][:34]:34}|{r['brands'][:70]}")

# focus/core mismatch heuristics
print("\n--- cosmetic focus but foodservice core (sample) ---")
n = 0
for r in rows:
    fl = r["focus"].lower()
    cl = r["core"].lower()
    if "cosmetic" in fl and ("foodservice" in cl or "grocery" in cl or "wholesale" in cl):
        print(r["company"], "|", r["core"][:50], "|", r["focus"][:50])
        n += 1
print("count", n)

print("\n--- foodservice focus but cosmetic core ---")
n = 0
for r in rows:
    fl = r["focus"].lower()
    cl = r["core"].lower()
    if "cosmetic" in cl and ("foodservice" in fl or "culinary" in fl or "grocery" in fl):
        print(r["company"], "|", r["core"][:50], "|", r["focus"][:50])
        n += 1
print("count", n)

Path(
    r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global\_avocado_generic_brands.json"
).write_text(json.dumps([r["company"] for r in generic_brands], indent=2), encoding="utf-8")
