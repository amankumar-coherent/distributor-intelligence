#!/usr/bin/env python3
"""Audit avocado OFFICE_SLIM geo columns for mismatches."""
from __future__ import annotations

import json
import re
from pathlib import Path

rows = json.loads(
    Path(
        r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global"
        r"\_avocado_geo_audit_snapshot.json"
    ).read_text(encoding="utf-8")
)

COUNTRY_TO_CONT = {
    "usa": "North America",
    "united states": "North America",
    "u.s.": "North America",
    "canada": "North America",
    "mexico": "North America",
    "uk": "Europe",
    "united kingdom": "Europe",
    "england": "Europe",
    "germany": "Europe",
    "france": "Europe",
    "netherlands": "Europe",
    "spain": "Europe",
    "italy": "Europe",
    "belgium": "Europe",
    "switzerland": "Europe",
    "sweden": "Europe",
    "norway": "Europe",
    "denmark": "Europe",
    "ireland": "Europe",
    "poland": "Europe",
    "portugal": "Europe",
    "austria": "Europe",
    "china": "Asia",
    "japan": "Asia",
    "india": "Asia",
    "singapore": "Asia",
    "malaysia": "Asia",
    "thailand": "Asia",
    "indonesia": "Asia",
    "vietnam": "Asia",
    "korea": "Asia",
    "hong kong": "Asia",
    "uae": "Asia",
    "dubai": "Asia",
    "saudi": "Asia",
    "israel": "Asia",
    "turkey": "Asia / Europe",
    "australia": "Oceania",
    "new zealand": "Oceania",
    "brazil": "South America",
    "chile": "South America",
    "argentina": "South America",
    "peru": "South America",
    "colombia": "South America",
    "south africa": "Africa",
    "kenya": "Africa",
    "egypt": "Africa",
    "morocco": "Africa",
}

mism = []
for r in rows:
    hq = (r["hq"] or "").lower()
    cont = (r["continent"] or "").lower()
    expected = None
    for k, v in COUNTRY_TO_CONT.items():
        if k in hq:
            expected = v.lower()
            break
    if expected and expected.split()[0] not in cont and "global" not in cont and "multi" not in cont:
        # turkey special
        if "turkey" in hq and ("asia" in cont or "europe" in cont):
            continue
        mism.append((r["company"], r["hq"], r["continent"], r["presence"][:90]))

print(f"HQ vs Continent mismatches: {len(mism)}")
for m in mism[:50]:
    print(" | ".join(m))

print("\n--- all rows (compact) ---")
for i, r in enumerate(rows, 1):
    print(
        f"{i:03d}|{r['company']}|{r['founded']}|{r['hq']}|{r['continent']}|{r['presence']}"
    )
