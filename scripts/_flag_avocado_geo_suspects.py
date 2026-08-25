#!/usr/bin/env python3
"""Flag avocado geo rows needing web verification."""
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

def countries_in(text: str) -> set[str]:
    t = (text or "").lower()
    # crude tokens after split
    parts = re.split(r"[;,/|()]+", t)
    return {p.strip() for p in parts if len(p.strip()) > 2}


suspect = []
for r in rows:
    reasons = []
    hq = r["hq"]
    presence = r["presence"]
    cont = r["continent"]
    co = r["company"]

    # Anchor-style: HQ country not in presence and presence is different single country
    hq_l = hq.lower()
    pr_l = presence.lower()
    if "australia" in hq_l and "australia" not in pr_l:
        reasons.append("HQ Australia missing from presence")
    if "united kingdom" in hq_l or ", uk" in hq_l or hq_l.endswith(" uk"):
        if "united kingdom" not in pr_l and "uk" not in pr_l and "england" not in pr_l:
            reasons.append("HQ UK missing from presence")
    if "germany" in hq_l and "germany" not in pr_l:
        reasons.append("HQ Germany missing from presence")
    if ("united states" in hq_l or ", usa" in hq_l or hq_l.endswith(" usa") or "u.s." in hq_l) and (
        "united states" not in pr_l and "usa" not in pr_l and "u.s." not in pr_l
    ):
        reasons.append("HQ USA missing from presence")
    if ("netherlands" in hq_l or "holland" in hq_l) and "netherlands" not in pr_l:
        reasons.append("HQ NL missing from presence")

    # presence only one token but continent lists many countries inside parens
    cont_countries = re.findall(r"\(([^)]+)\)", cont)
    flat = ";".join(cont_countries)
    if flat.count(",") >= 2 and presence.count(";") == 0 and presence.count(",") == 0:
        reasons.append("continent lists multi countries but presence single")

    # known bad website patterns already in snapshot? skip

    if reasons:
        suspect.append({**r, "reasons": reasons})

print(f"suspect={len(suspect)}")
for s in suspect:
    print(f"- {s['company']}: {'; '.join(s['reasons'])}")
    print(f"  HQ={s['hq']} | C={s['continent']} | P={s['presence']}")

Path(
    r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global\_avocado_geo_suspects.json"
).write_text(json.dumps(suspect, indent=2, ensure_ascii=False), encoding="utf-8")
