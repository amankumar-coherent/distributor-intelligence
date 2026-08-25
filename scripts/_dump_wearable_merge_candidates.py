# -*- coding: utf-8 -*-
"""Dump wearable company+website for merge candidates (PowerShell-safe)."""
import json
from pathlib import Path

d = json.loads(
    Path(
        r"D:\vendor-intel\output\web_expand\global_wearable_medical_devices_market_global"
        r"\_wearable_website_unique_audit.json"
    ).read_text(encoding="utf-8")
)
by = {r["Company"]: r for r in d["rows"]}
cands = []
for fam in d["brand_fams"].values():
    cands.extend(fam)
cands.extend(d["legal_suffix_names"])
for k, v in d["multi_legal_core"].items():
    if k.startswith("medi") or k.startswith("health") or k in {
        "amazon",
        "walmart",
        "b. braun",
        "apex medical",
        "vwr",
        "transmedic",
        "unipharm",
        "pacific healthcare",
        "medlife",
    }:
        cands.extend(v)

seen = set()
lines = []
for c in cands:
    if c in seen:
        continue
    seen.add(c)
    r = by[c]
    p = d["all_probes"].get(c, {})
    lines.append(
        f"{p.get('ok')}|{str(p.get('reason', ''))[:24]}|{c}|{r['Website']}"
    )
Path(
    r"D:\vendor-intel\output\web_expand\global_wearable_medical_devices_market_global"
    r"\_wearable_merge_candidates.txt"
).write_text("\n".join(lines), encoding="utf-8")
print("wrote", len(lines), "lines")
print("\n".join(lines[:80]))
