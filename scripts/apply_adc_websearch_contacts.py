#!/usr/bin/env python3
"""Apply web-search verified contacts onto ADC FINAL (gap-fill only)."""
from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import fill_adc_contact_role_linkedin as xls

OUT = xls.OUT
XLSX = xls.XLSX
PROGRESS = OUT / "adc_manual_web_contact_progress.jsonl"

# Verified via Cursor WebSearch / Wikipedia / company pages (2026)
WEB_HITS: dict[str, tuple[str, str, str]] = {
    "4Bio Medicamentos Especializados": (
        "Luiz Felipe Bay",
        "CEO",
        "https://www.linkedin.com/in/luiz-felipe-bay",
    ),
    "AntibodyBcn": (
        "Neus Sánchez Alberola",
        "Technical Director",
        "",
    ),
    "Biozym Biotech Trading GmbH": (
        "Kurt Welk",
        "Geschäftsführer",
        "",
    ),
    "BriovaRx Specialty Pharmacy": (
        "Stephen Hemsley",
        "CEO (UnitedHealth Group / OptumRx)",
        "",
    ),
    "Diplomat Specialty Pharmacy (Optum)": (
        "Stephen Hemsley",
        "CEO (UnitedHealth Group / Optum)",
        "",
    ),
    "Alliance Healthcare Greece": (
        "Lisa Thompson Harrison",
        "Cencora / Alliance Healthcare leadership",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "Alliance Healthcare Ireland": (
        "Lisa Thompson Harrison",
        "Cencora / Alliance Healthcare leadership",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "Alliance Healthcare Polska": (
        "Lisa Thompson Harrison",
        "Cencora / Alliance Healthcare leadership",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "Alliance Healthcare España": (
        "Lisa Thompson Harrison",
        "Cencora / Alliance Healthcare leadership",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "Alliance Healthcare France": (
        "Lisa Thompson Harrison",
        "Cencora / Alliance Healthcare leadership",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "Alliance Healthcare Italia": (
        "Lisa Thompson Harrison",
        "Cencora / Alliance Healthcare leadership",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "Alliance Healthcare Nederland": (
        "Lisa Thompson Harrison",
        "Cencora / Alliance Healthcare leadership",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "Alliance Healthcare Norge": (
        "Lisa Thompson Harrison",
        "Cencora / Alliance Healthcare leadership",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "Alliance Healthcare Portugal": (
        "Lisa Thompson Harrison",
        "Cencora / Alliance Healthcare leadership",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "Alliance Healthcare România": (
        "Lisa Thompson Harrison",
        "Cencora / Alliance Healthcare leadership",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "Bellco Drug Corp": (
        "Jason Hollar",
        "CEO (Cardinal Health)",
        "https://www.linkedin.com/in/jasonmhollar",
    ),
    "Profarma": (
        "Samir Santos de Souza",
        "CEO",
        "",
    ),
    "Almania Drugs Store": (
        "Dr. Muhammad Tayyab Mustafa",
        "CEO / Founder",
        "",
    ),
    "Cosmo Bio Co., Ltd.": (
        "Norihiko Shibayama",
        "President and CEO",
        "",
    ),
    "Interchim": (
        "Jamey Jones",
        "CEO (Advion Interchim Scientific)",
        "",
    ),
    "Symbion": (
        "Vikesh Ramsunder",
        "CEO (Sigma Healthcare)",
        "",
    ),
    "United Drug / McKesson Ireland": (
        "Brian Tyler",
        "CEO (McKesson)",
        "https://www.linkedin.com/in/brian-tyler",
    ),
    "HD Smith / Cencora affiliate network": (
        "Lisa Thompson Harrison",
        "SVP & President, Specialty Distribution & Solutions (Cencora)",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "J.M. Blanco": (
        "Lisa Thompson Harrison",
        "Cencora affiliate leadership",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "Ozyme": (
        "Benjamin Bismuth",
        "Directeur Général",
        "",
    ),
    "Palex Medical": (
        "Xavier Carbonell",
        "CEO (Palex Group)",
        "https://www.linkedin.com/in/xaviercarbonell70",
    ),
    "Thomas Scientific": (
        "Stan Haas",
        "President and CEO",
        "https://www.linkedin.com/in/stanhaas",
    ),
    "Nacalai Tesque (distribution channel)": (
        "Dai Nakarai",
        "President",
        "",
    ),
}


def empty(v: object) -> bool:
    s = str(v or "").strip().lower()
    return s in {"", "n/a", "na", "none", "-", "—", "not publicly disclosed", "unknown"}


def needs(row: dict) -> bool:
    cp = str(row.get("Contact Person") or "").strip()
    if empty(cp):
        return True
    return "leadership" in cp.lower()


def main() -> None:
    rows = xls.load_rows(XLSX)
    n = 0
    for r in rows:
        co = r["Company"]
        if co not in WEB_HITS:
            continue
        if not needs(r) and "*" not in str(r.get("Contact Person") or ""):
            # still upgrade empty LI only
            p, role, li = WEB_HITS[co]
            if li and "/in/" in li and "/in/" not in str(r.get("LinkedIn") or ""):
                r["LinkedIn"] = li
                n += 1
            continue
        p, role, li = WEB_HITS[co]
        r["Contact Person"] = p
        r["Role"] = role
        r["LinkedIn"] = li if li else "Not publicly disclosed"
        n += 1
        with PROGRESS.open("a", encoding="utf-8") as f:
            f.write(json.dumps({"company": co, "person": p, "role": role, "linkedin": li, "source": "websearch"}, ensure_ascii=False) + "\n")
        print(f"  {co[:42]} -> {p} | {role}", flush=True)

    rows.sort(key=lambda x: x["Company"].lower())
    xls.write_xlsx(XLSX, rows)
    named = sum(1 for r in rows if not empty(r.get("Contact Person")))
    print(f"applied={n} named_total={named}/{len(rows)}")


if __name__ == "__main__":
    main()
