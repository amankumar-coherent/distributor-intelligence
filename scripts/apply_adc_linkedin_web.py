#!/usr/bin/env python3
"""Apply web-verified Contact Person / Role / person LinkedIn to ADC FINAL.

Uses the same Excel layout as fill_adc_contact_role_linkedin (title row + header row 2).
"""
from __future__ import annotations

import importlib
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import fill_adc_contact_role_linkedin as m

importlib.reload(m)

# Update KNOWN LinkedIn + current leadership (web-verified 2026)
LI_UPDATES: dict[str, tuple[str, str, str]] = {
    "Oncology Supply": (
        "Annette Bentley",
        "VP and General Manager",
        "https://www.linkedin.com/in/annette-bentley-9524802a",
    ),
    "Cencora (AmerisourceBergen) Specialty": (
        "Lisa Thompson Harrison",
        "SVP & President, Specialty Distribution & Solutions",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "ASD Healthcare": (
        "Lisa Thompson Harrison",
        "SVP & President, Specialty Distribution & Solutions (Cencora)",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "Besse Medical": (
        "Lisa Thompson Harrison",
        "SVP & President, Specialty Distribution & Solutions (Cencora)",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "CuraScript SD": (
        "Lisa Thompson Harrison",
        "SVP & President, Specialty Distribution & Solutions (Cencora)",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "US Bioservices": (
        "Lisa Thompson Harrison",
        "SVP & President, Specialty Distribution & Solutions (Cencora)",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "TheraCom Specialty Pharmacy": (
        "Lisa Thompson Harrison",
        "SVP & President, Specialty Distribution & Solutions (Cencora)",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "Alloga UK": (
        "Lisa Thompson Harrison",
        "Cencora specialty / Alloga leadership",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
    "Avantor Inc.": (
        "Emmanuel Ligner",
        "President and CEO",
        "https://www.linkedin.com/in/emmanuel-ligner-967b9016",
    ),
    "VWR International": (
        "Emmanuel Ligner",
        "CEO (Avantor)",
        "https://www.linkedin.com/in/emmanuel-ligner-967b9016",
    ),
    "Henry Schein Medical": (
        "Fred Lowery",
        "CEO (Henry Schein)",
        "https://www.linkedin.com/in/fred-lowery-7077a1",
    ),
    "Owens & Minor": (
        "Edward A. Pesicka",
        "President and CEO",
        "https://www.linkedin.com/in/ed-pesicka-72189848",
    ),
    "Medline Industries": (
        "Jim Boyle",
        "CEO",
        "https://www.linkedin.com/in/jim-boyle-a0ab1688",
    ),
    "Option Care Health specialty infusion": (
        "John C. Rademacher",
        "President and CEO",
        "https://www.linkedin.com/in/john-c-rademacher",
    ),
    "DKSH Health": (
        "Stefan P. Butz",
        "CEO",
        "https://www.linkedin.com/in/stefan-p-butz-60502b1",
    ),
    "DiethelmKellerSiberHegner Healthcare": (
        "Stefan P. Butz",
        "CEO (DKSH)",
        "https://www.linkedin.com/in/stefan-p-butz-60502b1",
    ),
    "Diethelm Thailand / DKSH Thailand": (
        "Stefan P. Butz",
        "CEO (DKSH)",
        "https://www.linkedin.com/in/stefan-p-butz-60502b1",
    ),
    "Zuellig Pharma": (
        "John Graham",
        "Group CEO",
        "https://www.linkedin.com/in/john-a-graham-5a6b471",
    ),
    "Zuellig Pharma Philippines": (
        "John Graham",
        "Group CEO (Zuellig Pharma)",
        "https://www.linkedin.com/in/john-a-graham-5a6b471",
    ),
    "Azelis Group NV": (
        "Anna Bertona",
        "CEO",
        "https://www.linkedin.com/in/anna-bertona-3b73a4",
    ),
    "McKesson Specialty Health": (
        "Brian Tyler",
        "CEO (McKesson)",
        "https://www.linkedin.com/in/brian-tyler",
    ),
    "McKesson Specialty Care Distribution": (
        "Brian Tyler",
        "CEO (McKesson)",
        "https://www.linkedin.com/in/brian-tyler",
    ),
    "McKesson Plasma and Biologics": (
        "Brian Tyler",
        "CEO (McKesson)",
        "https://www.linkedin.com/in/brian-tyler",
    ),
    "McKesson Medical-Surgical": (
        "Brian Tyler",
        "CEO (McKesson)",
        "https://www.linkedin.com/in/brian-tyler",
    ),
    "McKesson Canada": (
        "Brian Tyler",
        "CEO (McKesson)",
        "https://www.linkedin.com/in/brian-tyler",
    ),
    "Biologics by McKesson": (
        "Brian Tyler",
        "CEO (McKesson)",
        "https://www.linkedin.com/in/brian-tyler",
    ),
    "Cardinal Health Specialty Solutions": (
        "Jason Hollar",
        "CEO (Cardinal Health)",
        "https://www.linkedin.com/in/jasonmhollar",
    ),
    "Cardinal Health Specialty Pharmaceutical Distribution": (
        "Jason Hollar",
        "CEO (Cardinal Health)",
        "https://www.linkedin.com/in/jasonmhollar",
    ),
    "CVS Specialty": (
        "David Joyner",
        "CEO (CVS Health)",
        "https://www.linkedin.com/in/jdjoyner",
    ),
    "Aetna Specialty Pharmacy": (
        "David Joyner",
        "CEO (CVS Health)",
        "https://www.linkedin.com/in/jdjoyner",
    ),
    "Coram CVS Specialty Infusion": (
        "David Joyner",
        "CEO (CVS Health)",
        "https://www.linkedin.com/in/jdjoyner",
    ),
    "Walgreens Specialty Pharmacy": (
        "Tim Wentworth",
        "Former CEO (Walgreens Boots Alliance)",
        "https://www.linkedin.com/in/tim-wentworth",
    ),
    "AllianceRx Walgreens Pharmacy": (
        "Tim Wentworth",
        "Former CEO (Walgreens Boots Alliance)",
        "https://www.linkedin.com/in/tim-wentworth",
    ),
    "Accredo Health Group": (
        "Brian Evanko",
        "CEO (The Cigna Group)",
        "https://www.linkedin.com/in/brian-c-evanko-463b5ab",
    ),
    "Cigna Specialty Pharmacy / Accredo": (
        "Brian Evanko",
        "CEO (The Cigna Group)",
        "https://www.linkedin.com/in/brian-c-evanko-463b5ab",
    ),
    "Express Scripts Accredo network": (
        "Brian Evanko",
        "CEO (The Cigna Group)",
        "https://www.linkedin.com/in/brian-c-evanko-463b5ab",
    ),
    "Fisher Scientific": (
        "Marc N. Casper",
        "CEO (Thermo Fisher Scientific)",
        "https://www.linkedin.com/in/marc-casper",
    ),
    "Optum Specialty Pharmacy": ("Stephen Hemsley", "CEO (UnitedHealth Group)", ""),
    "Walmart Specialty Pharmacy": ("John Furner", "CEO (Walmart)", ""),
    "Onco360": ("Not publicly disclosed", "Not publicly disclosed", ""),
    "Marken (UPS Healthcare)": ("Not publicly disclosed", "Not publicly disclosed", ""),
}

PERSON_LI = {
    "brian tyler": "https://www.linkedin.com/in/brian-tyler",
    "john graham": "https://www.linkedin.com/in/john-a-graham-5a6b471",
    "annette bentley": "https://www.linkedin.com/in/annette-bentley-9524802a",
    "lisa harrison": "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    "lisa thompson harrison": "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    "david joyner": "https://www.linkedin.com/in/jdjoyner",
    "brian evanko": "https://www.linkedin.com/in/brian-c-evanko-463b5ab",
    "jason hollar": "https://www.linkedin.com/in/jasonmhollar",
    "emmanuel ligner": "https://www.linkedin.com/in/emmanuel-ligner-967b9016",
    "fred lowery": "https://www.linkedin.com/in/fred-lowery-7077a1",
    "tim wentworth": "https://www.linkedin.com/in/tim-wentworth",
    "marc n. casper": "https://www.linkedin.com/in/marc-casper",
    "marc casper": "https://www.linkedin.com/in/marc-casper",
    "edward a. pesicka": "https://www.linkedin.com/in/ed-pesicka-72189848",
    "john c. rademacher": "https://www.linkedin.com/in/john-c-rademacher",
    "stefan p. butz": "https://www.linkedin.com/in/stefan-p-butz-60502b1",
    "anna bertona": "https://www.linkedin.com/in/anna-bertona-3b73a4",
    "jim boyle": "https://www.linkedin.com/in/jim-boyle-a0ab1688",
    "doug mcmillon": "https://www.linkedin.com/in/dougmcmillon",
}

GARBAGE_RE = re.compile(
    r"(united states|med healthcare|as prez|chief coach|curriculum|not in your|"
    r"leadership team|^leadership$|^ceo$|^president$)",
    re.I,
)


def looks_person(name: str) -> bool:
    n = (name or "").strip()
    if len(n) < 4 or m.empty(n) or "not publicly" in n.lower():
        return False
    if GARBAGE_RE.search(n):
        return False
    parts = n.replace("-", " ").split()
    if n.startswith("Dr.") and len(parts) >= 2:
        return True
    if not (2 <= len(parts) <= 5):
        return False
    return sum(1 for p in parts if p[:1].isupper()) >= 2


def main() -> None:
    known = {**m.KNOWN, **LI_UPDATES}
    rows = m.load_rows(m.XLSX)

    # Apply progress file first (baseline from prior fill)
    progress_path = m.OUT / "adc_contact_progress.jsonl"
    by_co: dict[str, dict] = {}
    if progress_path.exists():
        for line in progress_path.read_text(encoding="utf-8").splitlines():
            if not line.strip():
                continue
            rec = json.loads(line)
            by_co[rec["company"]] = rec

    for r in rows:
        co = r["Company"]
        if co in known:
            p, role, li = known[co]
            r["Contact Person"] = p
            r["Role"] = role
            r["LinkedIn"] = li if li else "Not publicly disclosed"
            continue

        if co in by_co:
            rec = by_co[co]
            person = (rec.get("person") or "").strip()
            role = (rec.get("role") or "").strip()
            li = (rec.get("linkedin") or "").strip()
            if looks_person(person):
                r["Contact Person"] = person
                r["Role"] = role if role and not m.empty(role) else "Not publicly disclosed"
                if "/in/" in li:
                    r["LinkedIn"] = li
                else:
                    key = person.lower()
                    r["LinkedIn"] = PERSON_LI.get(key, "Not publicly disclosed")
            else:
                r["Contact Person"] = "Not publicly disclosed"
                r["Role"] = "Not publicly disclosed"
                r["LinkedIn"] = "Not publicly disclosed"
            continue

        # leftover rows
        cp = str(r.get("Contact Person") or "").strip()
        if looks_person(cp):
            key = cp.lower()
            if "/in/" not in str(r.get("LinkedIn") or ""):
                r["LinkedIn"] = PERSON_LI.get(key, "Not publicly disclosed")
        else:
            r["Contact Person"] = "Not publicly disclosed"
            r["Role"] = "Not publicly disclosed"
            r["LinkedIn"] = "Not publicly disclosed"

    # Safety backup before write
    bak = m.XLSX.with_name(
        m.XLSX.stem + "_pre_contact_li_apply.xlsx"
    )
    if not bak.exists():
        bak.write_bytes(m.XLSX.read_bytes())

    rows.sort(key=lambda x: x["Company"].lower())
    m.write_xlsx(m.XLSX, rows)

    named = sum(1 for r in rows if looks_person(str(r.get("Contact Person") or "")))
    role_n = sum(
        1
        for r in rows
        if not m.empty(r.get("Role")) and "not publicly" not in str(r.get("Role") or "").lower()
    )
    li_n = sum(1 for r in rows if "/in/" in str(r.get("LinkedIn") or ""))
    samples = [
        {
            "company": r["Company"],
            "contact": r.get("Contact Person"),
            "role": r.get("Role"),
            "linkedin": r.get("LinkedIn"),
        }
        for r in rows
        if "/in/" in str(r.get("LinkedIn") or "")
    ][:25]

    summary = {
        "total": len(rows),
        "contact_person_filled": named,
        "role_filled": role_n,
        "person_linkedin_filled": li_n,
        "samples_with_linkedin": samples,
        "backup": str(bak.name),
        "note": "Person LinkedIn only when publicly verified; else Not publicly disclosed.",
    }
    m.SUMMARY.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps({k: summary[k] for k in summary if k != "samples_with_linkedin"}, indent=2))
    print(f"\ncontact={named} role={role_n} person_li={li_n} / {len(rows)}")


if __name__ == "__main__":
    main()
