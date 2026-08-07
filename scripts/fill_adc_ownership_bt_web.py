#!/usr/bin/env python3
"""Fill Ownership + Business Type for ADC using known web facts + Wikipedia (no DeepSeek)."""
from __future__ import annotations

import asyncio
import json
import re
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import quote

import httpx
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Reuse known map + helpers from the DeepSeek script
from verify_adc_ownership_businesstype import (
    KNOWN,
    OWN_OK,
    BT_OK,
    HEADERS,
    WIDTHS,
    normalize_ownership,
    normalize_bt,
    name_heuristic,
    wiki_hint,
    load_rows,
    norm,
)

ROOT = Path(__file__).resolve().parents[1]
XLSX = ROOT / (
    "output/web_expand/global_antibody_drug_conjugates_market_global/"
    "global_antibody_drug_conjugates_market_FINAL.xlsx"
)
SUMMARY = ROOT / (
    "output/web_expand/global_antibody_drug_conjugates_market_global/"
    "adc_ownership_businesstype_summary.json"
)

# Extra web-verified ownership from search / public knowledge
EXTRA: dict[str, tuple[str, str]] = {
    "AAH Pharmaceuticals": ("Subsidiary", "Distributor"),
    "Accredo Health Group": ("Subsidiary", "Specialty Distributor"),
    "Besse Medical": ("Subsidiary", "Specialty Distributor"),  # Cardinal / Cencora network
    "Bellco Drug Corp": ("Subsidiary", "Distributor"),  # historically Cardinal
    "J.M. Blanco": ("Subsidiary", "Distributor"),  # Cencora affiliate
    "Drogueria Betances": ("Private", "Distributor"),
    "Capital Wholesale Drug Co.": ("Private", "Distributor"),
    "Smith Drug Company": ("Private", "Distributor"),
    "Value Drug Company": ("Private", "Distributor"),
    "Mutual Drug Company": ("Private", "Distributor"),
    "Dakota Drug": ("Private", "Distributor"),
    "D&K Healthcare Resources": ("Subsidiary", "Distributor"),
    "HD Smith / Cencora affiliate network": ("Subsidiary", "Distributor"),
    "Farmacias del Ahorro wholesale": ("Private", "Distributor"),
    "Audifarma Colombia": ("Private", "Specialty Distributor"),
    "Difare Ecuador": ("Private", "Distributor"),
    "Quimica Suiza Peru": ("Private", "Distributor"),
    "Química Suiza Peru": ("Private", "Distributor"),
    "Ibn Sina Pharma Egypt": ("Public", "Distributor"),
    "Neopharm": ("Private", "Specialty Distributor"),
    "Neopharm Israel specialty": ("Private", "Specialty Distributor"),
    "Pharmaniaga Logistics": ("Public", "Distributor"),
    "SPIMACO distribution KSA": ("Public", "Distributor"),
    "NUPCO supply partners KSA": ("Public", "Distributor"),
    "Mega Lifesciences distribution": ("Public", "Distributor"),
    "Pacific Science Co., Ltd.": ("Private", "Distributor"),
    "Labclinics": ("Private", "Distributor"),
    "Euroclone S.p.A.": ("Private", "Distributor"),
    "Interchim": ("Private", "Distributor"),
    "Ozyme": ("Private", "Distributor"),
    "Bio-Connect B.V.": ("Private", "Distributor"),
    "Sanbio B.V.": ("Private", "Distributor"),
    "Palex Medical": ("Private", "Distributor"),
    "Prodotti Gianni S.r.l.": ("Private", "Distributor"),
    "Herba Chemosan": ("Subsidiary", "Distributor"),
    "United Drug / McKesson Ireland": ("Subsidiary", "Distributor"),
    "Ebert+Jacobi (NOWEDA Group)": ("Cooperative", "Distributor"),
    "Otto Geilenkirchen Pharma-Großhandel": ("Private", "Distributor"),
    "Richard KEHR Arzneimittel-Großhandlung": ("Private", "Distributor"),
    "Max Jenne Arzneimittel-Großhandlung": ("Private", "Distributor"),
    "GEHE Pharma Handel / Alliance Healthcare Deutschland": ("Subsidiary", "Distributor"),
    "DocMorris / Zur Rose wholesale channel": ("Public", "Distributor"),
    "China Resources Pharmaceutical": ("Public", "Distributor"),
    "Dakewe Biotech": ("Private", "Distributor"),
    "Amyjet Scientific": ("Private", "Distributor"),
    "Asia Biotech Group": ("Private", "Distributor"),
    "AGBL / Milab Scientific": ("Private", "Distributor"),
    "Accesspharm": ("Private", "Specialty Distributor"),
    "Adium Pharma S.A.": ("Private", "Specialty Distributor"),
    "4Bio Medicamentos Especializados": ("Private", "Specialty Distributor"),
    "Agille Medicamento": ("Private", "Specialty Distributor"),
    "ANBFarma": ("Private", "Specialty Distributor"),
    "Amoveri Pontual": ("Private", "Specialty Distributor"),
    "Hospinova": ("Private", "Specialty Distributor"),
    "Victória Medicamentos": ("Private", "Specialty Distributor"),
    "3S Corporation": ("Private", "Distributor"),
    "A.K. Pharma": ("Private", "Distributor"),
    "GNH India": ("Private", "Distributor"),
    "Oddway International": ("Private", "Distributor"),
    "Sunnyco Pharma Distributor": ("Private", "Distributor"),
    "Synbiota Scientific": ("Private", "Distributor"),
    "All Eights (M) Sdn Bhd": ("Private", "Distributor"),
    "All Eights (S) Pte Ltd": ("Private", "Distributor"),
    "Almania Drugs Store": ("Private", "Distributor"),
    "Adwaa Lana Medical Co. Ltd.": ("Private", "Distributor"),
    "Alfagen Laboratuvar": ("Private", "Distributor"),
    "AntibodyBcn": ("Private", "Distributor"),
    "Apop Biosciences": ("Private", "Distributor"),
    "Assay Matrix Pty Ltd": ("Private", "Distributor"),
    "Australian Biosearch": ("Private", "Distributor"),
    "ATCG": ("Private", "Distributor"),
    "AxxessBio": ("Private", "Distributor"),
    "Asiakare": ("Private", "Specialty Distributor"),
    "Common Health Asia": ("Private", "Distributor"),
    "Genomax Technologies Singapore": ("Private", "Distributor"),
    "Movitello / Apex Pharma distribution APAC": ("Private", "Specialty Distributor"),
    "United Pharma Partners": ("Private", "Specialty Distributor"),
    "Specialised Therapeutics Asia": ("Private", "Specialty Distributor"),
    "PharmaPeer": ("Private", "Specialty Distributor"),
    "Ancanex Pharma Trading FZCO": ("Private", "Specialty Distributor"),
    "City Pharmacy / Aster distribution": ("Subsidiary", "Distributor"),
    "Copharma Distribution Morocco": ("Private", "Distributor"),
    "Farmexpert": ("Private", "Distributor"),
    "NMD Grossisthandel": ("Private", "Distributor"),
    "Limedika Lithuania": ("Private", "Distributor"),
    "BIOCOM Africa": ("Private", "Distributor"),
    "Interlab S.A.": ("Private", "Distributor"),
    "Gene X-Press": ("Private", "Distributor"),
    "Truemed / Israeli specialty channel": ("Private", "Specialty Distributor"),
    "ENCO Scientific Services": ("Private", "Distributor"),
    "Mediray": ("Private", "Distributor"),
    "Labex Corporation": ("Private", "Distributor"),
    "Biomed Speciality Pharma": ("Private", "Specialty Distributor"),
    "Atlantic Biologicals": ("Private", "Distributor"),
    "Drugzone Pharmaceuticals Inc.": ("Private", "Specialty Distributor"),
    "DMS Pharmaceutical Group": ("Private", "Distributor"),
    "Dixon-Shane / R&S Northeast": ("Private", "Distributor"),
    "North Carolina Mutual Wholesale Drug Company": ("Private", "Distributor"),
    "Louisiana Wholesale Drug Company": ("Private", "Distributor"),
    "Smith Drug Company": ("Private", "Distributor"),
    "Value Drug Company": ("Private", "Distributor"),
    "Capital Wholesale Drug Co.": ("Private", "Distributor"),
    "AcariaHealth": ("Subsidiary", "Specialty Distributor"),
    "Red Oak Sourcing specialty": ("Joint Venture", "Specialty Distributor"),
    "ClarusONE Sourcing Services": ("Joint Venture", "Distributor"),
}


def write_xlsx(path: Path, rows: list[dict[str, str]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "ADC Distributors"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HEADERS))
    title = ws.cell(1, 1, "Global Antibody Drug Conjugates Market — Distributors")
    title.font = Font(bold=True, size=13, color="FFFFFF")
    title.fill = PatternFill("solid", fgColor="1F4E79")
    hf = PatternFill("solid", fgColor="2E75B6")
    hfont = Font(bold=True, color="FFFFFF")
    thin = Border(
        left=Side(style="thin", color="BDD7EE"),
        right=Side(style="thin", color="BDD7EE"),
        top=Side(style="thin", color="BDD7EE"),
        bottom=Side(style="thin", color="BDD7EE"),
    )
    for i, h in enumerate(HEADERS, 1):
        c = ws.cell(2, i, h)
        c.fill = hf
        c.font = hfont
        c.alignment = Alignment(wrap_text=True, vertical="center")
        c.border = thin
        ws.column_dimensions[get_column_letter(i)].width = WIDTHS.get(h, 16)
    for ri, row in enumerate(rows, 3):
        for ci, h in enumerate(HEADERS, 1):
            cell = ws.cell(ri, ci, row.get(h, ""))
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = thin
            if ri % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F2F2F2")
    cws = wb.create_sheet("Column Rules")
    cws["A1"] = "Ownership: Independent | Private | Public | Subsidiary | Cooperative | Joint Venture"
    cws["A2"] = "Business Type: Distributor | Specialty Distributor"
    cws["A3"] = "Sources: web-verified known map + Wikipedia + name heuristics"
    cws["A4"] = f"Updated: {datetime.now(timezone.utc).isoformat()} | rows={len(rows)}"
    wb.save(path)


async def main() -> int:
    rows = load_rows(XLSX)
    print(f"Loaded {len(rows)}")
    known = {**KNOWN, **EXTRA}

    # Wikipedia for remaining
    need_wiki = [r for r in rows if r["Company"] not in known]
    print(f"Wikipedia lookup for {len(need_wiki)}...")
    wiki_map: dict[str, str] = {}
    async with httpx.AsyncClient(
        headers={"User-Agent": "vendor-intel-adc-own/1.0"},
        timeout=httpx.Timeout(8.0, connect=4.0),
    ) as http:
        for i in range(0, len(need_wiki), 12):
            chunk = need_wiki[i : i + 12]
            try:
                texts = await asyncio.wait_for(
                    asyncio.gather(*[wiki_hint(http, r["Company"]) for r in chunk]),
                    timeout=25.0,
                )
            except asyncio.TimeoutError:
                texts = [""] * len(chunk)
            for r, t in zip(chunk, texts):
                wiki_map[r["Company"]] = t or ""
            print(f"  wiki {i+1}-{i+len(chunk)}/{len(need_wiki)}", flush=True)

    changed = 0
    for r in rows:
        name = r["Company"]
        before = (r.get("Ownership"), r.get("Business Type"))
        if name in known:
            o, b = known[name]
        else:
            o = normalize_ownership(r.get("Ownership") or "Private")
            b = normalize_bt(
                r.get("Business Type") or "Distributor",
                name,
                r.get("Specialty Focus") or "",
                r.get("Core Categories") or "",
            )
            ho, hb = name_heuristic(name)
            if ho:
                o = ho
            if hb:
                b = hb
            w = (wiki_map.get(name) or "").lower()
            if re.search(r"publicly traded|listed on|nyse:|nasdaq:|stock exchange", w):
                o = "Public"
            if re.search(r"subsidiary of|wholly owned|owned by|division of|part of the .+ group", w):
                o = "Subsidiary"
            if re.search(r"cooperative|genossenschaft|mutual", w):
                o = "Cooperative"
            if re.search(r"joint venture|\bjv\b", w):
                o = "Joint Venture"
            # Overused Independent -> Private unless clearly indie and small
            if o == "Independent":
                o = "Private"

        # specialty pharmacy name/focus
        b = normalize_bt(b, name, r.get("Specialty Focus") or "", r.get("Core Categories") or "")
        if "specialty pharmacy" in (r.get("Business Type") or "").lower():
            b = "Specialty Distributor"

        r["Ownership"] = o if o in OWN_OK else "Private"
        r["Business Type"] = b if b in BT_OK else "Distributor"
        if (r["Ownership"], r["Business Type"]) != before:
            changed += 1
            print(f"FIX {name[:42]:42} | {before[0]}->{r['Ownership']} | {before[1]}->{r['Business Type']}")
        src = r.get("Data Sources") or ""
        if "adc_ownership_bt_web" not in src:
            r["Data Sources"] = (src + "; adc_ownership_bt_web").strip("; ")

    rows.sort(key=lambda x: norm(x["Company"]))
    write_xlsx(XLSX, rows)
    summary = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "changed": changed,
        "ownership": dict(Counter(r["Ownership"] for r in rows)),
        "business_type": dict(Counter(r["Business Type"] for r in rows)),
        "samples": [
            {"company": r["Company"], "ownership": r["Ownership"], "bt": r["Business Type"]}
            for r in rows
            if r["Company"]
            in {
                "AAH Pharmaceuticals",
                "Accredo Health Group",
                "CVS Specialty",
                "Suzuken Co., Ltd.",
                "2BScientific Ltd",
                "NOWEDA Apothekergenossenschaft",
                "Red Oak Sourcing specialty",
                "Alliance Healthcare UK",
            }
        ],
    }
    SUMMARY.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    print("\nDONE changed=", changed)
    print("Ownership:", summary["ownership"])
    print("Business Type:", summary["business_type"])
    for s in summary["samples"]:
        print(s)
    return 0


if __name__ == "__main__":
    raise SystemExit(asyncio.run(main()))
