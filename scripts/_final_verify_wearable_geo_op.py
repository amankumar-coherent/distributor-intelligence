# -*- coding: utf-8 -*-
"""
Verify + fill Continent / Geography and Operational Presence.

DEFINITIONS (user):
- Continent / Geography = countries with PHYSICAL ops
  (HQ, offices, warehouses, plants, stores). Format: AF(South Africa, Namibia)
- Operational Presence = countries of COMMERCIAL sell/serve reach.
  Format: South Africa; Namibia

Only real/verified countries. Never invent.
Does NOT modify Company, Website, Founded, Headquarters.
"""
from __future__ import annotations

import json
import re
import shutil
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(
    r"D:\vendor-intel\output\web_expand\global_wearable_medical_devices_market_global"
)
XLSX = ROOT / "global_wearable_medical_devices_market_global_FINAL.xlsx"
AUDIT = ROOT / "_wearable_geo_op_final_verify_audit.json"

COUNTRY_REGION = {
    "United States": "NA",
    "Canada": "NA",
    "Mexico": "NA",
    "Brazil": "LATAM",
    "Argentina": "LATAM",
    "Chile": "LATAM",
    "Colombia": "LATAM",
    "Peru": "LATAM",
    "Bolivia": "LATAM",
    "Ecuador": "LATAM",
    "Uruguay": "LATAM",
    "Paraguay": "LATAM",
    "Venezuela": "LATAM",
    "Guatemala": "LATAM",
    "Honduras": "LATAM",
    "El Salvador": "LATAM",
    "Nicaragua": "LATAM",
    "Costa Rica": "LATAM",
    "Panama": "LATAM",
    "Dominican Republic": "LATAM",
    "Puerto Rico": "LATAM",
    "United Kingdom": "EU",
    "Ireland": "EU",
    "France": "EU",
    "Germany": "EU",
    "Spain": "EU",
    "Portugal": "EU",
    "Italy": "EU",
    "Netherlands": "EU",
    "Belgium": "EU",
    "Luxembourg": "EU",
    "Switzerland": "EU",
    "Austria": "EU",
    "Sweden": "EU",
    "Norway": "EU",
    "Denmark": "EU",
    "Finland": "EU",
    "Poland": "EU",
    "Czechia": "EU",
    "Romania": "EU",
    "Hungary": "EU",
    "Greece": "EU",
    "Lithuania": "EU",
    "Latvia": "EU",
    "Estonia": "EU",
    "Slovakia": "EU",
    "Slovenia": "EU",
    "Croatia": "EU",
    "Bulgaria": "EU",
    "Turkey": "EU",
    "South Africa": "AF",
    "Namibia": "AF",
    "Botswana": "AF",
    "Mauritius": "AF",
    "Nigeria": "AF",
    "Kenya": "AF",
    "Ghana": "AF",
    "Egypt": "AF",
    "Morocco": "AF",
    "Saudi Arabia": "ME",
    "United Arab Emirates": "ME",
    "Qatar": "ME",
    "Kuwait": "ME",
    "Bahrain": "ME",
    "Oman": "ME",
    "Jordan": "ME",
    "Iraq": "ME",
    "Israel": "ME",
    "Japan": "APAC",
    "China": "APAC",
    "Hong Kong": "APAC",
    "Macau": "APAC",
    "Taiwan": "APAC",
    "South Korea": "APAC",
    "India": "APAC",
    "Singapore": "APAC",
    "Malaysia": "APAC",
    "Thailand": "APAC",
    "Indonesia": "APAC",
    "Philippines": "APAC",
    "Vietnam": "APAC",
    "Myanmar": "APAC",
    "Cambodia": "APAC",
    "Brunei": "APAC",
    "Australia": "OC",
    "New Zealand": "OC",
}
REGION_ORDER = ["NA", "LATAM", "EU", "AF", "ME", "APAC", "OC"]

# PHYSICAL ops countries (verified). Sources in audit notes.
VERIFIED_GEO: dict[str, list[str]] = {
    "Transmedic": [
        "Singapore",
        "Malaysia",
        "Indonesia",
        "Thailand",
        "Vietnam",
        "Philippines",
        "Hong Kong",
    ],  # locations page
    "IDS Medical Systems": [
        "Singapore",
        "Malaysia",
        "Indonesia",
        "Philippines",
        "Thailand",
        "Vietnam",
    ],  # idsmed regional
    "Alliance Healthcare": [
        "United Kingdom",
        "France",
        "Germany",
        "Spain",
        "Netherlands",
        "Czechia",
        "Norway",
        "Romania",
        "Turkey",
        "Lithuania",
    ],  # Cencora / Alliance European network
    "Alloga": ["France", "Spain"],  # FR HQ + ES warehouses
    "Watsons": [
        "Hong Kong",
        "China",
        "Taiwan",
        "Macau",
        "Singapore",
        "Malaysia",
        "Thailand",
        "Philippines",
        "Indonesia",
        "Vietnam",
        "Turkey",
        "Qatar",
        "United Arab Emirates",
        "Saudi Arabia",
        "Bahrain",
    ],  # A.S. Watson Watsons brand store markets
    "Aster Pharmacy": [
        "United Arab Emirates",
        "India",
        "Saudi Arabia",
        "Oman",
        "Qatar",
        "Bahrain",
        "Jordan",
    ],  # pharmacy stores in 7 countries
    "BinSina Pharmacy": ["United Arab Emirates", "Oman"],  # UAE HQ + Oman stores
    "Device Technologies": [
        "Australia",
        "New Zealand",
        "Singapore",
        "Malaysia",
        "Thailand",
        "Philippines",
        "Hong Kong",
    ],  # AU/NZ + LinkedIn Asia offices
    "EBOS Healthcare": ["Australia", "New Zealand"],
    "BioD Medica": ["Malaysia", "Singapore"],
    "GrafiMedics BeLux": ["Belgium"],  # Nazareth HQ; LU commercial only
    "DocMorris": ["Germany", "Switzerland", "Netherlands", "Spain", "France"],
    "Mediq": ["Netherlands", "Germany"],
    "Farmacias Cruz Verde": ["Chile", "Colombia"],
    "Farmacias del Dr. Simi": [
        "Mexico",
        "Chile",
        "Colombia",
        "El Salvador",
        "United States",
    ],  # stores/wiki + press (MX/CL/CO/SV/US)
    "Healthcare 21": ["Ireland", "United Kingdom"],
    "Henry Schein": [
        "United States",
        "Canada",
        "United Kingdom",
        "Germany",
        "France",
        "Spain",
        "Netherlands",
        "Australia",
        "New Zealand",
        "Brazil",
        "Singapore",
    ],  # documented major markets (not inventing full 34)
    "McKesson Corporation": ["United States", "Canada"],  # NA core; Europe exited
    "McKesson Canada": ["Canada"],
    "Cencora": ["United States"],
    "AmerisourceBergen": ["United States"],
    "Cardinal Health": ["United States"],  # do not invent 40-country list
    "Medline Industries": ["United States"],
    "Owens & Minor": ["United States"],
    "Patterson Companies": ["United States", "Canada"],
    "Walmart": ["United States", "Canada"],
    "Amazon": ["United States", "Canada"],
    "Giant Eagle": ["United States", "India"],  # US retail + Bengaluru GCC office
    "Zuellig Pharma": [
        # 13 distribution markets (warehouses) + 4 Clinical Reach office markets
        "Singapore",
        "Malaysia",
        "Thailand",
        "Philippines",
        "Vietnam",
        "Indonesia",
        "Hong Kong",
        "Taiwan",
        "South Korea",
        "Brunei",
        "Cambodia",
        "Myanmar",
        "China",
        "India",
        "Australia",
        "Japan",
    ],
    "Bunzl": [
        # Bunzl business case / Wikipedia countries of operation (physical)
        "United Kingdom",
        "Ireland",
        "United States",
        "Canada",
        "Mexico",
        "Puerto Rico",
        "Germany",
        "France",
        "Italy",
        "Spain",
        "Netherlands",
        "Belgium",
        "Denmark",
        "Norway",
        "Finland",
        "Switzerland",
        "Austria",
        "Czechia",
        "Hungary",
        "Romania",
        "Poland",
        "Israel",
        "Turkey",
        "Brazil",
        "Chile",
        "Colombia",
        "Peru",
        "Uruguay",
        "Australia",
        "New Zealand",
        "China",
        "Singapore",
    ],
    "Invacare Europe": ["Switzerland"],
    "Primary Care Supplies": ["United Kingdom"],
    "Al-Bahar Medical": ["Kuwait"],
    "Al-Ahli Medical": ["Qatar"],
    "Al-Mansour Medical": ["Iraq"],
    "Farmacias Carol": ["Dominican Republic"],
    "Farmacias Siman": ["Honduras"],
    "Farmacias La Paz": ["Honduras"],
    "Farmacias Kielsa": ["Honduras"],
    "Farmacorp": ["Bolivia"],
    "Noweda": ["Germany"],
    "CERP Rouen": ["France"],
    "B. Braun": ["Germany"],
    "Easmed": ["Singapore"],
    "Healthmark": ["Singapore"],
}

# COMMERCIAL reach countries (verified). May equal GEO or be narrower/wider when proven.
VERIFIED_OP: dict[str, list[str]] = {
    "Transmedic": VERIFIED_GEO["Transmedic"],
    "IDS Medical Systems": VERIFIED_GEO["IDS Medical Systems"],
    "Alliance Healthcare": VERIFIED_GEO["Alliance Healthcare"],
    "Alloga": ["France", "Spain"],
    "Watsons": VERIFIED_GEO["Watsons"],
    "Aster Pharmacy": VERIFIED_GEO["Aster Pharmacy"],
    "BinSina Pharmacy": ["United Arab Emirates", "Oman"],
    "Device Technologies": VERIFIED_GEO["Device Technologies"],
    "EBOS Healthcare": ["Australia", "New Zealand"],
    "BioD Medica": ["Malaysia", "Singapore"],
    "GrafiMedics BeLux": ["Belgium", "Luxembourg"],  # BeLux sell/serve
    "DocMorris": VERIFIED_GEO["DocMorris"],
    "Mediq": ["Netherlands", "Germany"],
    "Farmacias Cruz Verde": ["Chile", "Colombia"],
    "Farmacias del Dr. Simi": VERIFIED_GEO["Farmacias del Dr. Simi"],
    "Healthcare 21": ["Ireland", "United Kingdom"],
    "Henry Schein": VERIFIED_GEO["Henry Schein"],
    "McKesson Corporation": ["United States", "Canada"],
    "McKesson Canada": ["Canada"],
    "Cencora": ["United States"],
    "AmerisourceBergen": ["United States"],
    "Cardinal Health": ["United States"],
    "Medline Industries": ["United States"],
    "Owens & Minor": ["United States"],
    "Patterson Companies": ["United States", "Canada"],
    "Walmart": ["United States", "Canada"],
    "Amazon": ["United States", "Canada"],
    "Giant Eagle": ["United States"],  # retail US only; India is support office
    "Zuellig Pharma": VERIFIED_GEO["Zuellig Pharma"],
    "Bunzl": VERIFIED_GEO["Bunzl"],
    "Invacare Europe": ["Switzerland"],
    "Primary Care Supplies": ["United Kingdom"],
    "Al-Bahar Medical": ["Kuwait"],
    "Al-Ahli Medical": ["Qatar"],
    "Al-Mansour Medical": ["Iraq"],
    "Farmacias Carol": ["Dominican Republic"],
    "Farmacias Siman": ["Honduras"],
    "Farmacias La Paz": ["Honduras"],
    "Farmacias Kielsa": ["Honduras"],
    "Farmacorp": ["Bolivia"],
    "Noweda": ["Germany"],
    "CERP Rouen": ["France"],
    "B. Braun": ["Germany"],
    "Easmed": ["Singapore"],
    "Healthmark": ["Singapore"],
    "Cognita Labs": ["United States"],
    "Ecolab": ["United States"],
    "Days Healthcare": ["United Kingdom"],
}


def country_from_hq(hq: str) -> str | None:
    s = (hq or "").strip()
    if not s:
        return None
    if "," in s:
        last = s.split(",")[-1].strip()
        return canon(last)
    return canon(s)


def canon(token: str) -> str | None:
    t = (token or "").strip()
    if not t:
        return None
    low = t.lower()
    alias = {
        "usa": "United States",
        "us": "United States",
        "uk": "United Kingdom",
        "uae": "United Arab Emirates",
        "korea": "South Korea",
        "czech republic": "Czechia",
        "czechia": "Czechia",
        "hong kong sar": "Hong Kong",
        "türkiye": "Turkey",
        "turkiye": "Turkey",
    }
    if low in alias:
        return alias[low]
    specials = {
        "united states": "United States",
        "united kingdom": "United Kingdom",
        "united arab emirates": "United Arab Emirates",
        "south africa": "South Africa",
        "south korea": "South Korea",
        "new zealand": "New Zealand",
        "hong kong": "Hong Kong",
        "saudi arabia": "Saudi Arabia",
        "dominican republic": "Dominican Republic",
        "el salvador": "El Salvador",
        "puerto rico": "Puerto Rico",
    }
    if low in specials:
        return specials[low]
    if t in COUNTRY_REGION:
        return t
    return t if re.search(r"[A-Za-z]", t) else None


def dedupe(seq: list[str]) -> list[str]:
    seen = set()
    out = []
    for x in seq:
        c = canon(x)
        if c and c not in seen:
            seen.add(c)
            out.append(c)
    return out


def format_geo(countries: list[str]) -> str:
    groups: dict[str, list[str]] = {}
    for c in countries:
        # Watsons/Aster GCC: Turkey with ME block when alongside GCC
        region = COUNTRY_REGION.get(c, "APAC")
        if c in {
            "Saudi Arabia",
            "United Arab Emirates",
            "Qatar",
            "Kuwait",
            "Bahrain",
            "Oman",
            "Jordan",
            "Iraq",
        }:
            region = "ME"
        # Turkey with Watsons-style ME retail → ME; Alliance Europe → EU (default map)
        groups.setdefault(region, [])
        if c not in groups[region]:
            groups[region].append(c)
    # Move Turkey to ME if group already has other ME countries (Watsons case)
    if "EU" in groups and "Turkey" in groups["EU"] and "ME" in groups:
        groups["EU"].remove("Turkey")
        if not groups["EU"]:
            del groups["EU"]
        if "Turkey" not in groups["ME"]:
            groups["ME"].append("Turkey")
    parts = []
    for region in REGION_ORDER:
        if region in groups and groups[region]:
            parts.append(f"{region}({', '.join(groups[region])})")
    for region, clist in groups.items():
        if region not in REGION_ORDER and clist:
            parts.append(f"{region}({', '.join(clist)})")
    return "; ".join(parts)


def format_op(countries: list[str]) -> str:
    return "; ".join(countries)


def apply_sheet_format(ws) -> None:
    headers = [c.value for c in ws[1]]
    widths = {
        "Company": 34,
        "Website": 36,
        "Founded": 10,
        "Headquarters": 28,
        "Continent / Geography": 52,
        "Operational Presence": 48,
        "Ownership": 18,
        "Employees": 12,
        "Core Categories": 28,
        "Specialty Focus": 36,
        "Key Brands Represented": 32,
        "Retail / E-commerce / Both": 16,
        "Distribution Type": 22,
        "Summary": 48,
    }
    thin = Border(
        left=Side(style="thin", color="D0D5DD"),
        right=Side(style="thin", color="D0D5DD"),
        top=Side(style="thin", color="D0D5DD"),
        bottom=Side(style="thin", color="D0D5DD"),
    )
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    body_font = Font(name="Calibri", size=10)
    wrap = Alignment(wrap_text=True, vertical="top", horizontal="left")
    for i, h in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(i)].width = widths.get(str(h or ""), 16)
    ws.row_dimensions[1].height = 32
    for i in range(1, len(headers) + 1):
        cell = ws.cell(1, i)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = thin
    geo_i = headers.index("Continent / Geography") + 1
    op_i = headers.index("Operational Presence") + 1
    for r in range(2, ws.max_row + 1):
        longest = max(
            len(str(ws.cell(r, geo_i).value or "")),
            len(str(ws.cell(r, op_i).value or "")),
            20,
        )
        ws.row_dimensions[r].height = 15 * max(2, min(8, longest // 48 + 1))
        for i in range(1, len(headers) + 1):
            cell = ws.cell(r, i)
            cell.font = body_font
            cell.alignment = wrap
            cell.border = thin
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = ROOT / f"global_wearable_medical_devices_market_global_FINAL_BEFORE_FINAL_GEO_VERIFY_{stamp}.xlsx"
    shutil.copy2(XLSX, bak)

    wb = load_workbook(XLSX)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    ic = {h: i + 1 for i, h in enumerate(headers) if h}

    updates = []
    for r in range(2, ws.max_row + 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        hq = str(ws.cell(r, ic["Headquarters"]).value or "").strip()
        old_geo = str(ws.cell(r, ic["Continent / Geography"]).value or "").strip()
        old_op = str(ws.cell(r, ic["Operational Presence"]).value or "").strip()
        hq_c = country_from_hq(hq)

        if co in VERIFIED_GEO:
            geo_list = dedupe(VERIFIED_GEO[co])
        else:
            geo_list = [hq_c] if hq_c else []

        if co in VERIFIED_OP:
            op_list = dedupe(VERIFIED_OP[co])
        else:
            op_list = [hq_c] if hq_c else []

        # Ensure HQ country always in GEO
        if hq_c and hq_c not in geo_list:
            geo_list = [hq_c] + geo_list

        new_geo = format_geo(geo_list)
        new_op = format_op(op_list)

        # Never touch core identity cols
        ws.cell(r, ic["Continent / Geography"]).value = new_geo
        ws.cell(r, ic["Operational Presence"]).value = new_op
        if new_geo != old_geo or new_op != old_op:
            updates.append({"company": co, "geo": new_geo, "op": new_op, "old_geo": old_geo, "old_op": old_op})

    apply_sheet_format(ws)
    wb.save(XLSX)

    # core check vs bak
    wb0 = load_workbook(bak)
    ws0 = wb0.active
    h0 = [c.value for c in ws0[1]]
    ic0 = {h: i + 1 for i, h in enumerate(h0) if h}
    core_bad = 0
    for r in range(2, ws0.max_row + 1):
        for c in ["Company", "Website", "Founded", "Headquarters"]:
            if str(ws0.cell(r, ic0[c]).value or "").strip() != str(ws.cell(r, ic[c]).value or "").strip():
                core_bad += 1

    audit = {
        "definitions": {
            "Continent / Geography": "Countries with physical ops — CODE(Country, Country)",
            "Operational Presence": "Countries of commercial reach — Country; Country",
        },
        "backup": bak.name,
        "updated": len(updates),
        "rows": ws.max_row - 1,
        "verified_geo_companies": len(VERIFIED_GEO),
        "verified_op_companies": len(VERIFIED_OP),
        "core_mismatches": core_bad,
        "samples": [
            u
            for u in updates
            if u["company"]
            in {
                "Al-Bahar Medical",
                "Transmedic",
                "Alliance Healthcare",
                "Watsons",
                "Bunzl",
                "BinSina Pharmacy",
                "Farmacias del Dr. Simi",
                "Giant Eagle",
                "Device Technologies",
                "Zuellig Pharma",
                "AAH Pharmaceuticals",
            }
        ]
        or updates[:15],
    }
    # always include key samples even if unchanged
    key_samples = []
    for r in range(2, ws.max_row + 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        if co in {
            "Al-Bahar Medical",
            "Transmedic",
            "Alliance Healthcare",
            "Watsons",
            "Bunzl",
            "BinSina Pharmacy",
            "Farmacias del Dr. Simi",
            "Giant Eagle",
            "Device Technologies",
            "Zuellig Pharma",
            "AAH Pharmaceuticals",
            "Adcock Ingram",
        }:
            key_samples.append(
                {
                    "company": co,
                    "geo": ws.cell(r, ic["Continent / Geography"]).value,
                    "op": ws.cell(r, ic["Operational Presence"]).value,
                }
            )
    audit["key_samples"] = key_samples
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print("backup", bak.name)
    print("updated", len(updates), "core_bad", core_bad)
    for s in key_samples:
        print("---", s["company"])
        print(" GEO", s["geo"])
        print(" OP ", s["op"])


if __name__ == "__main__":
    main()
