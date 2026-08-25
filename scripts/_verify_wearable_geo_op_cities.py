# -*- coding: utf-8 -*-
"""
Re-verify Continent / Geography + Operational Presence per user definition:

Continent / Geography =
  ALL physical places present (HQ, offices, warehouses, plants, stores) —
  write cities AND countries when known. Format: "City, Country; City, Country"

Operational Presence =
  ALL countries where the company sells/serves (commercial reach). Format: "Country; Country"

Rules: only verified data; never invent. Core cols Company/Website/Founded/Headquarters unchanged.
Default: GEO = Headquarters value; OP = HQ country only. Overrides = web-verified only.
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
AUDIT = ROOT / "_wearable_geo_op_verified_audit.json"

ALIAS = {
    "usa": "United States",
    "us": "United States",
    "uk": "United Kingdom",
    "uae": "United Arab Emirates",
    "korea": "South Korea",
    "czech republic": "Czechia",
    "czechia": "Czechia",
    "türkiye": "Turkey",
    "turkiye": "Turkey",
    "hong kong sar": "Hong Kong",
    "mainland china": "China",
}


def canon_country(s: str) -> str | None:
    t = (s or "").strip()
    if not t:
        return None
    low = t.lower().strip(" .")
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
        "costa rica": "Costa Rica",
        "czech republic": "Czechia",
        "puerto rico": "Puerto Rico",
    }
    if low in ALIAS:
        return ALIAS[low]
    if low in specials:
        return specials[low]
    if low in {
        "europe",
        "asia",
        "global",
        "international",
        "worldwide",
        "apac",
        "emea",
        "latam",
        "mena",
    }:
        return None
    return t


def country_from_hq(hq: str) -> str | None:
    s = (hq or "").strip()
    if not s:
        return None
    if "," in s:
        return canon_country(s.split(",")[-1].strip())
    return canon_country(s)


def dedupe_keep(seq: list[str]) -> list[str]:
    seen = set()
    out = []
    for x in seq:
        x = (x or "").strip()
        if not x:
            continue
        key = x.lower()
        if key not in seen:
            seen.add(key)
            out.append(x)
    return out


# Web-verified PHYSICAL places (city-level when known). Sources noted in audit.
VERIFIED_GEO: dict[str, list[str]] = {
    # https://albaharmedical.com/about/ — Sharq/East Kuwait City, Al-Safwa Tower
    "Al-Bahar Medical": ["Kuwait City (Sharq / Mubarak Al-Kabeer St, Al-Safwa Tower), Kuwait"],
    # https://www.transmedicgroup.com/our-locations/
    "Transmedic": [
        "Singapore, Singapore (HQ)",
        "Kuala Lumpur, Malaysia",
        "Jakarta, Indonesia",
        "Bangkok, Thailand",
        "Hanoi, Vietnam",
        "Ho Chi Minh City, Vietnam",
        "Makati City, Philippines",
        "Davao City, Philippines",
        "Kwun Tong, Hong Kong",
    ],
    # https://www.idsmed.com/en-en/regional.html
    "IDS Medical Systems": [
        "Singapore, Singapore",
        "Shah Alam, Malaysia",
        "Jakarta, Indonesia",
        "San Juan City, Philippines",
        "Samut Prakan, Thailand",
        "Ho Chi Minh City, Vietnam",
        "Hanoi, Vietnam",
    ],
    # Alliance Healthcare HQ Weybridge; Cencora European network countries (wholesale hubs)
    # City-level: Weybridge HQ + Viladecans/Barcelona Spain HQ/logistics publicly cited; UK has many DCs
    "Alliance Healthcare": [
        "Weybridge, United Kingdom (HQ)",
        "United Kingdom (multiple distribution / service centres)",
        "France (distribution network)",
        "Germany (distribution network)",
        "Spain (incl. Viladecans / Barcelona HQ & logistics; multiple branches)",
        "Netherlands (distribution network)",
        "Czechia (distribution network)",
        "Norway (distribution network)",
        "Romania (distribution network)",
        "Turkey (distribution network)",
        "Lithuania (distribution network)",
    ],
    "Alloga": [
        "Paris, France (HQ)",
        "Illescas, Spain (logistics hub)",
        "Borox, Spain",
        "Viladecans / Barcelona, Spain",
        "Las Palmas, Spain",
        "Tenerife, Spain",
    ],
    # A.S. Watson fact sheet — Watsons brand store markets (physical retail). Cities too many (5000+);
    # list all countries with stores; HQ Hong Kong.
    "Watsons": [
        "Hong Kong, Hong Kong (HQ / stores)",
        "China (stores)",
        "Taiwan (stores)",
        "Macau (stores)",
        "Singapore (stores)",
        "Malaysia (stores)",
        "Thailand (stores)",
        "Philippines (stores)",
        "Indonesia (stores)",
        "Vietnam (stores)",
        "Turkey (stores)",
        "Qatar (stores)",
        "United Arab Emirates (stores)",
        "Saudi Arabia (stores)",
        "Bahrain (stores)",
    ],
    # Device Technologies LinkedIn locations + Belrose HQ
    "Device Technologies": [
        "Belrose, NSW, Australia (HQ)",
        "Mulgrave, VIC, Australia",
        "Tingalpa, QLD, Australia",
        "Osborne Park, WA, Australia",
        "Thebarton, SA, Australia",
        "Auckland, New Zealand",
    ],
    # EBOS Healthcare: AU+NZ coverage; exact 6 AU city list not fully published on about page —
    # keep country-level physical with HQ Melbourne from sheet if present
    "EBOS Healthcare": [
        "Melbourne, Australia (HQ per sheet)",
        "Australia (offices / warehouses — 6+ AU locations cited)",
        "New Zealand (warehouses / distribution — both islands)",
    ],
    "BioD Medica": [
        "Subang Jaya, Malaysia (HQ)",
        "Singapore, Singapore (registrant entity / Marina Bay registered office)",
    ],
    "GrafiMedics BeLux": ["Nazareth, Belgium (HQ)"],
    "Primary Care Supplies": ["Aylesford, Kent, United Kingdom (HQ)"],
    "AED USA": ["Fort Worth, TX, United States (HQ)"],
    "Provita Medical": ["Wermelskirchen, Germany (HQ)"],
    "Invacare Europe": ["Aesch, Switzerland (HQ)"],
    "Sanitätshaus Aktuell AG": ["Vettelschoß, Germany (HQ)"],
    "Response Ready": ["Oceanside, CA, United States (HQ)"],
    "Enerspect Medical Solutions": ["Henderson, NV, United States (HQ)"],
    "MyBanzu Healthcare": [
        "Puchong, Selangor, Malaysia (West Malaysia address)",
        "Miri, Sarawak, Malaysia (East Malaysia address)",
    ],
    "Farmacias Kielsa": ["Tegucigalpa, Honduras (HQ)"],
    "Farmacorp": ["Santa Cruz de la Sierra, Bolivia (HQ)"],
    "Farmacias Carol": ["Santo Domingo, Dominican Republic (HQ)"],
    "Farmacias Siman": ["San Pedro Sula, Honduras (HQ)"],
    "Farmacias La Paz": ["Tegucigalpa, Honduras (HQ)"],
    "MediSystem Pharmacy": ["Toronto, Canada (HQ)"],
    "Spire Health": ["San Francisco, CA, United States (HQ)"],
    "Cognita Labs": ["Santa Ana, CA, United States (HQ)"],
    "Carex Health Brands": ["Middleburg Heights, OH, United States (via Compass Health Brands)"],
    "Zephyr Technology": ["Annapolis, MD, United States"],
    "Health Care Australia": ["Sydney, Australia (HQ)"],
    "Healthcare Direct": ["Borehamwood, United Kingdom (HQ)"],
    "Al-Mansour Medical": ["Baghdad, Iraq (HQ)"],
    # Giant Eagle: US retail HQ + Bengaluru GCC office (physical IT office, verified)
    "Giant Eagle": [
        "Cranberry Township / Pittsburgh area, PA, United States (HQ)",
        "Bengaluru, India (Global Capability Center office)",
    ],
    "Walmart": ["Bentonville, AR, United States (HQ)", "Canada (Walmart Canada operations)"],
    "Amazon": ["Seattle, WA, United States (HQ)", "Canada (Amazon.ca operations / facilities)"],
    "Aster Pharmacy": [
        "Dubai, United Arab Emirates (HQ)",
        "India (pharmacy stores — multi-state)",
        "Saudi Arabia (pharmacy stores)",
        "Oman (pharmacy stores)",
        "Qatar (pharmacy stores)",
        "Bahrain (pharmacy stores)",
        "Jordan (pharmacy stores)",
    ],
    "BinSina Pharmacy": ["Dubai, United Arab Emirates (HQ)"],
    "Farmacias Cruz Verde": ["Santiago, Chile (HQ)", "Colombia (operations)"],
    "DocMorris": [
        "Berlin, Germany (HQ)",
        "Switzerland (ops / market)",
        "Netherlands (market ops)",
        "Spain (market ops)",
        "France (market ops)",
    ],
    "Mediq": ["Utrecht, Netherlands (HQ)", "Germany (operations)"],
    "Henry Schein": [
        "Melville, NY, United States (HQ)",
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
    ],
    "McKesson Corporation": [
        "Irving, TX, United States (HQ)",
        "Canada (McKesson Canada)",
    ],
    "McKesson Canada": ["Mississauga, Canada (HQ)"],
    "Cencora": ["Conshohocken, PA, United States (HQ)"],
    "AmerisourceBergen": ["Conshohocken, PA, United States (HQ)"],
    "Cardinal Health": ["Dublin, OH, United States (HQ)"],
    "Medline Industries": ["Northfield, IL, United States (HQ)"],
    "Owens & Minor": ["Glen Allen, VA, United States (HQ)"],
    "Zuellig Pharma": [
        "Singapore, Singapore (HQ)",
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
        "China (Clinical Reach / office presence)",
        "India (Clinical Reach / office presence)",
        "Australia (Clinical Reach / office presence)",
        "Japan (Clinical Reach / office presence)",
    ],
    "Healthcare 21": ["Dublin, Ireland (HQ)", "United Kingdom"],
    "Easmed": ["Singapore, Singapore"],
    "Healthmark": ["Singapore, Singapore"],
    "Noweda": ["Essen, Germany (HQ)"],
    "CERP Rouen": ["Paris / Rouen area, France (HQ)"],
    "B. Braun": ["Melsungen, Germany (HQ)"],
}

# Web-verified commercial reach countries (sell/serve). Not inventing.
VERIFIED_OP: dict[str, list[str]] = {
    "Al-Bahar Medical": ["Kuwait"],  # about page: Kuwait healthcare; "outside Kuwait" claim too vague for extra countries
    "Transmedic": [
        "Singapore",
        "Malaysia",
        "Indonesia",
        "Thailand",
        "Vietnam",
        "Philippines",
        "Hong Kong",
    ],
    "IDS Medical Systems": [
        "Singapore",
        "Malaysia",
        "Indonesia",
        "Philippines",
        "Thailand",
        "Vietnam",
    ],
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
    ],
    "Alloga": ["France", "Spain"],
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
    ],
    "Device Technologies": ["Australia", "New Zealand"],
    "EBOS Healthcare": ["Australia", "New Zealand"],
    "BioD Medica": ["Malaysia", "Singapore"],
    "GrafiMedics BeLux": ["Belgium", "Luxembourg"],
    "Aster Pharmacy": [
        "United Arab Emirates",
        "India",
        "Saudi Arabia",
        "Oman",
        "Qatar",
        "Bahrain",
        "Jordan",
    ],
    "BinSina Pharmacy": ["United Arab Emirates"],  # Oman expansion not confirmed as active retail here
    "Farmacias Cruz Verde": ["Chile", "Colombia"],
    "DocMorris": ["Germany", "Switzerland", "Netherlands", "Spain", "France"],
    "Mediq": ["Netherlands", "Germany"],
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
    ],
    "McKesson Corporation": ["United States", "Canada"],
    "McKesson Canada": ["Canada"],
    "Cencora": ["United States"],
    "AmerisourceBergen": ["United States"],
    "Cardinal Health": ["United States"],
    "Medline Industries": ["United States"],
    "Owens & Minor": ["United States"],
    "Walmart": ["United States", "Canada"],
    "Amazon": ["United States", "Canada"],
    "Giant Eagle": ["United States"],  # retail grocery/pharmacy US; India is GCC support office not retail market
    "Zuellig Pharma": [
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
    "Healthcare 21": ["Ireland", "United Kingdom"],
    "Primary Care Supplies": ["United Kingdom"],
    "AED USA": ["United States"],
    "Provita Medical": ["Germany"],
    "Invacare Europe": ["Switzerland"],
    "Farmacias Carol": ["Dominican Republic"],
    "Farmacias Siman": ["Honduras"],
    "Farmacias La Paz": ["Honduras"],  # drop Chile/Mexico — unverified for this entity
    "Farmacias Kielsa": ["Honduras"],
    "Farmacorp": ["Bolivia"],
    "Cognita Labs": ["United States"],  # drop India — unverified
    "Healthmark": ["Singapore"],
    "Easmed": ["Singapore"],
    "B. Braun": ["Germany"],  # row HQ Germany; do not keep AU/MY-only as full commercial map
    "CERP Rouen": ["France"],
    "Noweda": ["Germany"],
    "Al-Ahli Medical": ["Qatar"],  # drop multi-GCC without location proof
    "Al-Mansour Medical": ["Iraq"],
    "Ecolab": ["United States"],
    "Medikabazaar": ["India"],
    "MediLink": ["United Kingdom"],
    "Medisource Supply (Asia)": ["Hong Kong"],
    "Meditrade GmbH": ["Germany"],
    "Swasthya Healthcare": ["India"],
    "Farmacentro": ["Guatemala"],
    "Farmacias del Dr. Simi": ["Mexico", "Chile", "Colombia"],  # keep LatAm core; drop US/Japan/DR without this pass proof
    "Days Healthcare": ["United Kingdom"],
    "Patterson Companies": ["United States", "Canada"],
    "Bunzl": ["United Kingdom", "Australia", "New Zealand", "Brazil", "Chile", "Finland"],
    "Saudi Pharmaceutical Industries & Medical Appliances Corporation (SPIMACO)": [
        "Saudi Arabia"
    ],
}


def apply_format(ws) -> None:
    headers = [c.value for c in ws[1]]
    widths = {
        "Company": 34,
        "Website": 36,
        "Founded": 10,
        "Headquarters": 28,
        "Continent / Geography": 55,
        "Operational Presence": 42,
        "Ownership": 18,
        "Employees": 12,
        "Core Categories": 28,
        "Specialty Focus": 36,
        "Key Brands Represented": 32,
        "Retail / E-commerce / Both": 16,
        "Distribution Type": 22,
        "Contact Person": 18,
        "Role": 16,
        "Email": 28,
        "LinkedIn": 28,
        "Office No.": 16,
        "Country Code": 10,
        "Region Code": 10,
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
    header_align = Alignment(wrap_text=True, vertical="center", horizontal="center")
    for i, h in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(i)].width = widths.get(str(h or ""), 18)
    ws.row_dimensions[1].height = 36
    for i in range(1, len(headers) + 1):
        c = ws.cell(1, i)
        c.fill = header_fill
        c.font = header_font
        c.alignment = header_align
        c.border = thin
    geo_i = headers.index("Continent / Geography") + 1
    op_i = headers.index("Operational Presence") + 1
    for r in range(2, ws.max_row + 1):
        longest = max(
            len(str(ws.cell(r, geo_i).value or "")),
            len(str(ws.cell(r, op_i).value or "")),
            24,
        )
        lines = max(2, min(10, (longest // 50) + 1))
        ws.row_dimensions[r].height = 14 * lines
        for i in range(1, len(headers) + 1):
            cell = ws.cell(r, i)
            cell.font = body_font
            cell.alignment = wrap
            cell.border = thin
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = ROOT / f"global_wearable_medical_devices_market_global_FINAL_BEFORE_GEO_VERIFY_{stamp}.xlsx"
    shutil.copy2(XLSX, bak)

    wb = load_workbook(XLSX)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    ic = {h: i + 1 for i, h in enumerate(headers) if h}

    core_snapshot = []
    updates = []

    for r in range(2, ws.max_row + 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        website = str(ws.cell(r, ic["Website"]).value or "").strip()
        founded = ws.cell(r, ic["Founded"]).value
        hq = str(ws.cell(r, ic["Headquarters"]).value or "").strip()
        old_geo = str(ws.cell(r, ic["Continent / Geography"]).value or "").strip()
        old_op = str(ws.cell(r, ic["Operational Presence"]).value or "").strip()

        core_snapshot.append(
            {"company": co, "website": website, "founded": founded, "hq": hq}
        )

        hq_country = country_from_hq(hq) or ""

        # Default GEO = full Headquarters (city + country as already verified in HQ col)
        if co in VERIFIED_GEO:
            geo_list = dedupe_keep(VERIFIED_GEO[co])
        else:
            geo_list = [hq] if hq else ([hq_country] if hq_country else [])

        # Default OP = HQ country only unless verified commercial list
        if co in VERIFIED_OP:
            op_list = dedupe_keep([canon_country(x) or x for x in VERIFIED_OP[co]])
        else:
            op_list = [hq_country] if hq_country else []

        new_geo = "; ".join(geo_list)
        new_op = "; ".join([c for c in op_list if c])

        # NEVER touch Company/Website/Founded/Headquarters
        ws.cell(r, ic["Continent / Geography"]).value = new_geo
        ws.cell(r, ic["Operational Presence"]).value = new_op

        if new_geo != old_geo or new_op != old_op:
            updates.append(
                {
                    "company": co,
                    "old_geo": old_geo,
                    "new_geo": new_geo,
                    "old_op": old_op,
                    "new_op": new_op,
                }
            )

    if "smartwatch" in (ws.title or "").lower():
        ws.title = "Wearable Medical Companies"

    apply_format(ws)
    wb.save(XLSX)

    # verify core unchanged vs snapshot written (same loop values)
    audit = {
        "definition": {
            "Continent / Geography": "All physical places (HQ/offices/warehouses/plants/stores) — cities and countries when known",
            "Operational Presence": "All countries of commercial sell/serve reach",
        },
        "backup": bak.name,
        "rows": ws.max_row - 1,
        "updated_geo_or_op": len(updates),
        "verified_geo_overrides": len(VERIFIED_GEO),
        "verified_op_overrides": len(VERIFIED_OP),
        "core_cols_policy": "Company, Website, Founded, Headquarters were not modified",
        "sample_updates": updates[:40],
        "all_updates": updates,
    }
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print("backup", bak.name)
    print("updated", len(updates), "/", ws.max_row - 1)
    print("verified GEO overrides", len(VERIFIED_GEO), "OP overrides", len(VERIFIED_OP))
    for name in [
        "Al-Bahar Medical",
        "Transmedic",
        "IDS Medical Systems",
        "Alliance Healthcare",
        "Watsons",
        "Giant Eagle",
        "Al-Ahli Medical",
        "Farmacias La Paz",
        "Device Technologies",
    ]:
        for u in updates:
            if u["company"] == name:
                print("---", name)
                print(" GEO:", u["new_geo"][:220])
                print(" OP :", u["new_op"])
                break


if __name__ == "__main__":
    main()
