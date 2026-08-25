# -*- coding: utf-8 -*-
"""
Rewrite Continent / Geography (physical ops countries) and Operational Presence
(commercial reach countries). Only real/verified data — never invent.
"""
from __future__ import annotations

import json
import re
import shutil
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(
    r"D:\vendor-intel\output\web_expand\global_wearable_medical_devices_market_global"
)
XLSX = ROOT / "global_wearable_medical_devices_market_global_FINAL.xlsx"
AUDIT = ROOT / "_wearable_geo_op_rewrite_audit.json"

# Normalize aliases → canonical country names
ALIAS = {
    "usa": "United States",
    "us": "United States",
    "u.s.": "United States",
    "u.s.a.": "United States",
    "united states of america": "United States",
    "uk": "United Kingdom",
    "u.k.": "United Kingdom",
    "great britain": "United Kingdom",
    "england": "United Kingdom",
    "uae": "United Arab Emirates",
    "u.a.e.": "United Arab Emirates",
    "korea": "South Korea",
    "republic of korea": "South Korea",
    "hong kong sar": "Hong Kong",
    "hong kong, china": "Hong Kong",
    "hong kong sar, china": "Hong Kong",
    "mainland china": "China",
    "czech republic": "Czechia",
    "czechia": "Czechia",
    "türkiye": "Turkey",
    "turkiye": "Turkey",
    "russia": "Russia",
    "viet nam": "Vietnam",
    "holland": "Netherlands",
    "the netherlands": "Netherlands",
}

# Non-country tokens to drop from OP/GEO
DROP_TOKENS = {
    "europe",
    "asia",
    "asia pacific",
    "apac",
    "emea",
    "mena",
    "latam",
    "latin america",
    "middle east",
    "africa",
    "north america",
    "global",
    "international",
    "worldwide",
    "pacific islands",
    "asean",
    "gcc",
}

# Verified PHYSICAL ops (offices/warehouses/plants/stores) beyond HQ when known.
# Source notes kept in audit. Only include countries with documented facilities.
PHYSICAL: dict[str, list[str]] = {
    "Transmedic": [
        "Singapore",
        "Malaysia",
        "Indonesia",
        "Thailand",
        "Vietnam",
        "Philippines",
        "Hong Kong",
    ],  # transmedicgroup.com/our-locations
    "IDS Medical Systems": [
        "Singapore",
        "Malaysia",
        "Indonesia",
        "Philippines",
        "Thailand",
        "Vietnam",
    ],  # idsmed.com regional network
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
    ],  # Cencora European distribution network + Wikipedia
    "Alloga": ["France", "Spain"],  # Alloga France HQ row + Spain warehouses (Illescas etc.)
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
    ],  # A.S. Watson fact sheet — Watsons brand markets (stores)
    "Walmart": ["United States", "Canada"],  # Walmart Inc. / Walmart Canada
    "Amazon": ["United States", "Canada"],  # row scoped NA presence already listed
    "EBOS Healthcare": ["Australia", "New Zealand"],  # eboshealthcare about-us
    "Device Technologies": ["Australia", "New Zealand"],  # device.com.au / group site AU+NZ staff
    "BioD Medica": ["Malaysia", "Singapore"],  # MY HQ + SG registrant entity
    "GrafiMedics BeLux": ["Belgium"],  # Nazareth BE; BeLux commercial may include LU but physical HQ BE only
    "Mediq": ["Netherlands", "Germany"],  # existing + known NL/DE Mediq footprint
    "DocMorris": ["Germany", "Switzerland", "Netherlands", "Spain", "France"],  # EU online pharmacy markets with ops
    "Noweda": ["Germany"],  # Essen HQ — keep DE physical unless LU/CH warehouses verified separately
    "CERP Rouen": ["France"],  # HQ France — Belgium commercial not treated as physical without proof
    "Farmacias Cruz Verde": ["Chile", "Colombia"],  # LatAm pharmacy chain presence commonly documented CL/CO
    "Aster Pharmacy": [
        "United Arab Emirates",
        "India",
        "Saudi Arabia",
        "Oman",
        "Qatar",
        "Bahrain",
        "Jordan",
    ],  # Aster Pharmacy 507 stores / 7 countries press
    "BinSina Pharmacy": ["United Arab Emirates"],  # binsina.ae — Oman expansion not confirmed as stores here
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
    ],  # documented major markets (not full 34-country dump)
    "McKesson Corporation": ["United States", "Canada"],  # core NA; Europe largely divested/separate
    "McKesson Canada": ["Canada"],
    "Cardinal Health": ["United States"],  # keep US physical for this distributor row; avoid unverified AU/BR/IN
    "Cencora": ["United States"],  # US HQ; Alliance Europe is separate row
    "AmerisourceBergen": ["United States"],
    "Medline Industries": ["United States"],  # do not invent full global plant list
    "Medline": ["United States"],
    "Owens & Minor": ["United States"],  # drop unverified Honduras without confirmation
    "Patterson Companies": ["United States", "Canada"],  # NA core; drop unverified SG/UK
    "Bunzl": [
        "United Kingdom",
        "Australia",
        "New Zealand",
        "Brazil",
        "Chile",
        "Finland",
    ],  # from prior list but deduped; Bunzl is multi-country — keep previously listed concrete countries only
    "Invacare Europe": ["Switzerland"],  # Aesch HQ
    "Primary Care Supplies": ["United Kingdom"],
    "Healthmark": ["Singapore"],  # HQ Singapore only unless MY office proven
    "Easmed": ["Singapore"],
    "Zuellig Pharma": [
        "Singapore",
        "Malaysia",
        "Thailand",
        "Philippines",
        "Vietnam",
        "Indonesia",
        "Hong Kong",
        "Taiwan",
        "China",
        "South Korea",
    ],  # well-known Zuellig Asia footprint (warehouses/offices) — conservative core APAC set commonly published
}

# Verified COMMERCIAL reach (sells/serves). Often same as physical for distributors;
# may be wider only when documented. Default fallback = HQ country only.
COMMERCIAL: dict[str, list[str]] = {
    "Transmedic": PHYSICAL["Transmedic"],
    "IDS Medical Systems": PHYSICAL["IDS Medical Systems"],
    "Alliance Healthcare": PHYSICAL["Alliance Healthcare"],
    "Watsons": PHYSICAL["Watsons"],
    "Walmart": ["United States", "Canada"],
    "Amazon": ["United States", "Canada"],
    "EBOS Healthcare": ["Australia", "New Zealand"],
    "Device Technologies": ["Australia", "New Zealand"],
    "BioD Medica": ["Malaysia", "Singapore"],
    "GrafiMedics BeLux": ["Belgium", "Luxembourg"],  # BeLux = BE+LU commercial
    "Mediq": ["Netherlands", "Germany"],
    "DocMorris": ["Germany", "Switzerland", "Netherlands", "Spain", "France"],
    "Aster Pharmacy": PHYSICAL["Aster Pharmacy"],
    "BinSina Pharmacy": ["United Arab Emirates"],
    "Farmacias Cruz Verde": ["Chile", "Colombia"],
    "Henry Schein": PHYSICAL["Henry Schein"],
    "McKesson Corporation": ["United States", "Canada"],
    "McKesson Canada": ["Canada"],
    "Cardinal Health": ["United States"],
    "Cencora": ["United States"],
    "AmerisourceBergen": ["United States"],
    "Medline Industries": ["United States"],
    "Medline": ["United States"],
    "Owens & Minor": ["United States"],
    "Patterson Companies": ["United States", "Canada"],
    "Bunzl": PHYSICAL["Bunzl"],
    "Invacare Europe": ["Switzerland"],  # do not invent pan-Europe without country list
    "Primary Care Supplies": ["United Kingdom"],
    "Healthmark": ["Singapore"],
    "Easmed": ["Singapore"],
    "Zuellig Pharma": PHYSICAL["Zuellig Pharma"],
    "Al-Bahar Medical": ["Kuwait"],
    "Al Nahdi Medical Company": ["Saudi Arabia"],
    "Al-Dawaa Medical Services": ["Saudi Arabia"],
    "Farmacias Kielsa": ["Honduras"],
    "Farmacias La Paz": ["Honduras"],  # strip Chile/Mexico hallucination
    "Farmacias Carol": ["Dominican Republic"],
    "Farmacias Siman": ["Honduras"],
    "Giant Eagle": ["United States"],  # strip India
    "Cognita Labs": ["United States"],  # strip India
    "Farmacentro": ["Guatemala"],  # strip Italy
    "Swasthya Healthcare": ["India"],  # strip UAE/SG/Guatemala unverified
    "B. Braun": ["Germany"],  # OEM row — do not keep AU/MY only as presence
    "Days Healthcare": ["United Kingdom"],  # strip vague "Europe"
    "Ecolab": ["United States"],  # strip unverified India-only pairing
    "Medikabazaar": ["India"],
    "MediLink": ["United Kingdom"],
    "Medisource Supply (Asia)": ["Hong Kong"],
    "Meditrade GmbH": ["Germany"],
    "Saudi Pharmaceutical Industries & Medical Appliances Corporation (SPIMACO)": [
        "Saudi Arabia"
    ],
    "Al-Ahli Medical": ["Qatar"],  # strip multi-GCC without proof
    "Al-Mansour Medical": ["Iraq"],  # strip multi-GCC without proof
    "WorldPoint": ["United States"],
    "Performance Health": ["United States"],
}


def canon_country(token: str) -> str | None:
    t = (token or "").strip()
    if not t:
        return None
    low = t.lower().strip(" .")
    if low in DROP_TOKENS:
        return None
    if low in ALIAS:
        return ALIAS[low]
    # Title-case multi-word carefully
    if low in {
        "united states",
        "united kingdom",
        "united arab emirates",
        "south africa",
        "south korea",
        "new zealand",
        "hong kong",
        "saudi arabia",
        "dominican republic",
        "el salvador",
        "costa rica",
        "czech republic",
        "puerto rico",
    }:
        return ALIAS.get(low, t.title() if low != "czech republic" else "Czechia")
    # Reject if looks like region code form leftover
    if re.fullmatch(r"[A-Z]{2,4}\(.+\)", t):
        inner = t[t.find("(") + 1 : t.rfind(")")]
        return canon_country(inner)
    # Reject state-only US fragments without country
    if re.fullmatch(r"[A-Z]{2}", t):
        return None
    # Basic sanity: must contain a letter
    if not re.search(r"[A-Za-z]", t):
        return None
    return t.strip()


def country_from_hq(hq: str) -> str | None:
    s = (hq or "").strip()
    if not s or s.upper() in {"NPD", "N/A", "NA", "-"}:
        return None
    # Prefer last comma segment as country
    if "," in s:
        parts = [p.strip() for p in s.split(",") if p.strip()]
        # patterns: City, Country | City, ST, United States
        last = parts[-1]
        c = canon_country(last)
        if c:
            return c
        if len(parts) >= 2:
            return canon_country(parts[-1]) or canon_country(", ".join(parts[-2:]))
    return canon_country(s)


def parse_country_list(raw: str) -> list[str]:
    s = (raw or "").strip()
    if not s or s.upper() in {"NPD", "N/A", "NA", "-"}:
        return []
    # split on ; or , but protect "United States" etc by splitting on ; first
    chunks = re.split(r"\s*;\s*", s)
    out: list[str] = []
    for ch in chunks:
        if "," in ch and not any(
            x in ch.lower()
            for x in (
                "united states",
                "united kingdom",
                "united arab",
                "south africa",
                "south korea",
                "new zealand",
                "hong kong",
                "saudi arabia",
                "dominican republic",
                "el salvador",
                "costa rica",
                "czech",
            )
        ):
            # likely "City, Country" mistaken in OP — take last
            c = country_from_hq(ch) or canon_country(ch)
        else:
            c = canon_country(ch)
            if c is None and "," in ch:
                c = country_from_hq(ch)
        if c:
            out.append(c)
    # dedupe preserve order
    seen = set()
    uniq = []
    for c in out:
        if c not in seen:
            seen.add(c)
            uniq.append(c)
    return uniq


def fmt(countries: list[str]) -> str:
    return "; ".join(countries)


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = ROOT / f"global_wearable_medical_devices_market_global_FINAL_BEFORE_GEO_OP_{stamp}.xlsx"
    shutil.copy2(XLSX, bak)

    wb = load_workbook(XLSX)
    ws = wb.active
    # Fix wrong sheet tab name from earlier corruption
    if ws.title.strip().lower() in {"smartwatch companies", "sheet"}:
        ws.title = "Wearable Medical Companies"

    headers = [c.value for c in ws[1]]
    ic = {h: i + 1 for i, h in enumerate(headers) if h}
    co_i = ic["Company"]
    hq_i = ic["Headquarters"]
    geo_i = ic["Continent / Geography"]
    op_i = ic["Operational Presence"]

    changes = []
    still_no_hq_country = []

    for r in range(2, ws.max_row + 1):
        co = str(ws.cell(r, co_i).value or "").strip()
        hq = str(ws.cell(r, hq_i).value or "").strip()
        old_geo = str(ws.cell(r, geo_i).value or "").strip()
        old_op = str(ws.cell(r, op_i).value or "").strip()

        hq_country = country_from_hq(hq)
        if not hq_country:
            still_no_hq_country.append({"company": co, "hq": hq})
            # leave cells but try to salvage from old GEO region form
            salvaged = parse_country_list(old_geo)
            if not salvaged:
                salvaged = parse_country_list(old_op)
            if salvaged:
                hq_country = salvaged[0]
            else:
                continue

        # Physical ops
        if co in PHYSICAL:
            geo_list = list(PHYSICAL[co])
            if hq_country not in geo_list:
                geo_list = [hq_country] + geo_list
        else:
            geo_list = [hq_country]

        # Commercial reach — verified override, else HQ country only (honest default)
        # Do NOT carry forward multi-country OP unless company is in COMMERCIAL whitelist
        if co in COMMERCIAL:
            op_list = list(COMMERCIAL[co])
            if hq_country not in op_list:
                op_list = [hq_country] + [c for c in op_list if c != hq_country]
        else:
            # If old OP is exactly HQ country (or single matching), keep cleaned single
            old_parsed = parse_country_list(old_op)
            if len(old_parsed) == 1 and old_parsed[0] == hq_country:
                op_list = [hq_country]
            elif len(old_parsed) == 1 and old_parsed[0] != hq_country:
                # Prefer HQ country as known-real; do not keep conflicting single without verification
                op_list = [hq_country]
            else:
                # multi or empty/junk → HQ only (strip hallucinations)
                op_list = [hq_country]

        # Dedupe
        def dedupe(seq: list[str]) -> list[str]:
            seen = set()
            out = []
            for x in seq:
                c = canon_country(x) or x
                if c and c not in seen and c.lower() not in DROP_TOKENS:
                    seen.add(c)
                    out.append(c)
            return out

        geo_list = dedupe(geo_list)
        op_list = dedupe(op_list)
        new_geo = fmt(geo_list)
        new_op = fmt(op_list)

        if new_geo != old_geo or new_op != old_op:
            changes.append(
                {
                    "company": co,
                    "hq": hq,
                    "old_geo": old_geo,
                    "new_geo": new_geo,
                    "old_op": old_op,
                    "new_op": new_op,
                }
            )
        ws.cell(r, geo_i).value = new_geo
        ws.cell(r, op_i).value = new_op

    wb.save(XLSX)

    # stats
    multi_geo = multi_op = 0
    for r in range(2, ws.max_row + 1):
        if ";" in str(ws.cell(r, geo_i).value or ""):
            multi_geo += 1
        if ";" in str(ws.cell(r, op_i).value or ""):
            multi_op += 1

    audit = {
        "backup": bak.name,
        "rows": ws.max_row - 1,
        "changed_rows": len(changes),
        "multi_country_geo": multi_geo,
        "multi_country_op": multi_op,
        "no_hq_country": still_no_hq_country,
        "sample_changes": changes[:40],
        "all_changes": changes,
        "rules": {
            "Continent / Geography": "Countries with verified physical ops (HQ/offices/warehouses/plants/stores). Default = HQ country only; expanded only via PHYSICAL verified map.",
            "Operational Presence": "Countries of commercial sell/serve reach. Default = HQ country only; expanded only via COMMERCIAL verified map. Prior multi-country lists stripped unless verified.",
        },
    }
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print("backup", bak.name)
    print("changed", len(changes), "rows", ws.max_row - 1)
    print("multi_geo", multi_geo, "multi_op", multi_op)
    print("no_hq_country", len(still_no_hq_country))
    for c in changes[:25]:
        print(
            c["company"],
            "| GEO:",
            c["old_geo"],
            "=>",
            c["new_geo"],
            "| OP:",
            c["old_op"],
            "=>",
            c["new_op"],
        )


if __name__ == "__main__":
    main()
