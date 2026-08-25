# -*- coding: utf-8 -*-
"""
Fill Continent / Geography + Operational Presence for Global Flexible Packaging FINAL.

DEFINITIONS (user):
- Continent / Geography = countries with PHYSICAL ops
  (HQ, offices, warehouses, plants, stores). Format: NA(United States); EU(Germany)
- Operational Presence = countries of COMMERCIAL sell/serve reach.
  Format: United States; Germany

Rules: only real/verified countries. Never invent full continent macros.
Baseline = HQ country (physical + commercial minimum). Multi-country only from evidence.
"""
from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

FP = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\global_flexible_packaging_market_global_FINAL.xlsx"
)
AUDIT = FP.parent / "geo_op_fill_audit.json"

COUNTRY_REGION: dict[str, str] = {
    "United States": "NA",
    "Canada": "NA",
    "Mexico": "NA",
    "Puerto Rico": "NA",
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
    "Serbia": "EU",
    "Bosnia and Herzegovina": "EU",
    "Ukraine": "EU",
    "Turkey": "EU",
    "South Africa": "AF",
    "Tunisia": "AF",
    "Angola": "AF",
    "Nigeria": "AF",
    "Kenya": "AF",
    "Egypt": "AF",
    "Morocco": "AF",
    "Saudi Arabia": "ME",
    "United Arab Emirates": "ME",
    "Israel": "ME",
    "Japan": "APAC",
    "China": "APAC",
    "Hong Kong": "APAC",
    "Taiwan": "APAC",
    "South Korea": "APAC",
    "India": "APAC",
    "Singapore": "APAC",
    "Malaysia": "APAC",
    "Thailand": "APAC",
    "Indonesia": "APAC",
    "Philippines": "APAC",
    "Vietnam": "APAC",
    "Australia": "OC",
    "New Zealand": "OC",
}
REGION_ORDER = ["NA", "LATAM", "EU", "AF", "ME", "APAC", "OC"]

# Multi-country PHYSICAL ops — evidence-gated only (never invent remaining “N countries”).
VERIFIED_GEO: dict[str, list[str]] = {
    # Antalis ESG Report 2024 operating entities list
    "Antalis": [
        "Austria", "Belgium", "Bolivia", "Bulgaria", "Canada", "Chile", "Czechia",
        "Denmark", "Estonia", "Finland", "France", "Germany", "Hungary", "Ireland",
        "Italy", "Latvia", "Lithuania", "Mexico", "Netherlands", "Norway", "Peru",
        "Poland", "Portugal", "Romania", "Slovakia", "Spain", "Sweden", "Switzerland",
        "Turkey", "United Kingdom",
    ],
    # Ahlsell countries of operation page + sustainability report
    "Ahlsell": [
        "Sweden", "Norway", "Finland", "Denmark", "Estonia", "Poland", "China",
    ],
    # Uline locations page (DCs)
    "Uline": ["United States", "Canada", "Mexico"],
    # Lyreco direct subsidiaries (not partner markets)
    "Lyreco": [
        "Austria", "Belgium", "Czechia", "Denmark", "France", "Finland", "Germany",
        "Hungary", "Ireland", "Italy", "Luxembourg", "Netherlands", "Norway", "Poland",
        "Portugal", "Slovakia", "Spain", "Sweden", "Switzerland", "United Kingdom",
        "Hong Kong", "South Korea", "Malaysia", "Singapore", "Thailand",
    ],
    # Bunzl documented countries of operation (same evidence set as wearable verify)
    "Bunzl": [
        "United Kingdom", "Ireland", "United States", "Canada", "Mexico", "Puerto Rico",
        "Germany", "France", "Italy", "Spain", "Netherlands", "Belgium", "Denmark",
        "Norway", "Finland", "Switzerland", "Austria", "Czechia", "Hungary", "Romania",
        "Poland", "Israel", "Turkey", "Brazil", "Chile", "Colombia", "Peru", "Uruguay",
        "Australia", "New Zealand", "China", "Singapore",
    ],
    # Veritiv IR / wiki: US + Mexico DCs (Canada sold 2022)
    "Veritiv": ["United States", "Mexico"],
    # Imperial Dade: US + Canada (acquired Veritiv Canada)
    "Imperial Dade": ["United States", "Canada"],
    # Global Industrial: US HQ + Canada office/DCs
    "Global Industrial": ["United States"],
    "Global Industrial Canada": ["Canada"],
    # Associated Bag: US DCs only (Milwaukee + PA/TX/NV)
    "Associated Bag": ["United States"],
    # RAJA Group: claims 19 EU countries — list only evidenced markets (HQ FR + known brands)
    "RAJA": ["France", "United Kingdom"],
    "RAJAPACK UK": ["United Kingdom"],
    # Manutan: LinkedIn workforce countries + FR HQ (not inventing all 17)
    "Manutan": [
        "France", "Netherlands", "United Kingdom", "Italy", "Belgium", "Portugal",
        "Spain", "Czechia", "Germany", "Tunisia",
    ],
    "Mondoffice": ["France"],  # Survilliers HQ (do not invent IT footprint without offices list)
    "Viking Direct": ["Ireland"],
    "JPG": ["France"],
    # Spicers AU vs CA are separate rows
    "Spicers": ["Australia"],
    "Spicers Paper": ["Canada"],
    "Ball & Doggett": ["Australia"],
    "Officeworks": ["Australia"],
    "Blackwoods": ["Australia"],
    "Winc": ["Australia"],
    "Askul": ["Japan"],
    "MonotaRO": ["Japan"],
    "Staples": ["United States"],
    "Quill": ["United States"],
    "The Home Depot": ["United States"],
    "Fastenal": ["United States", "Canada", "Mexico"],  # Fastenal public footprint NA
    "McMaster-Carr": ["United States"],
    "HD Supply": ["United States"],
    "BradyPLUS": ["United States"],
    "NNZ": ["Netherlands"],  # HQ Groningen; do not invent full subsidiary list without page
    "NNZ GmbH": ["Germany"],
    "Brangs + Heinrich": ["Germany", "Netherlands", "Belgium"],  # DE HQ + NL/BE subsidiaries (company page)
    "Verpacken Online": ["Germany"],
    "Hoffmann Verpackung": ["Germany", "Switzerland"],  # DE + CH locations page
    "Papier-Mettler": ["Germany"],
    "Inapa": ["Portugal", "Germany", "France", "Spain", "Belgium", "Luxembourg", "Turkey", "Angola"],
    "Papier Union": ["Germany"],  # now Inapa Deutschland — DE
    "Syspex": ["Singapore", "Malaysia"],  # SG HQ + MY facilities (company materials)
    "Seow Khim Polythene": ["Singapore"],
    "Tantraco": ["Singapore"],
    "PakFactory": ["Canada"],  # Markham office; manufacturing partners overseas are not PakFactory-owned plants
    "Gertex Solutions": ["Canada"],
    "Impak Packaging Systems": ["Canada"],
    "Atlantic Packaging Products": ["Canada"],
    "Richards Packaging": ["Canada"],
    "Crawford Packaging": ["Canada"],
    "Polyair": ["Canada", "United States"],  # Polyair Canada + IPG US parent footprint via itape.com
    "Abco Kovex": ["Ireland"],  # Swords HQ (fix prior wrong UK-only geo)
    "McLernon Packaging": ["Ireland"],
    "The Packaging Centre": ["Ireland", "United Kingdom"],  # Dublin + Belfast/Saintfield offices
    "FlexPRO": ["Bulgaria"],
    "Flexi-Films": ["United Kingdom"],
    "Empaques del Norte": ["Mexico"],
    "Envases y Componentes Saro": ["Mexico"],
    "Allied Distributors": ["South Africa"],
    "Moglix": ["India"],
    "IndustryBuying": ["India"],
    "India Pack": ["India"],
    "Omni Distributors": ["India"],
    "Macfarlane Packaging": ["United Kingdom"],
    "National Flexible": ["United Kingdom"],
    "SupplyOne": ["United States"],
    "Packaging Corporation of America": ["United States"],
    "Shorr Packaging": ["United States"],
    "WebstaurantStore": ["United States"],
    "Aramsco": ["United States"],
    "Zoro Tools": ["United States"],
    "W.B. Mason": ["United States"],
    "Hisco": ["United States"],
    "Kaufman Container": ["United States"],
    "O.Berk": ["United States"],
    "SKS Bottle & Packaging": ["United States"],
    "Unicorr Packaging": ["United States"],
    "Uboxes": ["United States"],
    "The Bottle Crew": ["United States"],
    "Ratioform": ["Germany"],
    "Papyrus": ["Sweden"],  # historic Papyrus SE; leave SE only (Inapa merger — do not invent)
    "Igepa Group": ["Germany"],
    "Cospak": ["Australia"],
    "Signet": ["Australia"],
    "MaxPak": ["Australia"],
    "Ozpack": ["Australia"],
    "Pluspak": ["Australia"],
    "EcoPack": ["Australia"],
    "Gerdmans": ["Sweden"],
    "GWP Group": ["United Kingdom"],
    "Network Packaging": ["United Kingdom"],
    "Karton.eu": ["Germany"],
    "Poly Pak": ["Canada"],
    "Super Poly": ["Canada"],
    "Unipac Packaging Products": ["Canada"],
    "Polythene UK": ["United Kingdom"],
    "Swiftpak": ["United Kingdom"],
    "Transpack": ["United Kingdom"],
    "Transpak": ["Germany"],
    "UK Packaging Supplies": ["United Kingdom"],
    "Packaging2Buy": ["United Kingdom"],  # Bunzl UK brand / Stevenage
    "Packaging Express": ["United Kingdom"],
    "Seal Packaging": ["United Kingdom"],
    "GM Packaging": ["United Kingdom"],
    "Industrial Packaging": ["United Kingdom"],
    "Lesters Packaging": ["United Kingdom"],
    "Multiplastics": ["United Kingdom"],
    "Complete Packaging Solutions": ["United Kingdom"],
    "Diamond Packaging": ["United Kingdom"],
    "Able Packaging": ["United Kingdom"],
    "Pack Supplies": ["United Kingdom"],
    "Packaging World": ["United Kingdom"],
    "Furodet Emballages": ["France"],
    "Kalamazoo": ["Spain"],
    "ESWE Versandpack": ["Germany"],
    "Packaging House": ["New Zealand"],
    "Packaging Products": ["New Zealand"],
    "Premier Packaging": ["New Zealand"],
    "Great Pacific Packaging": ["United States"],
    "Innovative Packaging Group": ["United States"],
    "Preferred Plastics & Packaging": ["United States"],
    "Plastic Bag Source": ["United States"],
    "Packaging Partners": ["United States"],
    "Packaging Price": ["United States"],
    "Stamar Packaging": ["United States"],
    "Schermerhorn Bros.": ["United States"],
    "R.J. Schinner Co.": ["United States"],
    "The Boxery": ["United States"],
    "Imperial Supplies": ["United States"],
    "Office Depot": ["Mexico"],  # row marked Mexico — do not invent US parent footprint for this row
}

# Commercial reach (may equal GEO or be wider when partner/export is evidenced).
VERIFIED_OP: dict[str, list[str]] = {
    "Antalis": VERIFIED_GEO["Antalis"],
    "Ahlsell": [
        "Sweden", "Norway", "Finland", "Denmark", "Estonia", "Poland",
    ],  # China is sourcing office — not primary sell market
    "Uline": ["United States", "Canada", "Mexico"],
    # Lyreco: direct + partner commercial markets (company footprint page)
    "Lyreco": [
        "Austria", "Belgium", "Czechia", "Denmark", "France", "Finland", "Germany",
        "Hungary", "Ireland", "Italy", "Luxembourg", "Netherlands", "Norway", "Poland",
        "Portugal", "Slovakia", "Spain", "Sweden", "Switzerland", "United Kingdom",
        "Hong Kong", "South Korea", "Malaysia", "Singapore", "Thailand",
        "United States", "Canada", "Mexico", "Brazil", "Argentina",
        "Latvia", "Lithuania", "Estonia", "Romania", "Turkey",
        "Taiwan", "China", "India", "Australia", "New Zealand",
    ],
    "Bunzl": VERIFIED_GEO["Bunzl"],
    "Veritiv": ["United States", "Mexico"],
    "Imperial Dade": ["United States", "Canada"],
    "Global Industrial": ["United States"],
    "Global Industrial Canada": ["Canada"],
    "Fastenal": ["United States", "Canada", "Mexico"],
    "RAJA": ["France", "United Kingdom"],  # do not invent remaining 17 of claimed 19
    "RAJAPACK UK": ["United Kingdom"],
    "Manutan": VERIFIED_GEO["Manutan"],  # export to AF/ME/Asia claimed — countries not listed → omit
    "The Packaging Centre": ["Ireland", "United Kingdom"],
    "Hoffmann Verpackung": ["Germany", "Switzerland"],
    "Brangs + Heinrich": ["Germany", "Netherlands", "Belgium"],
    "Inapa": VERIFIED_GEO["Inapa"],
    "Syspex": ["Singapore", "Malaysia"],
    "Polyair": ["Canada", "United States"],
    "Abco Kovex": ["Ireland"],
    "Mondoffice": ["France"],
    "Packaging2Buy": ["United Kingdom"],
}


def canon(token: str) -> str | None:
    t = (token or "").strip()
    if not t:
        return None
    low = t.lower().rstrip(".")
    alias = {
        "usa": "United States",
        "us": "United States",
        "u.s.": "United States",
        "u.s.a.": "United States",
        "uk": "United Kingdom",
        "u.k.": "United Kingdom",
        "great britain": "United Kingdom",
        "england": "United Kingdom",
        "uae": "United Arab Emirates",
        "korea": "South Korea",
        "republic of korea": "South Korea",
        "czech republic": "Czechia",
        "czechia": "Czechia",
        "holland": "Netherlands",
        "the netherlands": "Netherlands",
        "deutschland": "Germany",
        "türkiye": "Turkey",
        "turkiye": "Turkey",
        "hong kong sar": "Hong Kong",
        "hong-kong": "Hong Kong",
        "viet nam": "Vietnam",
        "nz": "New Zealand",
        "au": "Australia",
        "ca": "Canada",
        "de": "Germany",
        "fr": "France",
        "ie": "Ireland",
        "sg": "Singapore",
        "in": "India",
        "mx": "Mexico",
        "za": "South Africa",
        "pt": "Portugal",
        "es": "Spain",
        "it": "Italy",
        "nl": "Netherlands",
        "ch": "Switzerland",
        "se": "Sweden",
        "no": "Norway",
        "dk": "Denmark",
        "fi": "Finland",
        "pl": "Poland",
        "at": "Austria",
        "be": "Belgium",
        "lu": "Luxembourg",
        "jp": "Japan",
        "cn": "China",
        "my": "Malaysia",
        "th": "Thailand",
        "bg": "Bulgaria",
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
        "bosnia and herzegovina": "Bosnia and Herzegovina",
    }
    if low in specials:
        return specials[low]
    if t in COUNTRY_REGION:
        return t
    # Title-case simple names
    titled = t.title() if t.islower() or t.isupper() else t
    if titled in COUNTRY_REGION:
        return titled
    return None


US_STATE_RE = re.compile(
    r"\b(AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|HI|ID|IL|IN|IA|KS|KY|LA|ME|MD|MA|MI|MN|MS|"
    r"MO|MT|NE|NV|NH|NJ|NM|NY|NC|ND|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VT|VA|WA|WV|WI|WY|"
    r"DC)\b",
    re.I,
)
CA_PROV_RE = re.compile(
    r"\b(Ontario|Quebec|Alberta|British Columbia|Manitoba|Saskatchewan|"
    r"Nova Scotia|New Brunswick|Newfoundland|Prince Edward Island|"
    r"ON|QC|AB|BC|MB|SK|NS|NB|NL|PE)\b",
    re.I,
)


def country_from_hq(hq: str) -> str | None:
    s = (hq or "").strip()
    if not s or "not publicly" in s.lower():
        return None
    # Strip parenthetical notes
    s_clean = re.sub(r"\([^)]*\)", " ", s)
    low = s_clean.lower()
    if "canada" in low:
        return "Canada"
    if any(x in low for x in ("united states", "usa", "u.s.")):
        return "United States"
    if any(x in low for x in ("united kingdom", "england", "scotland", "wales")):
        return "United Kingdom"
    # Prefer last comma segment (Country)
    if "," in s_clean:
        parts = [p.strip() for p in s_clean.split(",") if p.strip()]
        for part in reversed(parts):
            part0 = re.split(r"[;/]", part)[0].strip()
            c = canon(part0)
            if c:
                return c
    c = canon(s_clean)
    if c:
        return c
    if US_STATE_RE.search(s_clean):
        return "United States"
    if CA_PROV_RE.search(s_clean):
        return "Canada"
    return None


def uniq(countries: list[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for c in countries:
        cc = canon(c) or c
        if cc not in COUNTRY_REGION:
            continue
        if cc not in seen:
            seen.add(cc)
            out.append(cc)
    return out


def format_geo(countries: list[str]) -> str:
    cs = uniq(countries)
    if not cs:
        return "Not publicly disclosed"
    buckets: dict[str, list[str]] = {r: [] for r in REGION_ORDER}
    for c in cs:
        buckets[COUNTRY_REGION[c]].append(c)
    parts = []
    for r in REGION_ORDER:
        if buckets[r]:
            parts.append(f"{r}({', '.join(buckets[r])})")
    return "; ".join(parts)


def format_op(countries: list[str]) -> str:
    cs = uniq(countries)
    if not cs:
        return "Not publicly disclosed"
    return "; ".join(cs)


def main() -> None:
    wb = load_workbook(FP)
    ws = wb.active
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    assert headers[4] == "Continent / Geography"
    assert headers[5] == "Operational Presence"

    audit = {"updated_at": datetime.now().isoformat(timespec="seconds"), "rows": []}
    n_geo = n_op = n_hq_only = n_verified = 0

    for row in ws.iter_rows(min_row=2):
        name = str(row[0].value or "").strip()
        website = str(row[1].value or "").strip()
        hq = str(row[3].value or "").strip()
        old_geo = str(row[4].value or "").strip()
        old_op = str(row[5].value or "").strip()
        if not name:
            continue

        hq_country = country_from_hq(hq)
        source = "hq_baseline"

        if name in VERIFIED_GEO:
            geo_list = list(VERIFIED_GEO[name])
            source = "verified_geo"
            n_verified += 1
        elif hq_country:
            geo_list = [hq_country]
            source = "hq_baseline"
            n_hq_only += 1
        else:
            geo_list = []
            source = "none"

        # Ensure HQ country included when known
        if hq_country and hq_country not in geo_list and hq_country in COUNTRY_REGION:
            geo_list = [hq_country] + geo_list

        if name in VERIFIED_OP:
            op_list = list(VERIFIED_OP[name])
        elif geo_list:
            # Domestic / single-country distributors: commercial reach at least = physical country
            op_list = list(geo_list)
        else:
            op_list = []

        new_geo = format_geo(geo_list)
        new_op = format_op(op_list)

        if new_geo != old_geo:
            row[4].value = new_geo
            n_geo += 1
        else:
            row[4].value = new_geo  # normalize format even if same countries

        if new_op != old_op or not old_op:
            row[5].value = new_op
            n_op += 1
        else:
            row[5].value = new_op

        audit["rows"].append(
            {
                "company": name,
                "website": website,
                "hq": hq,
                "hq_country": hq_country,
                "source": source,
                "old_geo": old_geo,
                "new_geo": new_geo,
                "old_op": old_op,
                "new_op": new_op,
            }
        )

    wb.save(FP)
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")

    # Stats
    weak_geo = weak_op = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        g = str(row[4] or "")
        o = str(row[5] or "")
        if "not publicly" in g.lower() or not g.strip():
            weak_geo += 1
        if "not publicly" in o.lower() or not o.strip():
            weak_op += 1

    print(f"saved {FP}")
    print(f"geo_cells_written={n_geo} op_cells_written={n_op}")
    print(f"verified_overlays={n_verified} hq_baseline_only={n_hq_only}")
    print(f"still_weak_geo={weak_geo} still_weak_op={weak_op}")
    print(f"audit {AUDIT}")


if __name__ == "__main__":
    main()
