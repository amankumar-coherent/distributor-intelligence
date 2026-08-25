# -*- coding: utf-8 -*-
"""
Reformat Continent / Geography and Operational Presence to country-only style:

Continent / Geography: AF(South Africa, Namibia, Botswana, Mauritius)
Operational Presence:  South Africa; Namibia; Botswana; Mauritius

Does NOT change Company, Website, Founded, Headquarters.
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
AUDIT = ROOT / "_wearable_geo_op_country_format_audit.json"

# Country -> region code (match prior landscape style)
COUNTRY_REGION: dict[str, str] = {
    # NA
    "United States": "NA",
    "Canada": "NA",
    "Mexico": "NA",
    # LATAM / SA
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
    "Cuba": "LATAM",
    # EU
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
    "Czech Republic": "EU",
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
    "Ukraine": "EU",
    "Russia": "EU",
    "Turkey": "EU",  # often ME; Alliance lists as Europe network — keep EU for Alliance footprint
    # AF
    "South Africa": "AF",
    "Namibia": "AF",
    "Botswana": "AF",
    "Mauritius": "AF",
    "Nigeria": "AF",
    "Kenya": "AF",
    "Ghana": "AF",
    "Egypt": "AF",
    "Morocco": "AF",
    "Algeria": "AF",
    "Tunisia": "AF",
    "Ethiopia": "AF",
    "Tanzania": "AF",
    "Uganda": "AF",
    # ME
    "Saudi Arabia": "ME",
    "United Arab Emirates": "ME",
    "Qatar": "ME",
    "Kuwait": "ME",
    "Bahrain": "ME",
    "Oman": "ME",
    "Jordan": "ME",
    "Iraq": "ME",
    "Israel": "ME",
    "Lebanon": "ME",
    # APAC
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
    "Pakistan": "APAC",
    "Bangladesh": "APAC",
    "Sri Lanka": "APAC",
    # OC
    "Australia": "OC",
    "New Zealand": "OC",
}

# Preferred region order when multiple groups
REGION_ORDER = ["NA", "LATAM", "EU", "AF", "ME", "APAC", "OC", "AS"]

CANON = {
    "usa": "United States",
    "us": "United States",
    "u.s.": "United States",
    "uk": "United Kingdom",
    "u.k.": "United Kingdom",
    "uae": "United Arab Emirates",
    "korea": "South Korea",
    "republic of korea": "South Korea",
    "hong kong sar": "Hong Kong",
    "mainland china": "China",
    "czech republic": "Czechia",
    "czechia": "Czechia",
    "türkiye": "Turkey",
    "turkiye": "Turkey",
    "the netherlands": "Netherlands",
    "viet nam": "Vietnam",
}


def canon_country(token: str) -> str | None:
    t = (token or "").strip()
    if not t:
        return None
    # strip trailing annotations like "(HQ)", "(stores)", "(distribution network)"
    t = re.sub(r"\s*\([^)]*\)\s*", " ", t).strip()
    t = re.sub(r"\s+", " ", t)
    low = t.lower().strip(" .,")
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
        "africa",
        "north america",
        "multiple distribution / service centres",
        "distribution network",
        "operations",
        "market ops",
        "pharmacy stores",
        "stores",
        "offices / warehouses — 6+ au locations cited",
        "warehouses / distribution — both islands",
        "clinical reach / office presence",
        "via compass health brands",
        "global capability center office",
        "west malaysia address",
        "east malaysia address",
        "registrant entity / marina bay registered office",
        "walmart canada operations",
        "amazon.ca operations / facilities",
        "mckesson canada",
        "both islands",
    }:
        return None
    if low in CANON:
        return CANON[low]
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
    if low in specials:
        return specials[low]
    # If token already a known country
    if t in COUNTRY_REGION:
        return t
    # Extract country from "City, ST, United States" / "City, Country"
    if "," in t:
        parts = [p.strip() for p in t.split(",") if p.strip()]
        # try last part
        last = parts[-1]
        c = canon_country(last)
        if c:
            return c
        # try "United States" style two-word at end already handled
        if len(parts) >= 2:
            maybe = ", ".join(parts[-2:])
            if maybe in COUNTRY_REGION:
                return maybe
    # Known country as substring at end
    for cname in sorted(COUNTRY_REGION.keys(), key=len, reverse=True):
        if low == cname.lower() or low.endswith(" " + cname.lower()) or low.endswith("," + cname.lower()):
            return cname
        if cname.lower() in low and low.endswith(cname.lower()):
            return cname
    return None


def extract_countries(raw: str) -> list[str]:
    s = (raw or "").strip()
    if not s:
        return []
    # If already CODE(a, b) form — pull inside
    countries: list[str] = []
    for m in re.finditer(r"[A-Z]{2,5}\(([^)]*)\)", s):
        inner = m.group(1)
        for part in re.split(r"\s*,\s*", inner):
            c = canon_country(part)
            if c:
                countries.append(c)
    if countries:
        return dedupe(countries)

    # Split on ; first (OP style or multi geo)
    chunks = re.split(r"\s*;\s*", s)
    for ch in chunks:
        ch = ch.strip()
        if not ch:
            continue
        # CODE(inner) already handled; plain country or city,country
        c = canon_country(ch)
        if c:
            countries.append(c)
            continue
        # comma list of countries without code
        if "," in ch and not re.search(r"\b[A-Z]{2}\b", ch.split(",")[0]):
            # could be City, ST, United States OR Country, Country
            c2 = canon_country(ch)
            if c2:
                countries.append(c2)
            else:
                for part in re.split(r"\s*,\s*", ch):
                    c3 = canon_country(part)
                    if c3:
                        countries.append(c3)
    return dedupe(countries)


def dedupe(seq: list[str]) -> list[str]:
    seen = set()
    out = []
    for x in seq:
        if x and x not in seen:
            seen.add(x)
            out.append(x)
    return out


def format_geo(countries: list[str]) -> str:
    """AF(South Africa, Namibia) or APAC(a, b); ME(c, d)"""
    groups: dict[str, list[str]] = {}
    for c in countries:
        # Turkey for Watsons ME retail — if in ME markets list use ME; default map may say EU
        region = COUNTRY_REGION.get(c, "APAC")
        # Prefer ME for GCC + Turkey when mixed with Watsons-style — keep map as-is
        groups.setdefault(region, [])
        if c not in groups[region]:
            groups[region].append(c)
    parts = []
    for region in REGION_ORDER:
        if region in groups and groups[region]:
            parts.append(f"{region}({', '.join(groups[region])})")
    for region, clist in groups.items():
        if region not in REGION_ORDER:
            parts.append(f"{region}({', '.join(clist)})")
    return "; ".join(parts)


def format_op(countries: list[str]) -> str:
    return "; ".join(countries)


def apply_format(ws) -> None:
    headers = [c.value for c in ws[1]]
    widths = {
        "Company": 34,
        "Website": 36,
        "Founded": 10,
        "Headquarters": 28,
        "Continent / Geography": 48,
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
    ws.row_dimensions[1].height = 32
    for i in range(1, len(headers) + 1):
        cell = ws.cell(1, i)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_align
        cell.border = thin
    geo_i = headers.index("Continent / Geography") + 1
    op_i = headers.index("Operational Presence") + 1
    for r in range(2, ws.max_row + 1):
        longest = max(
            len(str(ws.cell(r, geo_i).value or "")),
            len(str(ws.cell(r, op_i).value or "")),
            20,
        )
        lines = max(2, min(6, (longest // 45) + 1))
        ws.row_dimensions[r].height = 15 * lines
        for i in range(1, len(headers) + 1):
            cell = ws.cell(r, i)
            cell.font = body_font
            cell.alignment = wrap
            cell.border = thin
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"


def main() -> None:
    # Special-case Turkey region for GCC/ME retail brands
    ME_FORCE = {
        "Watsons",
        "Aster Pharmacy",
        "BinSina Pharmacy",
        "Al Nahdi Medical Company",
        "Al-Dawaa Medical Services",
    }

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = ROOT / f"global_wearable_medical_devices_market_global_FINAL_BEFORE_COUNTRY_FMT_{stamp}.xlsx"
    shutil.copy2(XLSX, bak)

    wb = load_workbook(XLSX)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    ic = {h: i + 1 for i, h in enumerate(headers) if h}

    samples = []
    empty = []
    for r in range(2, ws.max_row + 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        hq = str(ws.cell(r, ic["Headquarters"]).value or "").strip()
        old_geo = str(ws.cell(r, ic["Continent / Geography"]).value or "").strip()
        old_op = str(ws.cell(r, ic["Operational Presence"]).value or "").strip()

        geo_countries = extract_countries(old_geo)
        op_countries = extract_countries(old_op)
        hq_country = canon_country(hq) if hq else None

        if not geo_countries and hq_country:
            geo_countries = [hq_country]
        if not op_countries and hq_country:
            op_countries = [hq_country]
        if hq_country and hq_country not in geo_countries:
            geo_countries = [hq_country] + geo_countries
        if hq_country and hq_country not in op_countries:
            # keep OP as commercial list; ensure HQ country present if list empty already handled
            pass

        # Force Turkey into ME for Watsons-like brands when formatting GEO groups
        if co in ME_FORCE:
            # temporarily remap Turkey to ME for this row's geo formatting
            pass

        # Build GEO with optional Turkey->ME for specific brands
        def format_geo_row(countries: list[str], company: str) -> str:
            groups: dict[str, list[str]] = {}
            for c in countries:
                if company in ME_FORCE and c == "Turkey":
                    region = "ME"
                elif c in {
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
                else:
                    region = COUNTRY_REGION.get(c, "APAC")
                groups.setdefault(region, [])
                if c not in groups[region]:
                    groups[region].append(c)
            parts = []
            for region in REGION_ORDER:
                if region in groups and groups[region]:
                    parts.append(f"{region}({', '.join(groups[region])})")
            for region, clist in groups.items():
                if region not in REGION_ORDER:
                    parts.append(f"{region}({', '.join(clist)})")
            return "; ".join(parts)

        new_geo = format_geo_row(geo_countries, co)
        new_op = format_op(op_countries)

        # Do not touch Company/Website/Founded/Headquarters
        ws.cell(r, ic["Continent / Geography"]).value = new_geo
        ws.cell(r, ic["Operational Presence"]).value = new_op

        if not new_geo or not new_op:
            empty.append(co)
        if co in {
            "Al-Bahar Medical",
            "Adcock Ingram",
            "Alliance Healthcare",
            "Transmedic",
            "Watsons",
            "AAH Pharmaceuticals",
            "Giant Eagle",
            "Device Technologies",
        }:
            samples.append(
                {
                    "company": co,
                    "geo": new_geo,
                    "op": new_op,
                    "old_geo": old_geo[:120],
                    "old_op": old_op[:120],
                }
            )

    apply_format(ws)
    wb.save(XLSX)

    audit = {
        "backup": bak.name,
        "format": {
            "Continent / Geography": "CODE(Country1, Country2)",
            "Operational Presence": "Country1; Country2",
        },
        "rows": ws.max_row - 1,
        "empty": empty,
        "samples": samples,
    }
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print("backup", bak.name)
    print("rows", ws.max_row - 1, "empty", len(empty))
    for s in samples:
        print("---", s["company"])
        print(" GEO", s["geo"])
        print(" OP ", s["op"])


if __name__ == "__main__":
    main()
