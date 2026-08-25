#!/usr/bin/env python3
"""Restore Continent / Geography to CODE(Country) format for all avocado rows.

Example: NA(United States); ME(United Arab Emirates, Qatar); AF(Egypt)

Uses Operational Presence countries first, then HQ country, then geo snapshot.
Never invents countries.
"""
from __future__ import annotations

import json
import re
import unicodedata
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment

ROOT = Path(r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global")
XLSX = ROOT / "global_avocado_oil_market.xlsx"
SNAPSHOT = ROOT / "_avocado_geo_audit_snapshot.json"
AUDIT = ROOT / "_avocado_restore_continent_countries_audit.json"

# Display-normalized country -> region code (file convention)
COUNTRY_REGION: dict[str, str] = {
    "united states": "NA",
    "usa": "NA",
    "u.s.": "NA",
    "u.s.a.": "NA",
    "united states of america": "NA",
    "puerto rico": "NA",
    "canada": "NA",
    "mexico": "NA",
    "costa rica": "LATAM",
    "guatemala": "LATAM",
    "honduras": "LATAM",
    "el salvador": "LATAM",
    "nicaragua": "LATAM",
    "panama": "LATAM",
    "colombia": "LATAM",
    "dominican republic": "LATAM",
    "chile": "LATAM",
    "brazil": "LATAM",
    "argentina": "LATAM",
    "peru": "LATAM",
    "ecuador": "LATAM",
    "uruguay": "LATAM",
    "venezuela": "LATAM",
    "curaçao": "LATAM",
    "curacao": "LATAM",
    "jamaica": "LATAM",
    "barbados": "LATAM",
    "aruba": "LATAM",
    "trinidad and tobago": "LATAM",
    "united states virgin islands": "LATAM",
    "u.s. virgin islands": "LATAM",
    "united kingdom": "EU",
    "england": "EU",
    "scotland": "EU",
    "wales": "EU",
    "ireland": "EU",
    "germany": "EU",
    "france": "EU",
    "netherlands": "EU",
    "belgium": "EU",
    "luxembourg": "EU",
    "spain": "EU",
    "italy": "EU",
    "portugal": "EU",
    "switzerland": "EU",
    "austria": "EU",
    "sweden": "EU",
    "norway": "EU",
    "denmark": "EU",
    "finland": "EU",
    "iceland": "EU",
    "poland": "EU",
    "czech republic": "EU",
    "czechia": "EU",
    "slovakia": "EU",
    "hungary": "EU",
    "romania": "EU",
    "greece": "EU",
    "bulgaria": "EU",
    "croatia": "EU",
    "serbia": "EU",
    "slovenia": "EU",
    "estonia": "EU",
    "latvia": "EU",
    "lithuania": "EU",
    "ukraine": "EU",
    "belarus": "EU",
    "moldova": "EU",
    "turkey": "EU",
    "georgia": "EU",
    "russia": "EU",
    "russian federation": "EU",
    "kazakhstan": "EU",
    "uzbekistan": "EU",
    "united arab emirates": "ME",
    "uae": "ME",
    "saudi arabia": "ME",
    "qatar": "ME",
    "oman": "ME",
    "bahrain": "ME",
    "kuwait": "ME",
    "jordan": "ME",
    "israel": "ME",
    "lebanon": "ME",
    "iraq": "ME",
    "pakistan": "APAC",
    "india": "APAC",
    "china": "APAC",
    "japan": "APAC",
    "korea": "APAC",
    "south korea": "APAC",
    "republic of korea": "APAC",
    "taiwan": "APAC",
    "hong kong": "APAC",
    "singapore": "APAC",
    "malaysia": "APAC",
    "indonesia": "APAC",
    "thailand": "APAC",
    "vietnam": "APAC",
    "philippines": "APAC",
    "australia": "APAC",
    "new zealand": "APAC",
    "south africa": "AF",
    "tunisia": "AF",
    "zimbabwe": "AF",
    "botswana": "AF",
    "namibia": "AF",
    "zambia": "AF",
    "malawi": "AF",
    "mozambique": "AF",
    "lesotho": "AF",
    "eswatini": "AF",
    "egypt": "AF",
    "ghana": "AF",
    "kenya": "AF",
    "morocco": "AF",
    "nigeria": "AF",
}

# Canonical display names
CANON = {
    "usa": "United States",
    "u.s.": "United States",
    "u.s.a.": "United States",
    "united states of america": "United States",
    "uae": "United Arab Emirates",
    "korea": "South Korea",
    "south korea": "South Korea",
    "republic of korea": "South Korea",
    "czechia": "Czech Republic",
    "russian federation": "Russian Federation",
    "england": "United Kingdom",
    "scotland": "United Kingdom",
    "wales": "United Kingdom",
    "curacao": "Curaçao",
}

ORDER = ["NA", "EU", "ME", "APAC", "LATAM", "AF"]


def norm_key(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower().replace("&", " and ").replace("*", " ")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def canon_country(raw: str) -> tuple[str, str] | None:
    """Return (display_name, region_code) or None."""
    t = (raw or "").strip().strip(".")
    if not t:
        return None
    low = t.lower()
    # strip trailing region noise like "Western Australia, Australia" handled by split
    if low in COUNTRY_REGION:
        disp = CANON.get(low, t.title() if low in CANON or low in {"usa", "uae"} else t)
        if low in CANON:
            disp = CANON[low]
        elif low == "united states":
            disp = "United States"
        elif low == "united arab emirates":
            disp = "United Arab Emirates"
        elif low == "united kingdom":
            disp = "United Kingdom"
        elif low == "new zealand":
            disp = "New Zealand"
        elif low == "south africa":
            disp = "South Africa"
        elif low == "saudi arabia":
            disp = "Saudi Arabia"
        elif low == "czech republic":
            disp = "Czech Republic"
        elif low == "dominican republic":
            disp = "Dominican Republic"
        elif low == "el salvador":
            disp = "El Salvador"
        elif low == "costa rica":
            disp = "Costa Rica"
        elif low == "hong kong":
            disp = "Hong Kong"
        elif low == "south korea":
            disp = "South Korea"
        elif low == "russian federation":
            disp = "Russian Federation"
        elif low == "trinidad and tobago":
            disp = "Trinidad and Tobago"
        elif low == "united states virgin islands":
            disp = "United States Virgin Islands"
        else:
            disp = t if t[:1].isupper() else t.title()
            # prefer Title Case of known key
            disp = " ".join(w.capitalize() if w not in {"and"} else w for w in low.split())
            specials = {
                "united states": "United States",
                "united kingdom": "United Kingdom",
                "united arab emirates": "United Arab Emirates",
                "new zealand": "New Zealand",
                "south africa": "South Africa",
                "saudi arabia": "Saudi Arabia",
                "czech republic": "Czech Republic",
                "dominican republic": "Dominican Republic",
                "el salvador": "El Salvador",
                "costa rica": "Costa Rica",
                "hong kong": "Hong Kong",
                "south korea": "South Korea",
                "papua new guinea": "Papua New Guinea",
            }
            disp = specials.get(low, disp)
        return disp, COUNTRY_REGION[low]

    # try aliases via CANON first
    if low in CANON:
        return canon_country(CANON[low])

    return None


def countries_from_text(text: str) -> list[str]:
    """Extract country tokens from presence/HQ-like text."""
    if not text or str(text).strip().lower() in {"", "n/a", "na", "none", "global", "worldwide"}:
        return []
    # presence is usually "A; B; C"
    parts = [p.strip() for p in re.split(r"[;|/]", text) if p.strip()]
    out: list[str] = []
    for p in parts:
        # if comma list inside one segment, split further only when all look like countries
        if "," in p and ";" not in text:
            # HQ style: "City, State, Country" or "City, Country"
            bits = [b.strip() for b in p.split(",") if b.strip()]
            # try last bit, then last two combined, then each bit
            candidates = []
            if bits:
                candidates.append(bits[-1])
            if len(bits) >= 2:
                candidates.append(bits[-1])
                candidates.append(", ".join(bits[-2:]))
            for b in bits:
                candidates.append(b)
            hit = None
            for c in candidates:
                if canon_country(c):
                    hit = canon_country(c)
                    break
            if hit:
                out.append(hit[0])
            continue

        # plain country name
        hit = canon_country(p)
        if hit:
            out.append(hit[0])
            continue
        # "Western Australia, Australia" already handled; try last comma piece
        if "," in p:
            last = p.split(",")[-1].strip()
            hit = canon_country(last)
            if hit:
                out.append(hit[0])
    # dedupe preserve order
    seen = set()
    uniq = []
    for c in out:
        if c.lower() in seen:
            continue
        seen.add(c.lower())
        uniq.append(c)
    return uniq


def format_continent(countries: list[str]) -> str:
    buckets: dict[str, list[str]] = defaultdict(list)
    for c in countries:
        hit = canon_country(c)
        if not hit:
            # already canonical display — look up via lower
            hit = canon_country(c.lower()) if c else None
        if not hit:
            continue
        disp, code = hit
        if disp not in buckets[code]:
            buckets[code].append(disp)
    chunks = []
    for code in ORDER:
        if code in buckets:
            chunks.append(f"{code}({', '.join(buckets[code])})")
    for code, lst in buckets.items():
        if code not in ORDER:
            chunks.append(f"{code}({', '.join(lst)})")
    return "; ".join(chunks)


def parse_coded_continent(text: str) -> list[str]:
    """Extract country names from NA(United States, Canada); EU(Germany) strings."""
    if not text or "(" not in text:
        return []
    countries: list[str] = []
    for m in re.finditer(r"\(([^)]+)\)", text):
        for part in m.group(1).split(","):
            p = part.strip()
            if p:
                countries.append(p)
    return countries


def main() -> None:
    snap_by = {}
    if SNAPSHOT.exists():
        for row in json.loads(SNAPSHOT.read_text(encoding="utf-8")):
            snap_by[norm_key(row.get("company") or "")] = row

    wb = openpyxl.load_workbook(XLSX)
    ws = wb.active
    hdr = [c.value for c in ws[2]]
    ic = {str(h): i + 1 for i, h in enumerate(hdr) if h}

    updated = []
    unresolved = []

    for r in range(3, ws.max_row + 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        if not co:
            continue
        presence = str(ws.cell(r, ic["Operational Presence"]).value or "").strip()
        hq = str(ws.cell(r, ic["Headquarters"]).value or "").strip()
        old = str(ws.cell(r, ic["Continent / Geography"]).value or "").strip()

        countries = countries_from_text(presence)
        if not countries:
            countries = countries_from_text(hq)

        # enrich from snapshot if coded continent has more countries
        snap = snap_by.get(norm_key(co))
        if snap:
            snap_countries = parse_coded_continent(str(snap.get("continent") or ""))
            # use snapshot when presence-derived is empty or much thinner
            if not countries:
                countries = snap_countries
            elif len(snap_countries) > len(countries) + 2:
                # prefer richer verified snapshot for multi-market distributors
                countries = snap_countries

        # if still only bare continent word in old field with countries elsewhere already tried
        new_val = format_continent(countries)
        if not new_val:
            unresolved.append({"company": co, "hq": hq, "presence": presence, "old": old})
            continue

        ws.cell(r, ic["Continent / Geography"]).value = new_val
        ws.cell(r, ic["Continent / Geography"]).alignment = Alignment(
            wrap_text=True, vertical="top"
        )
        updated.append({"company": co, "old": old, "new": new_val})

    try:
        wb.save(XLSX)
        saved = str(XLSX)
    except PermissionError:
        alt = ROOT / "global_avocado_oil_market_OFFICE_SLIM.xlsx"
        wb.save(alt)
        saved = str(alt) + " (main locked; saved here)"
    wb.close()

    audit = {
        "ts": datetime.now(timezone.utc).isoformat(),
        "saved": saved,
        "updated": len(updated),
        "unresolved": unresolved,
        "samples": updated[:25],
    }
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps({k: audit[k] for k in ("saved", "updated", "unresolved") if k != "unresolved"}, indent=2))
    print("unresolved_count", len(unresolved))
    for u in unresolved[:20]:
        print(" UNRES", u["company"], "|", u["presence"][:60], "|", u["hq"][:50])
    print("samples:")
    for s in updated[:12]:
        print(" +", s["company"][:40], "=>", s["new"][:90])


if __name__ == "__main__":
    main()
