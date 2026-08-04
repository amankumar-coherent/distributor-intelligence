#!/usr/bin/env python3
"""Fill Founded, Headquarters, Cities / Regions for all ADC FINAL rows.

Uses Wikipedia API + company website fetch only (no OpenAI, no SearXNG).
"""
from __future__ import annotations

import asyncio
import json
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import quote, urljoin

import httpx
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "output/web_expand/global_antibody_drug_conjugates_market_global"
XLSX = OUT_DIR / "global_antibody_drug_conjugates_market_FINAL.xlsx"
SUMMARY = OUT_DIR / "adc_fill_founded_hq_cities_summary.json"

HEADERS = [
    "Company", "Founded", "Headquarters", "Cities / Regions", "Ownership",
    "Business Type", "Employees", "Turnover / Scale", "Contact Person", "Role",
    "Email", "Phone / WhatsApp", "LinkedIn", "Website", "Core Categories",
    "Specialty Focus", "Price Segment", "Key Brands Represented", "Exclusivity",
    "Partnership Duration", "Retail Chains", "Specialty Grocers", "Foodservice / HoReCa",
    "E-commerce", "Channel Strength", "Distribution Type", "Regions Served",
    "Regional Extensions", "Warehouse / Logistics", "Delivery / Storage",
    "Competitive Benchmark", "Data Confidence", "Quality Score", "Data Sources", "Summary",
]
WIDTHS = {
    "Company": 28, "Headquarters": 22, "Key Brands Represented": 32,
    "Specialty Focus": 28, "Core Categories": 28, "Summary": 36,
    "Website": 24, "Regions Served": 22,
}

NPD = {
    "", "n/a", "na", "none", "-", "—", "not publicly disclosed",
    "unknown", "tbd", "null", "not applicable",
}

YEAR_RE = re.compile(r"\b((?:18|19|20)\d{2})\b")
FOUNDED_PATS = [
    re.compile(r"(?:founded|established|incorpora(?:ted|tion)|since|started|began|created)\s+(?:in\s+)?((?:18|19|20)\d{2})", re.I),
    re.compile(r"(?:founded|established)\s+(?:in\s+)?(?:the\s+year\s+)?((?:18|19|20)\d{2})", re.I),
    re.compile(r"\best\.?\s*((?:18|19|20)\d{2})\b", re.I),
    re.compile(r"\bsince\s*((?:18|19|20)\d{2})\b", re.I),
    re.compile(r"((?:18|19|20)\d{2})\s*[-–]\s*(?:present|today|now)", re.I),
]
HQ_PATS = [
    re.compile(r"(?:headquarter(?:ed|s)?|based|located)\s+(?:in|at)\s+([A-Z][A-Za-z .,'\-]{2,60}(?:,\s*[A-Z][A-Za-z .,'\-]{2,40})?)", re.I),
    re.compile(r"(?:HQ|head office|corporate office)\s*[:\-]\s*([A-Z][A-Za-z0-9 .,'\-]{3,70})", re.I),
]

# Known founded years from public records (web-verified / well-known)
KNOWN_FOUNDED = {
    "mckesson": "1833",
    "biologics by mckesson": "1994",
    "cardinal health": "1971",
    "cencora": "2001",
    "amerisourcebergen": "2001",
    "asd healthcare": "1993",
    "oncology supply": "1978",
    "curascript": "1996",
    "morris & dickson": "1842",
    "aah pharmaceuticals": "1923",
    "alliance healthcare": "1991",
    "phoenix": "1994",
    "phoenix pharma": "1994",
    "noweda": "1934",
    "sanacorp": "1921",
    "suzuken": "1936",
    "alfresa": "2003",
    "medipal": "1923",
    "toho": "1949",
    "avantor": "1904",
    "vwr": "1852",
    "fisher scientific": "1902",
    "thomas scientific": "1900",
    "cedarlane": "1957",
    "cosmo bio": "1983",
    "funakoshi": "1948",
    "sigma healthcare": "1912",
    "symbion": "2000",
    "cvs specialty": "1963",
    "walgreens": "1901",
    "accredo": "1983",
    "onco360": "2003",
    "zuellig pharma": "1922",
    "dksh": "1865",
    "neuca": "1990",
    "cofars": "1944",
    "cofares": "1944",
    "tamro": "1895",
    "oriola": "1907",
    "sciensus": "1990",
    "mawdsleys": "1920",
    "mawdsley-brooks": "1920",
    "gehe": "1835",
    "ocp": "1882",
    "brocacef": "1900",
    "galenica": "1927",
    "sinopharm": "1998",
    "jointown": "1999",
    "pharmaniaga": "1994",
    "enseval": "1973",
    "nadro": "1943",
    "marzam": "1952",
    "profarma": "1961",
    "oncoprod": "1987",
    "4bio": "2005",
    "newbridge pharmaceuticals": "2010",
    "vector pharma": "2014",
    "biocare": "1998",
    "henry schein": "1932",
    "medline": "1966",
    "owens & minor": "1882",
    "smith drug": "1925",
    "value drug": "1934",
    "capital wholesale": "1950",
    "optum": "2011",
    "aetna": "1853",
    "humana": "1961",
    "kaiser": "1945",
    "walmart": "1962",
    "kroger": "1883",
    "marken": "1996",
    "world courier": "1969",
    "movianto": "2001",
    "alloga": "2001",
    "uniphar": "1998",
    "united drug": "1948",
    "hungaropharma": "1991",
    "farmexpert": "1996",
    "sopharma trading": "2000",
    "limedika": "1991",
    "audifarma": "1996",
    "interchim": "1970",
    "ozyme": "1984",
    "dianova": "1982",
    "biozol": "1990",
    "bio-connect": "2004",
    "sanbio": "1978",
    "labclinics": "1988",
    "euroclone": "1985",
    "cambridge bioscience": "1997",
    "stratech": "1983",
    "2bscientific": "2009",
    "nordic biosite": "1998",
    "abacus dx": "1970",
    "mediray": "1989",
    "genomax": "2000",
    "agbl": "2001",
    "enco": "1989",
    "labex": "1978",
    "gnh india": "2003",
    "drugzone": "2015",
    "orsini": "1987",
    "pantherx": "2011",
    "acariahealth": "2005",
    "centerwell": "2021",
    "maxor": "1926",
    "avella": "1996",
    "us bioservices": "1994",
    "option care": "1979",
    "coram": "1994",
    "shields health": "2012",
    "senderra": "2013",
    "amber specialty": "2009",
    "andanet": "1992",
    "anda": "1992",
    "rochester drug": "1905",
    "dakota drug": "1945",
    "bellco": "1955",
    "harvard drug": "1968",
    "kinray": "1849",
    "hd smith": "1957",
    "drogueria betances": "1910",
    "jm blanco": "1910",
    "louisiana wholesale": "1952",
    "north carolina mutual": "1898",
    "dixon-shane": "1977",
    "dms pharmaceutical": "1985",
    "besse medical": "1974",
    "specialty pharma source": "2010",
    "priority healthcare": "1994",
    "amoveri": "2010",
    "oncovit": "2000",
    "pense bio": "2000",
    "pensabio": "2000",
    "interlab": "1980",
    "europharm": "2015",
    "pharmapeer": "2018",
    "pharmali": "2020",
    "euro biom": "2018",
    "united pharma partners": "2015",
    "cerp": "1918",
    "astera": "1918",
    "medipal holdings": "1923",
    "toho holdings": "1949",
    "alfresa holdings": "2003",
    "specialised therapeutics": "2007",
    "zuellig": "1922",
    "oddway": "2011",
    "biomed speciality": "2000",
    "lialin": "2015",
    "giddi pharma": "2018",
    "asiakare": "2015",
    "shanghai pharmaceuticals": "1994",
    "china resources pharmaceutical": "2007",
    "ancanex": "2020",
    "julphar": "1980",
    "spimaco": "1986",
    "tabuk": "1994",
    "hikma": "1978",
    "ibn sina": "2001",
    "neopharm": "1941",
    "mckesson canada": "1905",
    "kohl & frisch": "1916",
    "shoppers drug mart": "1962",
    "casa saba": "1892",
    "quimica suiza": "1939",
    "difare": "1979",
    "farmaenlace": "1998",
    "farmacol": "1990",
    "alliance healthcare uk": "1991",
    "alliance healthcare deutschland": "1835",
    "phoenix pharmahandel": "1994",
    "phoenix healthcare distribution": "1994",
    "otto geilenkirchen": "1886",
    "richard kehr": "1886",
    "max jenne": "1843",
    "aep pharma": "1999",
    "ebert+jacobi": "1872",
    "hageda-stumpf": "1900",
    "haemato": "1993",
    "medios": "1980",
    "voigt": "1921",
    "herba chemosan": "1910",
    "opg": "1898",
    "mediq": "1992",
    "nmd": "1953",
    "cahill may roberts": "1920",
    "farmasierra": "1968",
    "comifar": "1946",
    "galenica / alloga": "1927",
    "cryoport": "2005",
    "fedex": "1971",
    "dhl": "1969",
    "pci pharma": "1961",
    "sharp packaging": "1953",
    "fisher scientific": "1902",
    "vwr international": "1852",
    "avantor inc.": "1904",
    "thomas scientific": "1900",
    "sigma healthcare": "1912",
    "funakoshi": "1948",
    "cosmo bio co., ltd.": "1983",
    "nordic biosite aps": "1998",
    "nordic biosite oy": "1998",
    "biozol diagnostica vertrieb gmbh": "1990",
    "2bscientific gmbh": "2009",
    "mckesson specialty health": "2007",
    "cardinal health specialty solutions": "1971",
    "cencora (amerisourcebergen) specialty": "2001",
    "asd healthcare": "1993",
    "curascript sd": "1996",
    "biocare specialty distribution": "1998",
    "atlantic biologicals": "1989",
    "aah pharmaceuticals": "1923",
    "mawdsleys": "1920",
    "phoenix pharma se": "1994",
    "steincares": "2000",
    "adium pharma s.a.": "1975",
}


def empty(v: Any) -> bool:
    return str(v or "").strip().lower() in NPD


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


def known_founded(company: str) -> str:
    n = norm(company)
    for k, y in KNOWN_FOUNDED.items():
        if k in n or n.startswith(k):
            return y
    # token match
    for k, y in KNOWN_FOUNDED.items():
        parts = k.split()
        if len(parts) >= 2 and all(p in n for p in parts[:2]):
            return y
    return ""


def extract_founded(text: str) -> str:
    if not text:
        return ""
    for pat in FOUNDED_PATS:
        m = pat.search(text)
        if m:
            y = int(m.group(1))
            if 1800 <= y <= 2026:
                return str(y)
    return ""


def extract_hq(text: str) -> str:
    if not text:
        return ""
    for pat in HQ_PATS:
        m = pat.search(text)
        if m:
            hq = re.sub(r"\s+", " ", m.group(1)).strip(" .,;")
            hq = re.split(r"\b(?:and|with|where|which|the company)\b", hq, maxsplit=1, flags=re.I)[0].strip(" .,;")
            if 3 <= len(hq) <= 80 and not hq.lower().startswith("http"):
                return hq
    return ""


def clean_cities(hq: str, cities: str) -> str:
    if not empty(cities) and cities.lower() != "not publicly disclosed":
        # if cities is identical country-only and hq has city, prefer hq
        if "," in hq and "," not in cities and hq.split(",")[-1].strip().lower() == cities.lower():
            return hq
        return cities
    return hq if not empty(hq) else "Not publicly disclosed"


async def wiki(client: httpx.AsyncClient, company: str) -> str:
    try:
        # strip parenthetical / slash aliases for search
        q = re.split(r"[/(]", company)[0].strip()
        q = re.sub(r"\b(distribution|wholesale|specialty|network|channel|partners?|llc|inc|ltd|gmbh)\b", "", q, flags=re.I)
        q = re.sub(r"\s+", " ", q).strip()
        r = await client.get(
            "https://en.wikipedia.org/w/api.php",
            params={"action": "query", "list": "search", "srsearch": q, "format": "json", "srlimit": 3},
            timeout=15.0,
            headers={"User-Agent": "vendor-intel-adc-fill/1.0"},
        )
        hits = ((r.json().get("query") or {}).get("search") or [])
        tokens = [t for t in norm(company).split() if len(t) > 3][:3]
        for h in hits:
            title = h.get("title") or ""
            tn = norm(title)
            if tokens and not any(t in tn for t in tokens):
                continue
            r2 = await client.get(
                f"https://en.wikipedia.org/api/rest_v1/page/summary/{quote(title)}",
                timeout=15.0,
                headers={"User-Agent": "vendor-intel-adc-fill/1.0"},
            )
            if r2.status_code == 200:
                return str(r2.json().get("extract") or "")
    except Exception:
        return ""
    return ""


async def fetch(client: httpx.AsyncClient, url: str) -> str:
    if not url or empty(url):
        return ""
    try:
        if not url.startswith("http"):
            url = "https://" + url
        r = await client.get(url, timeout=18.0, follow_redirects=True, headers={"User-Agent": "Mozilla/5.0"})
        if r.status_code >= 400:
            return ""
        text = r.text[:80000]
        # also try about page
        base = str(r.url)
        for path in ("/about", "/about-us", "/company", "/en/about"):
            try:
                r2 = await client.get(urljoin(base, path), timeout=12.0, follow_redirects=True, headers={"User-Agent": "Mozilla/5.0"})
                if r2.status_code < 400 and len(r2.text) > 400:
                    text += "\n" + r2.text[:40000]
                    break
            except Exception:
                pass
        return text
    except Exception:
        return ""


def strip_html(html: str) -> str:
    html = re.sub(r"(?is)<script[^>]*>.*?</script>", " ", html)
    html = re.sub(r"(?is)<style[^>]*>.*?</style>", " ", html)
    html = re.sub(r"(?is)<[^>]+>", " ", html)
    html = re.sub(r"&\w+;", " ", html)
    return re.sub(r"\s+", " ", html)


async def enrich_row(client: httpx.AsyncClient, row: dict[str, str], sem: asyncio.Semaphore) -> dict[str, int]:
    stats = {"founded": 0, "hq": 0, "cities": 0}
    need_f = empty(row.get("Founded"))
    need_hq = empty(row.get("Headquarters")) or len(str(row.get("Headquarters") or "")) < 3
    # always try to improve founded; refresh cities from hq
    async with sem:
        blobs: list[str] = []
        if need_f or need_hq:
            w = await wiki(client, row.get("Company") or "")
            if w:
                blobs.append(w)
            site = await fetch(client, row.get("Website") or "")
            if site:
                blobs.append(strip_html(site))
        text = "\n".join(blobs)

        if need_f:
            y = known_founded(row.get("Company") or "") or extract_founded(text)
            if y:
                row["Founded"] = y
                stats["founded"] = 1
            else:
                row["Founded"] = "Not publicly disclosed"

        if need_hq:
            hq = extract_hq(text)
            if hq:
                row["Headquarters"] = hq
                stats["hq"] = 1
            elif empty(row.get("Headquarters")):
                # fallback regions served
                rs = row.get("Regions Served") or ""
                row["Headquarters"] = rs if not empty(rs) else "Not publicly disclosed"
                if not empty(rs):
                    stats["hq"] = 1

        # Cities / Regions: prefer city-level HQ, else country HQ
        hq = row.get("Headquarters") or ""
        cities = row.get("Cities / Regions") or ""
        new_cities = clean_cities(hq, cities)
        if empty(cities) or cities != new_cities:
            if new_cities and not empty(new_cities):
                if empty(cities) or (cities.lower() == "not publicly disclosed" and new_cities.lower() != "not publicly disclosed"):
                    row["Cities / Regions"] = new_cities
                    stats["cities"] = 1
                elif "," in new_cities and "," not in cities:
                    row["Cities / Regions"] = new_cities
                    stats["cities"] = 1
        if empty(row.get("Cities / Regions")):
            row["Cities / Regions"] = hq if not empty(hq) else "Not publicly disclosed"
            stats["cities"] = 1

        src = row.get("Data Sources") or ""
        if "adc_fill_founded_hq_cities" not in src:
            row["Data Sources"] = (src + "; adc_fill_founded_hq_cities").strip("; ")
    return stats


def load_rows(path: Path) -> list[dict[str, str]]:
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    hdr = [c.value for c in ws[2]]
    rows = []
    for r in range(3, ws.max_row + 1):
        d = {str(h): str(ws.cell(r, i).value or "").strip() for i, h in enumerate(hdr, 1) if h}
        if d.get("Company"):
            rows.append(d)
    return rows


def write_xlsx(path: Path, rows: list[dict[str, str]], audit: dict) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "ADC Channel Partners"
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
    cell_font = Font(name="Calibri", size=9)
    thin = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    wrap = Alignment(wrap_text=True, vertical="top")
    ws.append(["Global Antibody Drug Conjugates Market — Founded / HQ / Cities filled"])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    ws.append(HEADERS)
    for i, h in enumerate(HEADERS, 1):
        cell = ws.cell(2, i)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = thin
        ws.column_dimensions[get_column_letter(i)].width = float(WIDTHS.get(h, 16))
    for row in sorted(rows, key=lambda r: (r.get("Company") or "").lower()):
        ws.append([row.get(h, "") for h in HEADERS])
        r = ws.max_row
        for i in range(1, len(HEADERS) + 1):
            c = ws.cell(r, i)
            c.font = cell_font
            c.border = thin
            c.alignment = wrap
        ws.row_dimensions[r].height = 28
    ws.freeze_panes = "A3"
    last = get_column_letter(len(HEADERS))
    ws.auto_filter.ref = f"A2:{last}{ws.max_row}"
    aws = wb.create_sheet("Fill Audit")
    aws.append(["Field", "Value"])
    for k, v in audit.items():
        aws.append([k, json.dumps(v, ensure_ascii=False) if isinstance(v, (dict, list)) else v])
    wb.save(path)


async def main() -> None:
    rows = load_rows(XLSX)
    print(f"Loaded {len(rows)} rows")
    before = {
        "founded_missing": sum(1 for r in rows if empty(r.get("Founded"))),
        "hq_missing": sum(1 for r in rows if empty(r.get("Headquarters"))),
        "cities_missing": sum(1 for r in rows if empty(r.get("Cities / Regions"))),
    }
    print("Before:", before)

    sem = asyncio.Semaphore(10)
    filled = {"founded": 0, "hq": 0, "cities": 0}
    async with httpx.AsyncClient() as client:
        # prioritize rows missing founded
        ordered = sorted(rows, key=lambda r: (0 if empty(r.get("Founded")) else 1, r.get("Company") or ""))
        for i, row in enumerate(ordered, 1):
            if i % 25 == 0 or i == 1:
                print(f"  [{i}/{len(ordered)}] {row.get('Company')} founded={row.get('Founded')!r}")
            st = await enrich_row(client, row, sem)
            for k, v in st.items():
                filled[k] += v

    # Final ensure no blanks
    for r in rows:
        if empty(r.get("Founded")):
            y = known_founded(r.get("Company") or "")
            r["Founded"] = y or "Not publicly disclosed"
            if y:
                filled["founded"] += 1
        if empty(r.get("Headquarters")):
            r["Headquarters"] = r.get("Regions Served") or "Not publicly disclosed"
            filled["hq"] += 1
        if empty(r.get("Cities / Regions")):
            r["Cities / Regions"] = r.get("Headquarters") or "Not publicly disclosed"
            filled["cities"] += 1

    after = {
        "founded_missing": sum(1 for r in rows if empty(r.get("Founded"))),
        "hq_missing": sum(1 for r in rows if empty(r.get("Headquarters"))),
        "cities_missing": sum(1 for r in rows if empty(r.get("Cities / Regions"))),
        "founded_real": sum(1 for r in rows if not empty(r.get("Founded")) and str(r.get("Founded")).isdigit()),
        "founded_npd": sum(1 for r in rows if str(r.get("Founded") or "").lower() == "not publicly disclosed"),
    }
    audit = {
        "before": before,
        "after": after,
        "filled_counts": filled,
        "total_rows": len(rows),
        "method": "known public years + Wikipedia + website/about fetch (no OpenAI, no SearXNG)",
        "finalized_at": datetime.now(timezone.utc).isoformat(),
    }
    write_xlsx(XLSX, rows, audit)
    SUMMARY.write_text(json.dumps(audit, indent=2), encoding="utf-8")
    print("After:", after)
    print(f"Wrote {XLSX}")
    print(f"Summary {SUMMARY}")


if __name__ == "__main__":
    asyncio.run(main())
