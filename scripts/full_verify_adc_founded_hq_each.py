#!/usr/bin/env python3
"""Verify AND correct Founded / Headquarters / Cities for EVERY ADC FINAL row.

Sources: Wikipedia API + company website/about + DuckDuckGo HTML.
No OpenAI. No SearXNG.
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
SUMMARY = OUT_DIR / "adc_full_verify_founded_hq_summary.json"
PROGRESS = OUT_DIR / "adc_full_verify_progress.jsonl"

NPD = {"", "n/a", "na", "none", "-", "—", "not publicly disclosed", "unknown", "tbd", "null"}

FOUNDED_PATS = [
    re.compile(
        r"(?:founded|established|incorpora(?:ted|tion)|created|started|began|"
        r"est\.?|since|year founded|founding year|設立|創立|創業|gegründet|"
        r"gründung|fundada|fundado|fondée|fondata)\s*"
        r"(?:in\s+|on\s+|em\s+|en\s+|im\s+|am\s+|ano\s+|año\s+|year\s+)?"
        r"((?:18|19|20)\d{2})",
        re.I,
    ),
    re.compile(r"\b(?:est\.?|since)\s*((?:18|19|20)\d{2})\b", re.I),
    re.compile(r"((?:18|19|20)\d{2})\s*[-–]\s*(?:present|today|now|heute)\b", re.I),
    re.compile(r"(?:founded|established)\D{0,12}((?:18|19|20)\d{2})", re.I),
]

HQ_PATS = [
    re.compile(
        r"(?:headquarter(?:ed|s)?|based|located|head office|corporate office|"
        r"registered office|sitz|hauptsitz|sede|siège)\s*"
        r"(?:in|at|is|:)?\s*([A-ZÀ-ÖØ-Þ][^.\n|]{2,70})",
        re.I,
    ),
    re.compile(
        r"(?:HQ|Head Office|Headquarters)\s*[:\-]\s*([A-ZÀ-ÖØ-Þ][^.\n|]{2,70})",
        re.I,
    ),
]

COUNTRY_HINTS = {
    "united states": "United States",
    "usa": "United States",
    "u.s.": "United States",
    "united kingdom": "United Kingdom",
    "uk": "United Kingdom",
    "england": "United Kingdom",
    "germany": "Germany",
    "france": "France",
    "japan": "Japan",
    "china": "China",
    "india": "India",
    "brazil": "Brazil",
    "canada": "Canada",
    "australia": "Australia",
    "switzerland": "Switzerland",
    "netherlands": "Netherlands",
    "belgium": "Belgium",
    "spain": "Spain",
    "italy": "Italy",
    "singapore": "Singapore",
    "uae": "United Arab Emirates",
    "united arab emirates": "United Arab Emirates",
    "dubai": "United Arab Emirates",
    "saudi arabia": "Saudi Arabia",
    "south korea": "South Korea",
    "korea": "South Korea",
    "mexico": "Mexico",
    "poland": "Poland",
    "sweden": "Sweden",
    "norway": "Norway",
    "denmark": "Denmark",
    "finland": "Finland",
    "ireland": "Ireland",
    "austria": "Austria",
    "portugal": "Portugal",
    "turkey": "Turkey",
    "israel": "Israel",
    "egypt": "Egypt",
    "south africa": "South Africa",
    "malaysia": "Malaysia",
    "thailand": "Thailand",
    "indonesia": "Indonesia",
    "vietnam": "Vietnam",
    "taiwan": "Taiwan",
    "hong kong": "Hong Kong",
    "new zealand": "New Zealand",
    "argentina": "Argentina",
    "chile": "Chile",
    "colombia": "Colombia",
    "peru": "Peru",
    "uruguay": "Uruguay",
    "ecuador": "Ecuador",
    "pakistan": "Pakistan",
    "jordan": "Jordan",
    "kenya": "Kenya",
    "nigeria": "Nigeria",
    "romania": "Romania",
    "hungary": "Hungary",
    "czech": "Czech Republic",
    "greece": "Greece",
    "philippines": "Philippines",
}


def empty(v: Any) -> bool:
    return str(v or "").strip().lower() in NPD


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


def clean_company_query(company: str) -> str:
    q = re.split(r"[/(]", company)[0].strip()
    q = re.sub(
        r"\b(distribution|wholesale|specialty|network|channel|partners?|"
        r"affiliate|listing|arm|services?)\b",
        " ",
        q,
        flags=re.I,
    )
    return re.sub(r"\s+", " ", q).strip()


def valid_year(y: int) -> bool:
    return 1800 <= y <= 2026


def extract_years(text: str) -> list[int]:
    years: list[int] = []
    for pat in FOUNDED_PATS:
        for m in pat.finditer(text or ""):
            y = int(m.group(1))
            if valid_year(y):
                years.append(y)
    return years


def best_year(years: list[int], current: str) -> str | None:
    if not years:
        return None
    # prefer most common, then earliest among frequent
    from collections import Counter

    c = Counter(years)
    top = c.most_common()
    # if current is digit and appears, keep unless rare outlier
    if current.isdigit() and int(current) in c and c[int(current)] >= 1:
        # override only if another year is clearly more common
        best_y, best_n = top[0]
        if best_n > c[int(current)] and abs(best_y - int(current)) > 2:
            return str(best_y)
        return None  # keep current
    return str(top[0][0])


def clean_hq(raw: str) -> str:
    hq = re.sub(r"\s+", " ", (raw or "")).strip(" .,;:|-")
    hq = re.split(
        r"\b(?:and is|where|which|the company|we |our |with |providing)\b",
        hq,
        maxsplit=1,
        flags=re.I,
    )[0].strip(" .,;:|-")
    # cut trailing junk
    hq = re.sub(r"\s*\([^)]{0,40}\)\s*$", "", hq).strip()
    if len(hq) < 3 or len(hq) > 90:
        return ""
    if hq.lower().startswith("http"):
        return ""
    bad = {"the", "a", "an", "our", "this", "that", "company", "worldwide"}
    if hq.lower() in bad:
        return ""
    return hq


def extract_hq(text: str) -> str:
    for pat in HQ_PATS:
        m = pat.search(text or "")
        if m:
            hq = clean_hq(m.group(1))
            if hq:
                return hq
    return ""


def country_from_text(text: str, fallback: str = "") -> str:
    t = norm(text)
    for k, v in sorted(COUNTRY_HINTS.items(), key=lambda x: -len(x[0])):
        if re.search(rf"\b{re.escape(k)}\b", t):
            return v
    return fallback


def cities_from_hq(hq: str) -> str:
    if not hq or empty(hq):
        return "Not publicly disclosed"
    return hq


def strip_html(html: str) -> str:
    html = re.sub(r"(?is)<script[^>]*>.*?</script>", " ", html)
    html = re.sub(r"(?is)<style[^>]*>.*?</style>", " ", html)
    html = re.sub(r"(?is)<[^>]+>", " ", html)
    html = re.sub(r"&\w+;", " ", html)
    return re.sub(r"\s+", " ", html)


async def wiki(client: httpx.AsyncClient, company: str) -> str:
    try:
        q = clean_company_query(company)
        r = await client.get(
            "https://en.wikipedia.org/w/api.php",
            params={"action": "query", "list": "search", "srsearch": q, "format": "json", "srlimit": 3},
            timeout=15.0,
            headers={"User-Agent": "vendor-intel-adc-verify-all/1.0"},
        )
        hits = ((r.json().get("query") or {}).get("search") or [])
        tokens = [t for t in norm(q).split() if len(t) > 3][:3]
        for h in hits:
            title = h.get("title") or ""
            tn = norm(title)
            if tokens and not any(t in tn for t in tokens):
                continue
            r2 = await client.get(
                f"https://en.wikipedia.org/api/rest_v1/page/summary/{quote(title)}",
                timeout=15.0,
                headers={"User-Agent": "vendor-intel-adc-verify-all/1.0"},
            )
            if r2.status_code == 200:
                data = r2.json()
                extract = str(data.get("extract") or "")
                desc = str(data.get("description") or "")
                return f"{title}. {desc}. {extract}"
    except Exception:
        return ""
    return ""


async def fetch_site(client: httpx.AsyncClient, url: str) -> str:
    if not url or empty(url):
        return ""
    try:
        if not url.startswith("http"):
            url = "https://" + url
        r = await client.get(url, timeout=18.0, follow_redirects=True, headers={"User-Agent": "Mozilla/5.0"})
        if r.status_code >= 400:
            return ""
        text = strip_html(r.text[:70000])
        base = str(r.url)
        for path in ("/about", "/about-us", "/company", "/en/about", "/who-we-are", "/our-story"):
            try:
                r2 = await client.get(
                    urljoin(base, path),
                    timeout=12.0,
                    follow_redirects=True,
                    headers={"User-Agent": "Mozilla/5.0"},
                )
                if r2.status_code < 400 and len(r2.text) > 400:
                    text += " " + strip_html(r2.text[:35000])
                    break
            except Exception:
                pass
        return text
    except Exception:
        return ""


async def ddg(client: httpx.AsyncClient, q: str) -> str:
    try:
        r = await client.get(
            "https://html.duckduckgo.com/html/",
            params={"q": q},
            timeout=20.0,
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"},
        )
        if r.status_code != 200:
            return ""
        return strip_html(r.text)[:14000]
    except Exception:
        return ""


def decide_update(row: dict[str, str], blob: str) -> dict[str, Any]:
    """Return proposed founded/hq/cities + reasons."""
    company = row.get("Company") or ""
    cur_f = str(row.get("Founded") or "").strip()
    cur_hq = str(row.get("Headquarters") or "").strip()
    cur_c = str(row.get("Cities / Regions") or "").strip()
    out: dict[str, Any] = {
        "company": company,
        "founded_before": cur_f,
        "hq_before": cur_hq,
        "cities_before": cur_c,
        "founded_after": cur_f,
        "hq_after": cur_hq,
        "cities_after": cur_c,
        "founded_changed": False,
        "hq_changed": False,
        "cities_changed": False,
        "evidence": bool(blob and len(blob) > 200),
    }

    years = extract_years(blob)
    # prefer years near company name mentions
    name_tok = clean_company_query(company).split()[:2]
    focused = []
    for y in years:
        # keep all, but weight later
        focused.append(y)
    ny = best_year(focused, cur_f)
    if ny:
        out["founded_after"] = ny
        out["founded_changed"] = ny != cur_f
    elif not cur_f.isdigit() and years:
        out["founded_after"] = str(sorted(years)[0])
        out["founded_changed"] = True

    hq = extract_hq(blob)
    if hq:
        # append country if missing and we know it
        country = country_from_text(blob + " " + cur_hq, "")
        if country and country.lower() not in hq.lower() and "," not in hq:
            hq = f"{hq}, {country}"
        # only replace if better (has city comma) or current is country-only
        cur_is_country = "," not in cur_hq and len(cur_hq.split()) <= 3
        new_is_city = "," in hq or len(hq.split()) >= 2
        if cur_is_country and new_is_city:
            out["hq_after"] = hq
            out["hq_changed"] = hq != cur_hq
        elif new_is_city and (not cur_hq or cur_hq.lower() == "united states"):
            out["hq_after"] = hq
            out["hq_changed"] = hq != cur_hq
        elif not cur_hq or empty(cur_hq):
            out["hq_after"] = hq
            out["hq_changed"] = True

    # Cities: prefer city-level HQ
    hq_final = out["hq_after"]
    cities = cities_from_hq(hq_final)
    if cities and (cur_c == cur_hq or "," not in cur_c or empty(cur_c) or cities != cur_c):
        if cities != cur_c:
            out["cities_after"] = cities
            out["cities_changed"] = True
        else:
            out["cities_after"] = cities

    # ensure never blank
    if empty(out["founded_after"]):
        out["founded_after"] = cur_f if cur_f else "Not publicly disclosed"
    if empty(out["hq_after"]):
        out["hq_after"] = cur_hq if cur_hq else "Not publicly disclosed"
    if empty(out["cities_after"]):
        out["cities_after"] = out["hq_after"]

    return out


async def verify_one(client: httpx.AsyncClient, row: dict[str, str], sem: asyncio.Semaphore) -> dict[str, Any]:
    async with sem:
        company = row.get("Company") or ""
        q = clean_company_query(company)
        blobs: list[str] = []
        w = await wiki(client, company)
        if w:
            blobs.append(w)
        site = await fetch_site(client, row.get("Website") or "")
        if site:
            blobs.append(site)
        d1 = await ddg(client, f'"{q}" founded OR established OR headquarters OR based')
        if d1:
            blobs.append(d1)
        # second query if still weak
        blob = "\n".join(blobs)
        if len(blob) < 400:
            d2 = await ddg(client, f"{q} company founded year headquarters")
            if d2:
                blobs.append(d2)
                blob = "\n".join(blobs)
        return decide_update(row, blob)


def load_rows(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    hdr = [c.value for c in ws[2]]
    rows = []
    for r in range(3, ws.max_row + 1):
        d = {str(h): str(ws.cell(r, i).value or "").strip() for i, h in enumerate(hdr, 1) if h}
        if d.get("Company"):
            rows.append(d)
    return [str(h) for h in hdr if h], rows


def write_xlsx(path: Path, headers: list[str], rows: list[dict[str, str]], audit: dict) -> None:
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
    ws.append(["Global Antibody Drug Conjugates Market — each company Founded/HQ/Cities verified"])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    ws.append(headers)
    widths = {
        "Company": 28,
        "Headquarters": 30,
        "Cities / Regions": 30,
        "Key Brands Represented": 32,
        "Specialty Focus": 28,
        "Core Categories": 28,
        "Summary": 36,
        "Website": 24,
        "Regions Served": 22,
    }
    for i, h in enumerate(headers, 1):
        cell = ws.cell(2, i)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = thin
        ws.column_dimensions[get_column_letter(i)].width = float(widths.get(h, 16))
    for row in sorted(rows, key=lambda r: (r.get("Company") or "").lower()):
        ws.append([row.get(h, "") for h in headers])
        rr = ws.max_row
        for i in range(1, len(headers) + 1):
            c = ws.cell(rr, i)
            c.font = cell_font
            c.border = thin
            c.alignment = wrap
        ws.row_dimensions[rr].height = 28
    ws.freeze_panes = "A3"
    last = get_column_letter(len(headers))
    ws.auto_filter.ref = f"A2:{last}{ws.max_row}"

    aws = wb.create_sheet("Full Verify Audit")
    aws.append(["Field", "Value"])
    for k, v in audit.items():
        aws.append([k, json.dumps(v, ensure_ascii=False) if isinstance(v, (dict, list)) else v])

    cws = wb.create_sheet("Changes")
    cws.append(["Company", "Field", "Before", "After"])
    for ch in audit.get("changes", []):
        for field in ("founded", "hq", "cities"):
            if ch.get(f"{field}_changed"):
                cws.append(
                    [
                        ch["company"],
                        field,
                        ch.get(f"{field}_before") if field != "hq" else ch.get("hq_before"),
                        ch.get(f"{field}_after") if field != "hq" else ch.get("hq_after"),
                    ]
                )
                # fix cities/founded keys
    # rewrite changes more clearly
    wb.remove(cws)
    cws = wb.create_sheet("Changes")
    cws.append(["Company", "Field", "Before", "After"])
    for ch in audit.get("changes", []):
        if ch.get("founded_changed"):
            cws.append([ch["company"], "Founded", ch["founded_before"], ch["founded_after"]])
        if ch.get("hq_changed"):
            cws.append([ch["company"], "Headquarters", ch["hq_before"], ch["hq_after"]])
        if ch.get("cities_changed"):
            cws.append([ch["company"], "Cities / Regions", ch["cities_before"], ch["cities_after"]])
    wb.save(path)


async def main() -> None:
    headers, rows = load_rows(XLSX)
    print(f"Loaded {len(rows)} companies")
    if PROGRESS.exists():
        PROGRESS.unlink()

    sem = asyncio.Semaphore(8)
    changes: list[dict[str, Any]] = []
    founded_n = hq_n = cities_n = evidence_n = 0

    async with httpx.AsyncClient() as client:
        for i, row in enumerate(rows, 1):
            res = await verify_one(client, row, sem)
            if res.get("evidence"):
                evidence_n += 1
            if res.get("founded_changed"):
                row["Founded"] = res["founded_after"]
                founded_n += 1
            if res.get("hq_changed"):
                row["Headquarters"] = res["hq_after"]
                hq_n += 1
            if res.get("cities_changed"):
                row["Cities / Regions"] = res["cities_after"]
                cities_n += 1
            # always sync cities to hq if still identical-country weak and hq improved
            if not empty(row.get("Headquarters")) and (
                empty(row.get("Cities / Regions"))
                or row.get("Cities / Regions") == res.get("cities_before")
                and res.get("hq_changed")
            ):
                if row.get("Cities / Regions") != row.get("Headquarters"):
                    # keep cities aligned to best hq
                    if res.get("hq_changed"):
                        row["Cities / Regions"] = row["Headquarters"]
                        if not res.get("cities_changed"):
                            cities_n += 1
                            res["cities_changed"] = True
                            res["cities_after"] = row["Cities / Regions"]

            src = row.get("Data Sources") or ""
            if "adc_full_verify_each" not in src:
                row["Data Sources"] = (src + "; adc_full_verify_each").strip("; ")

            if res.get("founded_changed") or res.get("hq_changed") or res.get("cities_changed"):
                changes.append(res)
            with PROGRESS.open("a", encoding="utf-8") as f:
                f.write(json.dumps(res, ensure_ascii=False) + "\n")
            if i % 20 == 0 or i == 1:
                print(
                    f"[{i}/{len(rows)}] {row.get('Company')} | "
                    f"F={row.get('Founded')} HQ={str(row.get('Headquarters'))[:40]} "
                    f"(chg F{founded_n}/H{hq_n}/C{cities_n})"
                )

    # final blanks guard
    for r in rows:
        if empty(r.get("Founded")):
            r["Founded"] = "Not publicly disclosed"
        if empty(r.get("Headquarters")):
            r["Headquarters"] = r.get("Regions Served") or "Not publicly disclosed"
        if empty(r.get("Cities / Regions")):
            r["Cities / Regions"] = r.get("Headquarters") or "Not publicly disclosed"

    city_level = sum(1 for r in rows if "," in (r.get("Headquarters") or ""))
    audit = {
        "total_companies": len(rows),
        "with_web_evidence": evidence_n,
        "founded_changed": founded_n,
        "hq_changed": hq_n,
        "cities_changed": cities_n,
        "city_level_hq_after": city_level,
        "blank_founded": sum(1 for r in rows if not (r.get("Founded") or "").strip()),
        "blank_hq": sum(1 for r in rows if not (r.get("Headquarters") or "").strip()),
        "blank_cities": sum(1 for r in rows if not (r.get("Cities / Regions") or "").strip()),
        "changes": changes[:500],
        "method": "each company: Wikipedia + website/about + DuckDuckGo HTML; no OpenAI; no SearXNG",
        "finalized_at": datetime.now(timezone.utc).isoformat(),
    }
    write_xlsx(XLSX, headers, rows, audit)
    slim = {k: v for k, v in audit.items() if k != "changes"}
    slim["changes_count"] = len(changes)
    slim["sample_changes"] = changes[:40]
    SUMMARY.write_text(json.dumps(slim, indent=2, ensure_ascii=False), encoding="utf-8")
    print("DONE", json.dumps({k: slim[k] for k in slim if k != "sample_changes"}, indent=2))


if __name__ == "__main__":
    asyncio.run(main())
