#!/usr/bin/env python3
"""Aggressive city-level HQ + founded check for rows still country-only.

Uses Wikipedia + DuckDuckGo HTML + website. No OpenAI, no SearXNG.
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
SUMMARY = OUT_DIR / "adc_aggressive_city_hq_summary.json"

NPD = {"", "n/a", "na", "none", "-", "—", "not publicly disclosed", "unknown", "tbd", "null"}

YEAR_PATS = [
    re.compile(
        r"(?:founded|established|incorpora(?:ted|tion)|est\.?|since|created|started|"
        r"設立|創立|創業|gegründet|gründung)\D{0,20}((?:18|19|20)\d{2})",
        re.I,
    ),
]
HQ_PATS = [
    re.compile(
        r"(?:headquarter(?:ed|s)?|based|located|head office|registered office|"
        r"hauptsitz|sitz)\s*(?:in|at|:)?\s*([A-Z][^.\n|]{2,80})",
        re.I,
    ),
]


def empty(v: Any) -> bool:
    return str(v or "").strip().lower() in NPD


def is_country_only(hq: str) -> bool:
    hq = (hq or "").strip()
    if not hq:
        return True
    if "," in hq:
        return False
    # city-like multiword without comma still ok if long
    parts = hq.split()
    countries = {
        "united states", "united kingdom", "united arab emirates", "saudi arabia",
        "south korea", "south africa", "new zealand", "hong kong", "czech republic",
        "bosnia and herzegovina", "north macedonia", "sri lanka", "costa rica",
    }
    if hq.lower() in countries:
        return True
    if len(parts) <= 2 and hq[0].isupper():
        # Germany, Japan, Brazil, India, China, Australia, Canada, Mexico, Spain...
        return True
    return False


def clean_hq(raw: str) -> str:
    hq = re.sub(r"\s+", " ", raw or "").strip(" .,;:|-")
    # drop leading verb fragments
    hq = re.sub(r"^(?:are |is |was |were |located |based |in |at )+", "", hq, flags=re.I)
    hq = re.split(
        r"\b(?:and is|where|which|the company|providing|with a|with \$|employees|revenue|and has)\b",
        hq,
        maxsplit=1,
        flags=re.I,
    )[0]
    hq = hq.strip(" .,;:|-")
    # keep only up to country-ish end: City, State/Country
    # cut if too many commas
    parts = [p.strip() for p in hq.split(",") if p.strip()]
    if len(parts) > 3:
        parts = parts[:3]
        hq = ", ".join(parts)
    if not (3 <= len(hq) <= 70):
        return ""
    if hq.lower().startswith("http"):
        return ""
    # reject if contains digits money or weird tokens
    if re.search(r"[\$€£%]|\demployees\b|\drevenue\b|\bwith \d", hq, re.I):
        return ""
    if re.search(r"\b(located|based|headquarter)\b", hq, re.I):
        return ""
    # must look like place: at least one capital letter word
    if not re.search(r"[A-ZÀ-ÖØ-Þ]", hq):
        return ""
    return hq


def looks_city_level(hq: str) -> bool:
    hq = (hq or "").strip()
    if not hq or is_country_only(hq):
        return False
    # require comma OR known city pattern
    if "," in hq:
        return True
    return False


def strip_html(html: str) -> str:
    html = re.sub(r"(?is)<script[^>]*>.*?</script>", " ", html)
    html = re.sub(r"(?is)<style[^>]*>.*?</style>", " ", html)
    html = re.sub(r"(?is)<[^>]+>", " ", html)
    return re.sub(r"\s+", " ", html)


async def wiki(client: httpx.AsyncClient, company: str) -> str:
    try:
        q = re.split(r"[/(]", company)[0].strip()
        r = await client.get(
            "https://en.wikipedia.org/w/api.php",
            params={"action": "query", "list": "search", "srsearch": q, "format": "json", "srlimit": 5},
            timeout=20.0,
            headers={"User-Agent": "vendor-intel-adc-city/1.0"},
        )
        hits = ((r.json().get("query") or {}).get("search") or [])
        toks = [t for t in q.lower().split() if len(t) > 3][:3]
        for h in hits:
            title = h.get("title") or ""
            if toks and not any(t in title.lower() for t in toks):
                continue
            r2 = await client.get(
                f"https://en.wikipedia.org/api/rest_v1/page/summary/{quote(title)}",
                timeout=20.0,
                headers={"User-Agent": "vendor-intel-adc-city/1.0"},
            )
            if r2.status_code == 200:
                j = r2.json()
                return f"{title}. {j.get('description','')}. {j.get('extract','')}"
    except Exception:
        return ""
    return ""


async def ddg(client: httpx.AsyncClient, q: str) -> str:
    try:
        r = await client.get(
            "https://html.duckduckgo.com/html/",
            params={"q": q},
            timeout=25.0,
            headers={"User-Agent": "Mozilla/5.0"},
        )
        if r.status_code != 200:
            return ""
        return strip_html(r.text)[:16000]
    except Exception:
        return ""


async def site(client: httpx.AsyncClient, url: str) -> str:
    if not url or empty(url):
        return ""
    try:
        if not url.startswith("http"):
            url = "https://" + url
        r = await client.get(url, timeout=20.0, follow_redirects=True, headers={"User-Agent": "Mozilla/5.0"})
        if r.status_code >= 400:
            return ""
        text = strip_html(r.text[:60000])
        base = str(r.url)
        for path in ("/about", "/about-us", "/company", "/contact", "/en/about"):
            try:
                r2 = await client.get(urljoin(base, path), timeout=15.0, follow_redirects=True, headers={"User-Agent": "Mozilla/5.0"})
                if r2.status_code < 400 and len(r2.text) > 300:
                    text += " " + strip_html(r2.text[:30000])
                    break
            except Exception:
                pass
        return text
    except Exception:
        return ""


def extract(blob: str) -> tuple[str | None, str | None]:
    years = []
    for pat in YEAR_PATS:
        for m in pat.finditer(blob or ""):
            y = int(m.group(1))
            if 1800 <= y <= 2026:
                years.append(y)
    year = str(min(years)) if years else None  # earliest founded often correct
    # if many years, prefer most frequent
    if years:
        from collections import Counter
        year = str(Counter(years).most_common(1)[0][0])
    hq = None
    for pat in HQ_PATS:
        m = pat.search(blob or "")
        if m:
            hq = clean_hq(m.group(1))
            if hq:
                break
    return year, hq


async def main() -> None:
    wb = load_workbook(XLSX, data_only=True)
    ws = wb.active
    headers = [c.value for c in ws[2]]
    rows = []
    for r in range(3, ws.max_row + 1):
        d = {str(h): str(ws.cell(r, i).value or "").strip() for i, h in enumerate(headers, 1) if h}
        if d.get("Company"):
            rows.append(d)

    targets = [r for r in rows if is_country_only(r.get("Headquarters") or "")]
    print(f"Total {len(rows)}; country-only HQ targets {len(targets)}")

    sem = asyncio.Semaphore(6)
    hq_chg = founded_chg = 0
    changes = []

    async with httpx.AsyncClient() as client:
        async def one(row: dict[str, str]) -> None:
            nonlocal hq_chg, founded_chg
            async with sem:
                company = row["Company"]
                q = re.split(r"[/(]", company)[0].strip()
                blobs = []
                for txt in [
                    await wiki(client, company),
                    await site(client, row.get("Website") or ""),
                    await ddg(client, f'"{q}" headquarters OR "headquartered" OR "based in" OR founded'),
                    await ddg(client, f"{q} company founded established HQ city"),
                ]:
                    if txt:
                        blobs.append(txt)
                blob = "\n".join(blobs)
                year, hq = extract(blob)
                before = {"f": row.get("Founded"), "h": row.get("Headquarters"), "c": row.get("Cities / Regions")}
                changed = False
                if hq and looks_city_level(hq):
                    # keep country suffix if useful
                    country = before["h"]
                    if (
                        country
                        and country.lower() not in hq.lower()
                        and len(country.split()) <= 3
                        and is_country_only(country)
                    ):
                        hq = f"{hq}, {country}"
                    if looks_city_level(hq):
                        row["Headquarters"] = hq
                        row["Cities / Regions"] = hq
                        hq_chg += 1
                        changed = True
                if year and year.isdigit():
                    cur = str(row.get("Founded") or "")
                    # update if current looks rounded/guessy OR differs by >3 and evidence present
                    if cur != year:
                        if cur.endswith("0") or cur.endswith("5") or abs(int(cur) - int(year)) >= 3:
                            # only if year appears with founded language confidence
                            if re.search(rf"(?:founded|established|est\.?|since).{{0,20}}{year}", blob, re.I):
                                row["Founded"] = year
                                founded_chg += 1
                                changed = True
                if changed:
                    changes.append(
                        {
                            "company": company,
                            "before": before,
                            "after": {
                                "f": row.get("Founded"),
                                "h": row.get("Headquarters"),
                                "c": row.get("Cities / Regions"),
                            },
                            "blob_len": len(blob),
                        }
                    )
                src = row.get("Data Sources") or ""
                if "adc_aggressive_city_hq" not in src:
                    row["Data Sources"] = (src + "; adc_aggressive_city_hq").strip("; ")

        # process in batches for logging
        for i, row in enumerate(targets, 1):
            await one(row)
            if i % 25 == 0 or i == 1:
                print(f"[{i}/{len(targets)}] {row['Company']} -> {row.get('Headquarters')} (hq+{hq_chg} f+{founded_chg})")

    # ensure cities mirrors hq
    for r in rows:
        if empty(r.get("Cities / Regions")):
            r["Cities / Regions"] = r.get("Headquarters") or "Not publicly disclosed"
        if empty(r.get("Headquarters")):
            r["Headquarters"] = r.get("Cities / Regions") or "Not publicly disclosed"
        if empty(r.get("Founded")):
            r["Founded"] = "Not publicly disclosed"

    # write excel
    wb2 = Workbook()
    ws2 = wb2.active
    ws2.title = "ADC Channel Partners"
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
    ws2.append(["Global Antibody Drug Conjugates Market — each company Founded/HQ/Cities checked"])
    ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    hdr = [h for h in headers if h]
    ws2.append(hdr)
    widths = {"Company": 28, "Headquarters": 32, "Cities / Regions": 32, "Summary": 36, "Website": 24}
    for i, h in enumerate(hdr, 1):
        cell = ws2.cell(2, i)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = thin
        ws2.column_dimensions[get_column_letter(i)].width = float(widths.get(h, 16))
    for row in sorted(rows, key=lambda r: (r.get("Company") or "").lower()):
        ws2.append([row.get(h, "") for h in hdr])
        rr = ws2.max_row
        for i in range(1, len(hdr) + 1):
            c = ws2.cell(rr, i)
            c.font = cell_font
            c.border = thin
            c.alignment = wrap
        ws2.row_dimensions[rr].height = 28
    ws2.freeze_panes = "A3"
    last = get_column_letter(len(hdr))
    ws2.auto_filter.ref = f"A2:{last}{ws2.max_row}"

    city_level = sum(1 for r in rows if not is_country_only(r.get("Headquarters") or ""))
    audit = {
        "targets_country_only": len(targets),
        "hq_upgraded_to_city": hq_chg,
        "founded_updated": founded_chg,
        "city_level_hq_after": city_level,
        "country_only_hq_remaining": sum(1 for r in rows if is_country_only(r.get("Headquarters") or "")),
        "blank_founded": sum(1 for r in rows if not (r.get("Founded") or "").strip()),
        "blank_hq": sum(1 for r in rows if not (r.get("Headquarters") or "").strip()),
        "blank_cities": sum(1 for r in rows if not (r.get("Cities / Regions") or "").strip()),
        "digit_founded": sum(1 for r in rows if str(r.get("Founded") or "").isdigit()),
        "changes_sample": changes[:50],
        "changes_count": len(changes),
        "finalized_at": datetime.now(timezone.utc).isoformat(),
    }
    aws = wb2.create_sheet("Aggressive HQ Audit")
    aws.append(["Field", "Value"])
    for k, v in audit.items():
        aws.append([k, json.dumps(v, ensure_ascii=False) if isinstance(v, (dict, list)) else v])
    cws = wb2.create_sheet("HQ Changes")
    cws.append(["Company", "Founded Before", "Founded After", "HQ Before", "HQ After"])
    for ch in changes:
        cws.append(
            [
                ch["company"],
                ch["before"]["f"],
                ch["after"]["f"],
                ch["before"]["h"],
                ch["after"]["h"],
            ]
        )
    wb2.save(XLSX)
    SUMMARY.write_text(json.dumps({k: v for k, v in audit.items() if k != "changes_sample"}, indent=2, ensure_ascii=False), encoding="utf-8")
    print("DONE", json.dumps({k: audit[k] for k in audit if k != "changes_sample"}, indent=2))


if __name__ == "__main__":
    asyncio.run(main())
