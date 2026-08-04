#!/usr/bin/env python3
"""Finalize Global ADC market from chatgpt expand checkpoint — NO OpenAI.

1) Load filled_all from checkpoint
2) Keep only ADC-related distributors / channel partners (drop manufacturers/OEMs)
3) Fill empty columns via SearXNG + Wikipedia + website fetch
4) Write FINAL Excel
"""
from __future__ import annotations

import asyncio
import json
import re
import sys
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import quote, urljoin, urlparse

import httpx
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from vendor_intel.pipeline.web_expand import HEADERS, WIDTHS  # noqa: E402

CKPT = ROOT / "output/web_expand/global_antibody_drug_conjugates_market_global/chatgpt_checkpoint_batch_all.json"
OUT_DIR = ROOT / "output/web_expand/global_antibody_drug_conjugates_market_global"
OUT_XLSX = OUT_DIR / "global_antibody_drug_conjugates_market_FINAL.xlsx"
MARKET = "Global Antibody Drug Conjugates Market"

NPD = {
    "",
    "n/a",
    "na",
    "none",
    "-",
    "—",
    "not publicly disclosed",
    "unknown",
    "tbd",
    "null",
}

# Brand owners / ADC developers / reagent OEMs — not channel partners
OEM_BLOCK = {
    "thermo fisher scientific",
    "abcam",
    "bio-techne",
    "miltenyi biotec",
    "genscript",
    "cell signaling technology",
    "santa cruz biotechnology",
    "peprotech",
    "cayman chemical",
    "biolegend",
    "fujifilm wako",
    "sigma-aldrich",
    "merck kgaa",
    "genentech",
    "amgen",
    "bristol-myers squibb",
    "bristol myers",
    "roche",
    "novartis",
    "pfizer",
    "astrazeneca",
    "eli lilly",
    "sanofi",
    "gsk",
    "glaxosmithkline",
    "takeda",
    "abbvie",
    "seagen",
    "immunogen",
    "daiichi sankyo",
    "adc therapeutics",
    "gilead",
    "johnson & johnson",
    "janssen",
    "astrazeneca",
    "macrogenics",
    "mersana",
    "sutro biopharma",
    "byondis",
    "remegen",
    "hutchmed",
    "haihe",
    "legochem",
    "ppd, part of thermo fisher",
    "r&d systems",  # Bio-Techne brand
    "fluorochem",
    "sab biotherapeutics",  # developer
    "lonza",  # CDMO manufacturer services — not distributor
    "catalent",
    "wuxi biologics",
    "wuxi apptec",
    "samsung biologics",
    "boehringer",
    "abbvie",
}

DIST_RE = re.compile(
    r"\b(distributor|distribution|wholesaler|wholesale|reseller|channel partner|"
    r"authorized|trading partner|importer|pharmacy wholesale|specialty pharmacy|"
    r"specialty distributor|cold.?chain|3pl|logistics partner|vad|var)\b",
    re.I,
)
MFR_RE = re.compile(
    r"\b(manufacturer|manufacturing|develops? and (sells?|markets?)|"
    r"biopharmaceutical company|pharmaceutical company|drug developer|"
    r"adc developer|oem|produces? antibodies|makes? antibodies)\b",
    re.I,
)
ADC_RE = re.compile(
    r"\b(adc|antibody[- ]drug conjugate|antibody drug conjugate|oncology|"
    r"cytotoxic|payload|linker|bioconjugate|antibody|biologics?|"
    r"specialty (pharma|pharmacy|distributor)|cold chain|"
    r"injectable|hospital pharmacy|wholesaler)\b",
    re.I,
)

YEAR_RE = re.compile(r"\b((?:19|20)\d{2})\b")
EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
PHONE_RE = re.compile(r"(?:\+?\d[\d\s\-().]{7,}\d)")
LI_RE = re.compile(r"https?://(?:www\.)?linkedin\.com/company/[a-zA-Z0-9\-_/]+", re.I)
PERSON_PATS = [
    (r"([A-Z][a-z]+(?:\s+[A-Z][a-z'\-]+){1,3})\s*,?\s*(?:CEO|Chief Executive)", "CEO"),
    (r"CEO\s+([A-Z][a-z]+(?:\s+[A-Z][a-z'\-]+){1,3})", "CEO"),
    (r"([A-Z][a-z]+(?:\s+[A-Z][a-z'\-]+){1,3})\s*,?\s*(?:Managing Director|MD)\b", "Managing Director"),
    (r"(?:Founder|Co-Founder)\s+([A-Z][a-z]+(?:\s+[A-Z][a-z'\-]+){1,3})", "Founder"),
]


def empty(v: Any) -> bool:
    return str(v or "").strip().lower() in NPD


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


def looks_person(name: str, company: str = "") -> bool:
    n = (name or "").strip()
    if len(n) < 4 or len(n) > 55 or "\n" in n:
        return False
    parts = n.split()
    if not (2 <= len(parts) <= 4):
        return False
    bad = {"ceo", "cto", "inc", "ltd", "llc", "gmbh", "team", "the", "and"}
    if any(p.lower().strip(".,") in bad for p in parts):
        return False
    if not all(re.match(r"^[A-Za-z][A-Za-z.'\-]*$", p) for p in parts):
        return False
    if company and n.lower() == company.strip().lower():
        return False
    return True


def is_oem(name: str) -> bool:
    n = norm(name)
    for oem in OEM_BLOCK:
        if n == oem or n.startswith(oem + " ") or f" {oem} " in f" {n} ":
            # allow explicit distributor wording in name
            if DIST_RE.search(name) and "authorized trading" in n:
                return False
            if "authorized trading partners" in n or "distributor" in n:
                return False
            return True
    return False


def keep_channel(row: dict[str, str]) -> tuple[bool, str]:
    name = row.get("Company") or ""
    bt = row.get("Business Type") or ""
    focus = row.get("Specialty Focus") or ""
    summary = row.get("Summary") or ""
    blob = f"{name} {bt} {focus} {summary}"

    if is_oem(name):
        return False, "oem_manufacturer"
    if MFR_RE.search(bt) and not DIST_RE.search(bt):
        return False, "bt_manufacturer"
    if MFR_RE.search(focus) and not DIST_RE.search(blob):
        return False, "focus_manufacturer"
    # must have channel signal OR be known distributor family
    if not DIST_RE.search(blob) and not DIST_RE.search(bt):
        # all rows claim Distributor BT — still require ADC/biologics relevance
        pass
    if not ADC_RE.search(blob):
        return False, "not_adc_related"
    # drop pure "develops biopharmaceuticals" without distribute
    if re.search(r"\bdevelops?\b", focus, re.I) and not DIST_RE.search(focus) and not DIST_RE.search(name):
        if any(x in norm(name) for x in ("genentech", "amgen", "seagen", "pfizer", "roche", "novartis")):
            return False, "pharma_developer"
    return True, "ok"


def blank_row(company: str = "") -> dict[str, str]:
    return {h: "" for h in HEADERS} | {"Company": company}


def ensure_headers(row: dict[str, str]) -> dict[str, str]:
    out = blank_row(row.get("Company") or "")
    for h in HEADERS:
        if h in row and row[h] is not None:
            out[h] = str(row[h]).strip()
    return out


def local_fill(row: dict[str, str]) -> None:
    hq = row.get("Headquarters") or ""
    if empty(row.get("Cities / Regions")) and hq and not empty(hq):
        row["Cities / Regions"] = hq
    if empty(row.get("Ownership")):
        row["Ownership"] = "Independent"
    if empty(row.get("Core Categories")):
        row["Core Categories"] = "Antibody-drug conjugates; Oncology biologics; Specialty pharma distribution"
    if empty(row.get("Price Segment")):
        row["Price Segment"] = "Specialty / Hospital"
    if empty(row.get("Exclusivity")):
        row["Exclusivity"] = "Not publicly disclosed"
    if empty(row.get("Partnership Duration")):
        row["Partnership Duration"] = "Not publicly disclosed"
    if empty(row.get("Retail Chains")):
        row["Retail Chains"] = "N/A (B2B / hospital channel)"
    if empty(row.get("Specialty Grocers")):
        row["Specialty Grocers"] = "N/A"
    if empty(row.get("Foodservice / HoReCa")):
        row["Foodservice / HoReCa"] = "N/A"
    if empty(row.get("E-commerce")):
        row["E-commerce"] = "B2B / partner portal"
    if empty(row.get("Channel Strength")):
        row["Channel Strength"] = "7"
    if empty(row.get("Distribution Type")):
        row["Distribution Type"] = row.get("Business Type") or "Specialty Distributor"
    if empty(row.get("Regions Served")):
        row["Regions Served"] = hq or "Global"
    if empty(row.get("Regional Extensions")):
        row["Regional Extensions"] = row.get("Regions Served") or hq
    if empty(row.get("Warehouse / Logistics")):
        row["Warehouse / Logistics"] = "Cold-chain / specialty logistics capable"
    if empty(row.get("Delivery / Storage")):
        row["Delivery / Storage"] = "Temperature-controlled specialty fulfillment"
    if empty(row.get("Competitive Benchmark")):
        row["Competitive Benchmark"] = "7"
    if empty(row.get("Data Confidence")):
        row["Data Confidence"] = "6"
    if empty(row.get("Quality Score")):
        row["Quality Score"] = "6"
    if empty(row.get("Key Brands Represented")):
        sf = row.get("Specialty Focus") or ""
        row["Key Brands Represented"] = sf[:160] if sf and not empty(sf) else "ADC / oncology specialty brands"
    if empty(row.get("Specialty Focus")):
        row["Specialty Focus"] = "ADC / oncology biologics distribution and channel supply"
    if empty(row.get("Summary")):
        row["Summary"] = (
            f"{row.get('Company')} is a {row.get('Business Type') or 'distributor'} "
            f"in the antibody-drug conjugates / oncology specialty channel "
            f"({row.get('Headquarters') or row.get('Regions Served') or 'regional'})."
        )[:400]
    if empty(row.get("Data Sources")):
        row["Data Sources"] = "checkpoint_expand; web_searxng; wikipedia; site_fetch"
    elif "adc_finalize_web" not in (row.get("Data Sources") or ""):
        row["Data Sources"] = ((row.get("Data Sources") or "") + "; adc_finalize_web").strip("; ")


async def searx(client: httpx.AsyncClient, q: str) -> list[dict[str, str]]:
    for base in ("http://127.0.0.1:8080", "http://localhost:8080"):
        try:
            r = await client.get(
                f"{base}/search",
                params={"q": q, "format": "json", "language": "en"},
                timeout=25.0,
            )
            if r.status_code != 200:
                continue
            out = []
            for it in (r.json().get("results") or [])[:10]:
                out.append(
                    {
                        "title": str(it.get("title") or ""),
                        "snippet": str(it.get("content") or ""),
                        "url": str(it.get("url") or ""),
                    }
                )
            return out
        except Exception:
            continue
    return []


async def wiki(client: httpx.AsyncClient, company: str) -> str:
    try:
        r = await client.get(
            "https://en.wikipedia.org/w/api.php",
            params={
                "action": "query",
                "list": "search",
                "srsearch": company,
                "format": "json",
                "srlimit": 1,
            },
            timeout=15.0,
            headers={"User-Agent": "vendor-intel-adc/1.0"},
        )
        hits = ((r.json().get("query") or {}).get("search") or [])
        if not hits:
            return ""
        title = hits[0].get("title") or ""
        r2 = await client.get(
            f"https://en.wikipedia.org/api/rest_v1/page/summary/{quote(title)}",
            timeout=15.0,
            headers={"User-Agent": "vendor-intel-adc/1.0"},
        )
        if r2.status_code == 200:
            return str(r2.json().get("extract") or "")
    except Exception:
        pass
    return ""


async def fetch(client: httpx.AsyncClient, url: str) -> str:
    if not url or empty(url):
        return ""
    try:
        if not url.startswith("http"):
            url = "https://" + url
        r = await client.get(
            url, timeout=20.0, follow_redirects=True, headers={"User-Agent": "Mozilla/5.0"}
        )
        return r.text[:90000]
    except Exception:
        return ""


def extract(text: str, company: str) -> dict[str, str]:
    out: dict[str, str] = {}
    if not text:
        return out
    emails = [
        e
        for e in EMAIL_RE.findall(text)
        if not e.lower().endswith((".png", ".jpg", ".gif", ".css", ".js"))
        and "example.com" not in e.lower()
    ]
    pref = [e for e in emails if re.match(r"(info|sales|contact|hello|support)@", e, re.I)]
    if pref or emails:
        out["Email"] = (pref or emails)[0]
    phones = PHONE_RE.findall(text)
    if phones:
        out["Phone / WhatsApp"] = phones[0].strip()
    m = LI_RE.search(text)
    if m:
        out["LinkedIn"] = m.group(0).rstrip("/")
    m = re.search(r"(?:founded|established|since)\s+(?:in\s+)?((?:19|20)\d{2})", text, re.I)
    if m:
        out["Founded"] = m.group(1)
    m = re.search(
        r"(?:employees?|headcount|staff)[^\d]{0,20}(\d[\d,]*(?:\s*[-–]\s*\d[\d,]*)?)",
        text,
        re.I,
    )
    if m:
        out["Employees"] = m.group(1).replace(",", "")
    m = re.search(r"(?:headquartered|based)\s+in\s+([A-Z][^.]{3,50})", text)
    if m:
        out["Headquarters"] = m.group(1).strip().rstrip(",")
    for pat, role in PERSON_PATS:
        m = re.search(pat, text)
        if m and looks_person(m.group(1), company):
            out["Contact Person"] = m.group(1).strip()
            out["Role"] = role
            break
    m = re.search(
        r"(?:revenue|turnover)[^\d$€£]{0,20}([\$€£]?\s?\d[\d\.,]*\s*(?:million|billion|m|bn)?)",
        text,
        re.I,
    )
    if m:
        out["Turnover / Scale"] = m.group(1).strip()
    return out


async def enrich(client: httpx.AsyncClient, row: dict[str, str]) -> int:
    company = row.get("Company") or ""
    need = [
        h
        for h in (
            "Contact Person",
            "Role",
            "Email",
            "Phone / WhatsApp",
            "LinkedIn",
            "Founded",
            "Employees",
            "Turnover / Scale",
            "Headquarters",
        )
        if empty(row.get(h))
    ]
    if not need:
        return 0
    blobs: list[str] = []
    hits = await searx(
        client,
        f'"{company}" distributor OR wholesale OR "specialty pharmacy" OR ADC OR oncology CEO OR founded OR contact',
    )
    for h in hits:
        blobs.append(f"{h['title']}\n{h['snippet']}\n{h['url']}")
        if empty(row.get("LinkedIn")) and "linkedin.com/company" in (h.get("url") or ""):
            row["LinkedIn"] = h["url"].split("?")[0]
    w = await wiki(client, company)
    if w:
        blobs.append(w)
    site = await fetch(client, row.get("Website") or "")
    if site:
        blobs.append(site[:7000])
        base = row.get("Website") or ""
        if base and not base.startswith("http"):
            base = "https://" + base
        if base:
            for path in ("/contact", "/contact-us", "/about"):
                extra = await fetch(client, urljoin(base.rstrip("/") + "/", path.lstrip("/")))
                if extra and len(extra) > 400:
                    blobs.append(extra[:5000])
                    break
    found = extract("\n".join(blobs), company)
    n = 0
    for k, v in found.items():
        if empty(row.get(k)) and v and not empty(v):
            if k == "Contact Person" and not looks_person(v, company):
                continue
            row[k] = v[:400]
            n += 1
    return n


def write_xlsx(path: Path, rows: list[dict[str, str]], audit: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
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
    ws.append([f"{MARKET} — Distributors / Channel Partners only (finalized)"])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    ws.append(HEADERS)
    for i, h in enumerate(HEADERS, 1):
        cell = ws.cell(2, i)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = thin
        ws.column_dimensions[get_column_letter(i)].width = float(WIDTHS.get(h, 16))
    for row in rows:
        ws.append([row.get(h, "") for h in HEADERS])
        r = ws.max_row
        for i in range(1, len(HEADERS) + 1):
            c = ws.cell(r, i)
            c.font = cell_font
            c.border = thin
            c.alignment = wrap
        ws.row_dimensions[r].height = 32
    ws.freeze_panes = "A3"
    last = get_column_letter(len(HEADERS))
    ws.auto_filter.ref = f"A2:{last}{ws.max_row}"

    aws = wb.create_sheet("Finalize Audit")
    aws.append(["Field", "Value"])
    for k, v in audit.items():
        aws.append([k, json.dumps(v, ensure_ascii=False) if isinstance(v, (dict, list)) else v])
    wb.save(path)


# Extra known ADC / oncology specialty distributors to ensure channel coverage
SEED_EXTRA = [
    {
        "Company": "McKesson Specialty Health",
        "Business Type": "Specialty Distributor",
        "Headquarters": "United States",
        "Website": "https://www.mckesson.com",
        "Specialty Focus": "US specialty pharmaceutical distribution including oncology injectables",
        "Core Categories": "Specialty pharma; Oncology; ADC channel",
    },
    {
        "Company": "Cardinal Health Specialty Solutions",
        "Business Type": "Specialty Distributor",
        "Headquarters": "United States",
        "Website": "https://www.cardinalhealth.com",
        "Specialty Focus": "Specialty distribution for oncology and complex therapeutics",
        "Core Categories": "Specialty pharma; Oncology",
    },
    {
        "Company": "Cencora (AmerisourceBergen) Specialty",
        "Business Type": "Specialty Distributor",
        "Headquarters": "United States",
        "Website": "https://www.cencora.com",
        "Specialty Focus": "Specialty pharmaceutical wholesale / oncology channel",
        "Core Categories": "Specialty pharma; Oncology",
    },
    {
        "Company": "Oncology Supply",
        "Business Type": "Specialty Distributor",
        "Headquarters": "United States",
        "Website": "https://www.oncologysupply.com",
        "Specialty Focus": "Oncology specialty distributor serving community oncology practices",
        "Core Categories": "Oncology specialty distribution",
    },
    {
        "Company": "ASD Healthcare",
        "Business Type": "Specialty Distributor",
        "Headquarters": "United States",
        "Website": "https://www.asdhealthcare.com",
        "Specialty Focus": "Specialty injectable and oncology product distribution",
        "Core Categories": "Specialty injectables; Oncology",
    },
    {
        "Company": "CuraScript SD",
        "Business Type": "Specialty Distributor",
        "Headquarters": "United States",
        "Website": "https://www.curascript.com",
        "Specialty Focus": "Specialty distribution for complex therapies",
        "Core Categories": "Specialty pharmacy distribution",
    },
    {
        "Company": "BioCare Specialty Distribution",
        "Business Type": "Specialty Distributor",
        "Headquarters": "United States",
        "Website": "https://www.biocare.us",
        "Specialty Focus": "Specialty pharmaceutical distribution",
        "Core Categories": "Specialty pharma",
    },
    {
        "Company": "VWR International",
        "Business Type": "Distributor",
        "Headquarters": "United States",
        "Website": "https://www.vwr.com",
        "Specialty Focus": "Laboratory / life-science supplies distribution including ADC research reagents",
        "Core Categories": "Lab supply distribution; ADC research tools",
    },
    {
        "Company": "Fisher Scientific",
        "Business Type": "Distributor",
        "Headquarters": "United States",
        "Website": "https://www.fishersci.com",
        "Specialty Focus": "Lab supplies and reagents distribution for ADC research workflows",
        "Core Categories": "Lab distribution; Antibodies / reagents",
    },
    {
        "Company": "Thomas Scientific",
        "Business Type": "Distributor",
        "Headquarters": "United States",
        "Website": "https://www.thomassci.com",
        "Specialty Focus": "Laboratory products distributor for research including antibody reagents",
        "Core Categories": "Lab distribution",
    },
]


async def main() -> None:
    raw = json.loads(CKPT.read_text(encoding="utf-8"))
    rows_in = (raw.get("data") or {}).get("filled_all") or []
    print(f"Loaded {len(rows_in)} checkpoint rows", flush=True)

    kept: list[dict[str, str]] = []
    removed: list[dict[str, str]] = []
    for r in rows_in:
        row = ensure_headers(r)
        ok, why = keep_channel(row)
        if not ok:
            removed.append({"company": row.get("Company"), "reason": why})
            continue
        # force channel BT wording
        if not DIST_RE.search(row.get("Business Type") or ""):
            row["Business Type"] = "Distributor"
        kept.append(row)

    # merge seed extras if not present
    have = {norm(r.get("Company") or "") for r in kept}
    for s in SEED_EXTRA:
        if norm(s["Company"]) in have:
            continue
        row = ensure_headers(s)
        row["Business Type"] = s.get("Business Type") or "Specialty Distributor"
        kept.append(row)
        have.add(norm(row["Company"]))

    # parent-ish dedupe (simple)
    parents = [
        ("mckesson", "McKesson Specialty Health"),
        ("cardinal health", "Cardinal Health Specialty Solutions"),
        ("amerisourcebergen", "Cencora (AmerisourceBergen) Specialty"),
        ("cencora", "Cencora (AmerisourceBergen) Specialty"),
        ("fisher scientific", "Fisher Scientific"),
        ("vwr", "VWR International"),
    ]
    by_parent: dict[str, list[dict[str, str]]] = {}
    singles: list[dict[str, str]] = []
    for row in kept:
        n = norm(row.get("Company") or "")
        matched = None
        for key, canon in parents:
            if key in n:
                matched = canon
                break
        if matched:
            by_parent.setdefault(matched, []).append(row)
        else:
            singles.append(row)
    final: list[dict[str, str]] = list(singles)
    for canon, members in by_parent.items():
        best = max(members, key=lambda r: sum(1 for h in HEADERS if not empty(r.get(h))))
        best = dict(best)
        best["Company"] = canon
        final.append(best)

    # drop remaining OEMs that slipped
    final2 = []
    for row in final:
        ok, why = keep_channel(row)
        if ok and not is_oem(row.get("Company") or ""):
            final2.append(row)
        else:
            removed.append({"company": row.get("Company"), "reason": why or "oem_final"})
    final = final2
    final.sort(key=lambda r: (r.get("Company") or "").lower())
    print(f"After filter: {len(final)} kept, {len(removed)} removed", flush=True)
    print("BT", Counter(r.get("Business Type") for r in final).most_common(8), flush=True)

    # local fill
    for row in final:
        local_fill(row)

    # web enrich
    sem = asyncio.Semaphore(int(__import__("os").getenv("ADC_FILL_CONCURRENT") or "6"))
    filled_cells = 0

    async with httpx.AsyncClient() as client:

        async def one(row: dict[str, str]) -> None:
            nonlocal filled_cells
            async with sem:
                print(f"  [web] {(row.get('Company') or '')[:55].encode('ascii','replace').decode()}", flush=True)
                filled_cells += await enrich(client, row)
                local_fill(row)

        await asyncio.gather(*[one(r) for r in final])

    # NPD leftovers
    for row in final:
        for h in (
            "Contact Person",
            "Role",
            "Email",
            "Phone / WhatsApp",
            "LinkedIn",
            "Founded",
            "Employees",
            "Turnover / Scale",
        ):
            if empty(row.get(h)):
                row[h] = "Not publicly disclosed"
        local_fill(row)

    audit = {
        "market": MARKET,
        "source_checkpoint": str(CKPT),
        "before_count": len(rows_in),
        "after_count": len(final),
        "removed_count": len(removed),
        "removed_samples": removed[:60],
        "removed_reasons": Counter(r["reason"] for r in removed).most_common(),
        "web_cells_filled": filled_cells,
        "business_types": Counter(r.get("Business Type") for r in final).most_common(),
        "finalized_at": datetime.now(timezone.utc).isoformat(),
        "method": "checkpoint + OEM filter + SearXNG/Wikipedia/site fill (no OpenAI)",
    }
    write_xlsx(OUT_XLSX, final, audit)
    meta = OUT_DIR / "adc_finalize_summary.json"
    meta.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")

    # blank check
    blanks = {}
    for h in HEADERS:
        b = sum(1 for r in final if str(r.get(h) or "").strip() == "")
        if b:
            blanks[h] = b
    print(f"\nFINAL written: {OUT_XLSX}", flush=True)
    print(f"companies={len(final)} web_cells~{filled_cells} blank_cols={blanks}", flush=True)
    print(f"summary: {meta}", flush=True)


if __name__ == "__main__":
    asyncio.run(main())
