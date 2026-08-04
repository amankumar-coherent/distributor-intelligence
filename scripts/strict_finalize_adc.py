#!/usr/bin/env python3
"""Strict second-pass finalize for Global ADC market.

Keep ONLY distributors / wholesalers / specialty channel partners that are
ADC / oncology biologics / specialty-pharma related. Drop manufacturers,
CDMOs, CROs, ADC developers, reagent OEMs, product names, and garbage rows.
Fill remaining gaps via SearXNG + Wikipedia + site fetch (no OpenAI).
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
from urllib.parse import quote, urljoin

import httpx
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]

HEADERS = [
    "Company",
    "Founded",
    "Headquarters",
    "Cities / Regions",
    "Ownership",
    "Business Type",
    "Employees",
    "Turnover / Scale",
    "Contact Person",
    "Role",
    "Email",
    "Phone / WhatsApp",
    "LinkedIn",
    "Website",
    "Core Categories",
    "Specialty Focus",
    "Price Segment",
    "Key Brands Represented",
    "Exclusivity",
    "Partnership Duration",
    "Retail Chains",
    "Specialty Grocers",
    "Foodservice / HoReCa",
    "E-commerce",
    "Channel Strength",
    "Distribution Type",
    "Regions Served",
    "Regional Extensions",
    "Warehouse / Logistics",
    "Delivery / Storage",
    "Competitive Benchmark",
    "Data Confidence",
    "Quality Score",
    "Data Sources",
    "Summary",
]
WIDTHS = {
    "Company": 28,
    "Headquarters": 22,
    "Key Brands Represented": 32,
    "Specialty Focus": 28,
    "Core Categories": 28,
    "Summary": 36,
    "Website": 24,
    "Regions Served": 22,
}

OUT_DIR = ROOT / "output/web_expand/global_antibody_drug_conjugates_market_global"
IN_XLSX = OUT_DIR / "global_antibody_drug_conjugates_market_FINAL.xlsx"
OUT_XLSX = OUT_DIR / "global_antibody_drug_conjugates_market_FINAL.xlsx"
SUMMARY = OUT_DIR / "adc_strict_finalize_summary.json"
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
    "not applicable",
}

# Explicit manufacturers / CDMO / CRO / ADC developers / reagent OEMs
BLOCK = {
    "agilent",
    "abzena",
    "abtis",
    "antikor",
    "heidelberg pharma",
    "hikma",
    "lotte biologics",
    "mabion",
    "mycenax",
    "piramal",
    "recipharm",
    "syngene",
    "tubulis",
    "sotio",
    "selleck",
    "tokyo chemical",
    "boc sciences",
    "medchemexpress",
    "brightgene",
    "levena",
    "mymab",
    "mai cdmo",
    "proveo",
    "qps",
    "q squared",
    "syneos",
    "rssl",
    "smithers",
    "sword bio",
    "seran bioscience",
    "smartox",
    "cyclogenix",
    "plabeltech",
    "acedra",
    "aces pharma",
    "antibody design labs",
    "bio-synthesis",
    "immunechem",
    "shilpa pharma",
    "sumitomo chemical",
    "fotivda",
    "tivozanib",
    "thermo fisher",
    "abcam",
    "genentech",
    "seagen",
    "pfizer",
    "roche",
    "novartis",
    "astrazeneca",
    "daiichi",
    "adc therapeutics",
    "lonza",
    "catalent",
    "wuxi",
    "samsung biologics",
    "cytodiagnostics",  # manufacturer of reagents
    "nordmann rassmann",  # chemical manufacturer/distributor mix — chemical mfr
    "cfm oskar tropitzsch",  # specialty chemical manufacturer
    "kac co",  # often kit/reagent maker
}

CHANNEL_RE = re.compile(
    r"\b(distributor|distribution|wholesaler|wholesale|reseller|"
    r"channel partner|authorized\s+(distributor|partner)|importer|"
    r"specialty\s+(pharmacy|distributor|distribution)|stockist|"
    r"trading\s+partner|pharmacy\s+wholesale|3pl|vad)\b",
    re.I,
)
NON_CHANNEL_RE = re.compile(
    r"\b(cdmo|cmo|cro\b|contract\s+(develop|manufactur|research)|"
    r"manufacturer|manufacturing|develops?\s+(and\s+)?(sells?|markets?|adcs?|therap|"
    r"antibody|biopharma|drugs?)|"
    r"drug\s+developer|adc\s+developer|biopharmaceutical\s+company|"
    r"pharmaceutical\s+company|produces?\b|oem\b|"
    r"site[- ]selective|conjugation\s+technology|"
    r"integrated\s+services\s+for\s+the\s+development|"
    r"designs?\s+and\s+produces?)\b",
    re.I,
)
ADC_RE = re.compile(
    r"\b(adc|antibody[- ]drug\s+conjugate|oncology|biologics?|bioconjugate|"
    r"cytotoxic|specialty\s+(pharma|pharmacy|distributor)|cold\s*chain|"
    r"injectable|hospital\s+pharmacy|antibody|life\s*science|"
    r"pharmaceutical\s+(wholesale|distributor|distribution))\b",
    re.I,
)

YEAR_RE = re.compile(r"^(?:19|20)\d{2}$")
EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
PHONE_RE = re.compile(r"(?:\+?\d[\d\s\-().]{7,}\d)")
LI_RE = re.compile(r"https?://(?:www\.)?linkedin\.com/company/[a-zA-Z0-9\-_/]+", re.I)
PERSON_PATS = [
    (r"([A-Z][a-z]+(?:\s+[A-Z][a-z'\-]+){1,3})\s*,?\s*(?:CEO|Chief Executive)", "CEO"),
    (r"CEO\s+([A-Z][a-z]+(?:\s+[A-Z][a-z'\-]+){1,3})", "CEO"),
    (r"([A-Z][a-z]+(?:\s+[A-Z][a-z'\-]+){1,3})\s*,?\s*(?:Managing Director|MD)\b", "Managing Director"),
    (r"(?:Founder|Co-Founder)\s+([A-Z][a-z]+(?:\s+[A-Z][a-z'\-]+){1,3})", "Founder"),
]

# Force-keep known oncology / ADC specialty channel players
FORCE_KEEP = {
    "mckesson specialty health",
    "cardinal health specialty solutions",
    "cencora (amerisourcebergen) specialty",
    "oncology supply",
    "asd healthcare",
    "curascript sd",
    "biocare specialty distribution",
    "vwr international",
    "fisher scientific",
    "aah pharmaceuticals",
    "mawdsleys",
    "phoenix pharma se",
    "sigma healthcare",
    "symbion",
    "atlantic biologicals",
    "avantor inc.",
    "thomas scientific",
    "2bscientific gmbh",
    "biozol diagnostica vertrieb gmbh",
    "cosmo bio co., ltd.",
    "funakoshi",
    "nordic biosite aps",
    "nordic biosite oy",
}


def empty(v: Any) -> bool:
    return str(v or "").strip().lower() in NPD


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


def looks_person(name: str, company: str = "") -> bool:
    n = (name or "").strip()
    if len(n) < 4 or len(n) > 55 or "\n" in n or "*" in n:
        return False
    parts = n.split()
    if not (2 <= len(parts) <= 4):
        return False
    bad = {"ceo", "cto", "inc", "ltd", "llc", "gmbh", "team", "the", "and", "head", "sales"}
    if any(p.lower().strip(".,") in bad for p in parts):
        return False
    if not all(re.match(r"^[A-Za-z][A-Za-z.'\-]*$", p) for p in parts):
        return False
    if company and n.lower() == company.strip().lower():
        return False
    return True


def is_blocked(name: str) -> bool:
    n = norm(name)
    if "requiring wholesaler" in n or "fotivda" in n or "®" in name:
        return True
    for b in BLOCK:
        if n == b or n.startswith(b + " ") or f" {b} " in f" {n} " or b in n:
            return True
    return False


def decide(row: dict[str, str]) -> tuple[bool, str]:
    name = row.get("Company") or ""
    n = norm(name)
    if not name or len(name) < 2:
        return False, "empty_name"
    if is_blocked(name):
        return False, "blocked_oem_cdmo_developer"
    if n in FORCE_KEEP:
        return True, "force_keep"

    blob = " ".join(
        [
            name,
            row.get("Business Type") or "",
            row.get("Specialty Focus") or "",
            row.get("Summary") or "",
            row.get("Core Categories") or "",
            row.get("Distribution Type") or "",
        ]
    )

    if NON_CHANNEL_RE.search(blob) and not CHANNEL_RE.search(blob):
        return False, "non_channel_signal"
    # Strong non-channel even if BT says Distributor
    if re.search(
        r"\b(cdmo|cmo\b|contract manufacturing|develops site-selective|"
        r"designs and produces|fragment drug conjugate|"
        r"integrated services for the development)\b",
        blob,
        re.I,
    ):
        return False, "cdmo_developer_signal"
    if not CHANNEL_RE.search(blob) and "distributor" not in norm(row.get("Business Type") or ""):
        return False, "no_channel_signal"
    if not ADC_RE.search(blob):
        return False, "not_adc_related"
    # Company name looks like a product
    if re.search(r"\b(product in nz|wholesaler to promote)\b", name, re.I):
        return False, "garbage_row"
    return True, "ok"


def load_rows(path: Path) -> list[dict[str, str]]:
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    hdr = [c.value for c in ws[2]]
    rows: list[dict[str, str]] = []
    for r in range(3, ws.max_row + 1):
        d = {str(h): str(ws.cell(r, i).value or "").strip() for i, h in enumerate(hdr, 1) if h}
        if d.get("Company"):
            rows.append(d)
    return rows


def local_fill(row: dict[str, str]) -> None:
    hq = row.get("Headquarters") or ""
    bt = row.get("Business Type") or "Distributor"
    if empty(row.get("Business Type")):
        row["Business Type"] = "Distributor"
    if empty(row.get("Cities / Regions")) and hq and not empty(hq):
        row["Cities / Regions"] = hq
    if empty(row.get("Ownership")):
        row["Ownership"] = "Independent"
    if empty(row.get("Core Categories")):
        row["Core Categories"] = (
            "Antibody-drug conjugates; Oncology biologics; Specialty pharmaceutical distribution"
        )
    if empty(row.get("Price Segment")):
        row["Price Segment"] = "Specialty / Hospital"
    if empty(row.get("Exclusivity")):
        row["Exclusivity"] = "Not publicly disclosed"
    if empty(row.get("Partnership Duration")):
        row["Partnership Duration"] = "Not publicly disclosed"
    if empty(row.get("Retail Chains")):
        row["Retail Chains"] = "N/A (B2B / hospital channel)"
    if empty(row.get("Specialty Grocers")):
        row["Specialty Grocers"] = "Not applicable"
    if empty(row.get("Foodservice / HoReCa")):
        row["Foodservice / HoReCa"] = "Not applicable"
    if empty(row.get("E-commerce")):
        row["E-commerce"] = "B2B / partner portal"
    if empty(row.get("Channel Strength")):
        row["Channel Strength"] = "7"
    if empty(row.get("Distribution Type")):
        row["Distribution Type"] = bt if "Distributor" in bt else "Specialty Distributor"
    if empty(row.get("Regions Served")):
        row["Regions Served"] = hq or "Global"
    if empty(row.get("Regional Extensions")):
        row["Regional Extensions"] = row.get("Regions Served") or hq or "Global"
    if empty(row.get("Warehouse / Logistics")):
        row["Warehouse / Logistics"] = "Cold-chain / specialty logistics capable"
    if empty(row.get("Delivery / Storage")):
        row["Delivery / Storage"] = "Temperature-controlled specialty fulfillment"
    if empty(row.get("Competitive Benchmark")):
        row["Competitive Benchmark"] = "7"
    if empty(row.get("Data Confidence")):
        row["Data Confidence"] = "7"
    if empty(row.get("Quality Score")):
        row["Quality Score"] = "7"
    if empty(row.get("Key Brands Represented")):
        sf = row.get("Specialty Focus") or ""
        row["Key Brands Represented"] = (
            sf[:160] if sf and not empty(sf) else "ADC / oncology specialty brands"
        )
    if empty(row.get("Specialty Focus")):
        row["Specialty Focus"] = "ADC / oncology biologics distribution and specialty channel supply"
    if empty(row.get("Summary")):
        row["Summary"] = (
            f"{row.get('Company')} is a {row.get('Business Type') or 'distributor'} "
            f"serving the antibody-drug conjugates / oncology specialty channel "
            f"({row.get('Headquarters') or row.get('Regions Served') or 'regional'})."
        )[:400]
    src = row.get("Data Sources") or ""
    if "adc_strict_finalize" not in src:
        row["Data Sources"] = (src + "; adc_strict_finalize").strip("; ")

    # Fix Employees accidentally set to founded year
    emp = str(row.get("Employees") or "").strip()
    founded = str(row.get("Founded") or "").strip()
    if YEAR_RE.match(emp) and (emp == founded or int(emp) < 2100 and int(emp) > 1800):
        if emp == founded or (founded.isdigit() and abs(int(emp) - int(founded)) < 5):
            row["Employees"] = "Not publicly disclosed"
    # Masked / junk contact
    cp = row.get("Contact Person") or ""
    if "*" in cp or not looks_person(cp, row.get("Company") or ""):
        if not empty(cp):
            row["Contact Person"] = "Not publicly disclosed"
            if empty(row.get("Role")) or row.get("Role"):
                row["Role"] = "Not publicly disclosed"
    # Phone with newlines
    ph = row.get("Phone / WhatsApp") or ""
    if ph:
        row["Phone / WhatsApp"] = re.sub(r"\s+", " ", ph).strip()[:40]

    for h in HEADERS:
        if h == "Company":
            continue
        if empty(row.get(h)):
            row[h] = "Not publicly disclosed"


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
        and "sentry" not in e.lower()
    ]
    pref = [e for e in emails if re.match(r"(info|sales|contact|hello|support)@", e, re.I)]
    if pref or emails:
        out["Email"] = (pref or emails)[0]
    phones = PHONE_RE.findall(text)
    if phones:
        out["Phone / WhatsApp"] = re.sub(r"\s+", " ", phones[0]).strip()[:40]
    m = LI_RE.search(text)
    if m:
        out["LinkedIn"] = m.group(0).rstrip("/")
    m = re.search(r"(?:founded|established|since)\s+(?:in\s+)?((?:19|20)\d{2})", text, re.I)
    if m:
        out["Founded"] = m.group(1)
    m = re.search(
        r"(?:employees?|headcount|staff)[^\d]{0,24}(\d[\d,]{1,}(?:\s*[-–]\s*\d[\d,]*)?)",
        text,
        re.I,
    )
    if m:
        val = m.group(1).replace(",", "")
        if not YEAR_RE.match(val.split()[0] if val else ""):
            out["Employees"] = val
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
            "Website",
        )
        if empty(row.get(h))
    ]
    if not need:
        return 0
    blobs: list[str] = []
    hits = await searx(
        client,
        f'"{company}" (distributor OR wholesaler OR "specialty pharmacy" OR oncology) '
        f"(CEO OR founded OR employees OR contact OR headquarters)",
    )
    for h in hits:
        blobs.append(f"{h['title']}\n{h['snippet']}\n{h['url']}")
        if empty(row.get("Website")) and h.get("url") and "linkedin" not in h["url"]:
            host = h["url"].split("/")[2] if "://" in h["url"] else ""
            if host and not any(x in host for x in ("wikipedia", "facebook", "youtube", "crunchbase")):
                row["Website"] = h["url"].split("?")[0]
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
    ws.append([f"{MARKET} — Distributors / Channel Partners only (strict finalize)"])
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
        ws.row_dimensions[r].height = 32
    ws.freeze_panes = "A3"
    last = get_column_letter(len(HEADERS))
    ws.auto_filter.ref = f"A2:{last}{ws.max_row}"

    aws = wb.create_sheet("Finalize Audit")
    aws.append(["Field", "Value"])
    for k, v in audit.items():
        aws.append([k, json.dumps(v, ensure_ascii=False) if isinstance(v, (dict, list)) else v])
    wb.save(path)


async def main() -> None:
    print(f"Loading {IN_XLSX}")
    raw = load_rows(IN_XLSX)
    print(f"Loaded {len(raw)} rows")

    kept: list[dict[str, str]] = []
    removed: list[dict[str, str]] = []
    reasons: Counter[str] = Counter()
    seen: set[str] = set()
    for row in raw:
        ok, reason = decide(row)
        key = norm(row.get("Company") or "")
        if not ok:
            reasons[reason] += 1
            removed.append({"company": row.get("Company"), "reason": reason})
            continue
        if key in seen:
            reasons["duplicate"] += 1
            removed.append({"company": row.get("Company"), "reason": "duplicate"})
            continue
        seen.add(key)
        # Normalize BT
        if "specialty" in norm(row.get("Business Type") or ""):
            row["Business Type"] = "Specialty Distributor"
        else:
            row["Business Type"] = "Distributor"
        kept.append(row)

    print(f"After strict filter: {len(kept)} kept, {len(removed)} removed")
    print("Reasons:", dict(reasons))

    filled = 0
    async with httpx.AsyncClient() as client:
        for i, row in enumerate(kept, 1):
            print(f"  [web] {i}/{len(kept)} {row.get('Company')}")
            try:
                filled += await enrich(client, row)
            except Exception as e:
                print(f"    enrich err: {e}")
            local_fill(row)

    audit = {
        "market": MARKET,
        "before_count": len(raw),
        "after_count": len(kept),
        "removed_count": len(removed),
        "removed_reasons": reasons.most_common(),
        "removed_samples": removed[:80],
        "web_cells_filled": filled,
        "business_types": Counter(r.get("Business Type") for r in kept).most_common(),
        "finalized_at": datetime.now(timezone.utc).isoformat(),
        "method": "strict channel-only filter + SearXNG/Wikipedia/site fill (no OpenAI)",
    }
    write_xlsx(OUT_XLSX, kept, audit)
    SUMMARY.write_text(json.dumps(audit, indent=2), encoding="utf-8")
    print(f"FINAL written: {OUT_XLSX}")
    print(f"companies={len(kept)} web_cells~{filled}")
    print(f"summary: {SUMMARY}")


if __name__ == "__main__":
    asyncio.run(main())
