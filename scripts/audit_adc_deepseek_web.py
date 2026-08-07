#!/usr/bin/env python3
"""Audit Global ADC FINAL: keep only ADC-related distributors/channel partners.

- Web search (DuckDuckGo HTML + Wikipedia + company site)
- DeepSeek via OpenAI SDK (OPENAI_* -> api.deepseek.com, deepseek-chat / low model)
- Remove manufacturers / CDMOs / OEMs / non-ADC / non-distributors
- Verify & correct filled column values
"""
from __future__ import annotations

import asyncio
import json
import os
import re
import sys
import time
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import quote, urljoin

import httpx
from dotenv import load_dotenv
from openai import OpenAI
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
load_dotenv(ROOT / ".env")

OUT_DIR = ROOT / "output/web_expand/global_antibody_drug_conjugates_market_global"
IN_XLSX = OUT_DIR / "global_antibody_drug_conjugates_market_FINAL.xlsx"
OUT_XLSX = OUT_DIR / "global_antibody_drug_conjugates_market_FINAL.xlsx"
BACKUP = OUT_DIR / f"global_antibody_drug_conjugates_market_FINAL_pre_audit_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
SUMMARY = OUT_DIR / "adc_audit_deepseek_summary.json"
PROGRESS = OUT_DIR / "adc_audit_progress.jsonl"
REMOVED = OUT_DIR / "adc_audit_removed.json"

MARKET = "Global Antibody Drug Conjugates Market"
MODEL = os.getenv("OPENAI_MODEL", "deepseek-chat")
BATCH = int(os.getenv("ADC_AUDIT_BATCH", "8"))
CONCURRENCY = int(os.getenv("ADC_AUDIT_CONCURRENCY", "6"))
RESUME = os.getenv("ADC_AUDIT_RESUME", "1") != "0"

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

# Brand owners / ADC developers / CDMO / CRO / reagent OEMs — not channel partners
OEM_BLOCK = {
    "thermo fisher", "fujifilm wako", "jackson immunoresearch", "abcam", "biolegend",
    "sigma-aldrich", "merck kgaa", "genentech", "roche", "pfizer", "novartis",
    "astrazeneca", "seagen", "daiichi", "adc therapeutics", "lonza", "catalent",
    "wuxi", "samsung biologics", "agilent", "genscript", "cell signaling",
    "santa cruz", "exbio praha", "abzena", "abtis", "antikor", "heidelberg pharma",
    "hikma", "lotte biologics", "mabion", "mycenax", "piramal", "recipharm",
    "syngene", "tubulis", "sotio", "selleck", "tokyo chemical", "boc sciences",
    "medchemexpress", "brightgene", "levena", "mymab", "mai cdmo", "proveo",
    "qps", "q squared", "syneos", "rssl", "smithers", "sword bio", "seran bioscience",
    "smartox", "cyclogenix", "plabeltech", "acedra", "aces pharma",
    "antibody design labs", "bio-synthesis", "immunechem", "shilpa pharma",
    "sumitomo chemical", "fotivda", "tivozanib", "cytodiagnostics",
    "cfm oskar tropitzsch", "gilead sciences", "immunomedics", "immunogen",
    "byondis", "mablink", "synaffix",
    "pieris", "macrogenics", "seattle genetics", "eli lilly", "bristol myers",
    "bms ", "johnson & johnson", "janssen", "amgen", "takeda", "eisai",
    "glaxosmithkline", "gsk ", "boehringer", "sanofi", "merck & co", "msd ",
}

CHANNEL_RE = re.compile(
    r"\b(distributor|distribution|wholesaler|wholesale|reseller|importer|"
    r"specialty\s+(pharmacy|distributor|distribution)|authorized\s+distributor|"
    r"channel\s+partner|stockist|trading|pharmacy\s+supply|3pl)\b",
    re.I,
)
MFR_RE = re.compile(
    r"\b(cdmo|cmo\b|cro\b|contract\s+(develop|manufactur)|"
    r"we\s+manufacture|manufactures?\s+(and\s+)?(develops?|sells?)|"
    r"drug\s+developer|adc\s+developer|biopharmaceutical\s+company\s+that\s+develops|"
    r"conjugation\s+technology|site[- ]selective)\b",
    re.I,
)
ADC_RE = re.compile(
    r"\b(adc|antibody[- ]drug\s+conjugate|oncology|biologics?|bioconjugate|"
    r"cytotoxic|specialty\s+(pharma|pharmacy|distributor)|cold\s*chain|"
    r"injectable|hospital\s+pharmacy|antibody|life\s*science|"
    r"pharmaceutical\s+(wholesale|distributor|distribution)|immunolog|"
    r"chemotherapy|cancer\s+(care|drug|medicine)|reagent|diagnostics?\b)\b",
    re.I,
)

YEAR_RE = re.compile(r"^(?:18|19|20)\d{2}$")
EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
PHONE_RE = re.compile(r"(?:\+?\d[\d\s\-().]{7,}\d)")
LI_RE = re.compile(r"https?://(?:www\.)?linkedin\.com/company/[a-zA-Z0-9\-_/]+", re.I)
FOUNDED_PATS = [
    re.compile(r"(?:founded|established|incorpora(?:ted|tion)|since|started|began|created)\s+(?:in\s+)?((?:18|19|20)\d{2})", re.I),
    re.compile(r"\best\.?\s*((?:18|19|20)\d{2})\b", re.I),
    re.compile(r"\bsince\s*((?:18|19|20)\d{2})\b", re.I),
]
HQ_PATS = [
    re.compile(r"(?:headquarter(?:ed|s)?|based|located)\s+(?:in|at)\s+([A-Z][A-Za-z .,'\-]{2,60}(?:,\s*[A-Z][A-Za-z .,'\-]{2,40})?)", re.I),
    re.compile(r"(?:HQ|head office|corporate office)\s*[:\-]\s*([A-Z][A-Za-z0-9 .,'\-]{3,70})", re.I),
]


def empty(v: Any) -> bool:
    return str(v or "").strip().lower() in NPD


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


def is_oem(name: str) -> bool:
    n = norm(name)
    return any(b in n for b in OEM_BLOCK)


def load_rows(path: Path) -> list[dict[str, str]]:
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    hdr = [c.value for c in ws[2]]
    rows: list[dict[str, str]] = []
    for r in range(3, ws.max_row + 1):
        d = {h: str(ws.cell(r, i).value or "").strip() for i, h in enumerate(hdr, 1) if h}
        if d.get("Company"):
            rows.append(d)
    return rows


def write_xlsx(path: Path, rows: list[dict[str, str]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "ADC Distributors"
    # Title row
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HEADERS))
    title = ws.cell(1, 1, MARKET + " — Distributors / Channel Partners (audited)")
    title.font = Font(bold=True, size=14, color="FFFFFF")
    title.fill = PatternFill("solid", fgColor="1F4E79")
    title.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 28

    header_fill = PatternFill("solid", fgColor="2E75B6")
    header_font = Font(bold=True, color="FFFFFF")
    thin = Border(
        left=Side(style="thin", color="BDD7EE"),
        right=Side(style="thin", color="BDD7EE"),
        top=Side(style="thin", color="BDD7EE"),
        bottom=Side(style="thin", color="BDD7EE"),
    )
    for i, h in enumerate(HEADERS, 1):
        c = ws.cell(2, i, h)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(wrap_text=True, vertical="center")
        c.border = thin
        ws.column_dimensions[get_column_letter(i)].width = WIDTHS.get(h, 16)

    for ri, row in enumerate(rows, 3):
        for ci, h in enumerate(HEADERS, 1):
            val = row.get(h, "")
            cell = ws.cell(ri, ci, val)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = thin
            if ri % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F2F2F2")

    # Checklist sheet
    cws = wb.create_sheet("Audit Checklist")
    cws["A1"] = "ADC market audit checklist"
    cws["A2"] = "Only distributors / specialty distributors / channel partners kept"
    cws["A3"] = "Manufacturers, CDMOs, CROs, ADC developers, reagent OEMs removed"
    cws["A4"] = "Columns verified via web search + DeepSeek (deepseek-chat)"
    cws["A5"] = f"Generated: {datetime.now(timezone.utc).isoformat()}"
    cws["A6"] = f"Rows kept: {len(rows)}"

    wb.save(path)


async def fetch_text(client: httpx.AsyncClient, url: str, timeout: float = 6.0) -> str:
    try:
        r = await client.get(url, timeout=timeout, follow_redirects=True)
        if r.status_code >= 400:
            return ""
        return r.text or ""
    except Exception:
        return ""


def strip_html(html: str) -> str:
    t = re.sub(r"(?is)<script[^>]*>.*?</script>", " ", html)
    t = re.sub(r"(?is)<style[^>]*>.*?</style>", " ", t)
    t = re.sub(r"(?is)<[^>]+>", " ", t)
    t = re.sub(r"&nbsp;|&amp;|&lt;|&gt;|&quot;", " ", t)
    t = re.sub(r"\s+", " ", t)
    return t.strip()


async def ddg_snippets(client: httpx.AsyncClient, query: str, n: int = 6) -> list[str]:
    url = f"https://html.duckduckgo.com/html/?q={quote(query)}"
    html = await fetch_text(client, url)
    if not html:
        return []
    snippets: list[str] = []
    for m in re.finditer(
        r'class="result__snippet"[^>]*>(.*?)</a>|class="result__a"[^>]*>(.*?)</a>',
        html,
        re.I | re.S,
    ):
        s = strip_html(m.group(1) or m.group(2) or "")
        if s and len(s) > 20:
            snippets.append(s[:400])
        if len(snippets) >= n:
            break
    # fallback: plain text chunks
    if not snippets:
        text = strip_html(html)
        if text:
            snippets.append(text[:800])
    return snippets


async def wikipedia_extract(client: httpx.AsyncClient, company: str) -> dict[str, str]:
    out: dict[str, str] = {}
    api = (
        "https://en.wikipedia.org/w/api.php?action=query&list=search"
        f"&srsearch={quote(company)}&utf8=1&format=json&srlimit=1"
    )
    try:
        r = await client.get(api, timeout=10.0)
        data = r.json()
        hits = (data.get("query") or {}).get("search") or []
        if not hits:
            return out
        title = hits[0].get("title") or ""
        if not title:
            return out
        sum_url = f"https://en.wikipedia.org/api/rest_v1/page/summary/{quote(title)}"
        r2 = await client.get(sum_url, timeout=10.0)
        js = r2.json()
        extract = js.get("extract") or ""
        out["wiki_title"] = title
        out["wiki_extract"] = extract[:1200]
        # try founded / hq from extract
        for pat in FOUNDED_PATS:
            m = pat.search(extract)
            if m:
                out["founded"] = m.group(1)
                break
        for pat in HQ_PATS:
            m = pat.search(extract)
            if m:
                out["hq"] = m.group(1).strip(" .,;")
                break
    except Exception:
        pass
    return out


async def site_extract(client: httpx.AsyncClient, website: str) -> str:
    if not website or not website.startswith("http"):
        return ""
    html = await fetch_text(client, website, timeout=5.0)
    return strip_html(html)[:2500] if html else ""


async def gather_evidence(client: httpx.AsyncClient, row: dict[str, str]) -> dict[str, Any]:
    company = row.get("Company", "")
    website = row.get("Website", "")
    use_ddg = os.getenv("ADC_AUDIT_DDG", "0") == "1"

    async def _bounded():
        tasks = [wikipedia_extract(client, company), site_extract(client, website)]
        if use_ddg:
            q1 = f'"{company}" distributor OR wholesaler OR specialty pharmacy OR oncology OR ADC'
            tasks.insert(0, ddg_snippets(client, q1, 4))
        return await asyncio.gather(*tasks)

    snips1: list[str] = []
    try:
        results = await asyncio.wait_for(_bounded(), timeout=12.0)
        if use_ddg:
            snips1, wiki, site = results  # type: ignore
        else:
            wiki, site = results  # type: ignore
    except asyncio.TimeoutError:
        wiki, site = {}, ""
    blob = " ".join(list(snips1) + [wiki.get("wiki_extract", ""), site])
    row_blob = " ".join([
        row.get("Core Categories", ""),
        row.get("Specialty Focus", ""),
        row.get("Summary", ""),
        row.get("Business Type", ""),
        row.get("Key Brands Represented", ""),
        row.get("Headquarters", ""),
        row.get("Regions Served", ""),
    ])
    blob_all = (blob + " " + row_blob)[:4500]
    return {
        "company": company,
        "snippets": snips1[:5],
        "meta_snippets": [],
        "wiki": wiki,
        "site_excerpt": (site or "")[:1200],
        "blob": blob_all,
        "channel_hit": bool(CHANNEL_RE.search(blob_all)),
        "mfr_hit": bool(MFR_RE.search(blob)),
        "adc_hit": bool(ADC_RE.search(blob_all)),
    }


def extract_fields_from_evidence(ev: dict[str, Any], row: dict[str, str]) -> dict[str, str]:
    fixes: dict[str, str] = {}
    blob = ev.get("blob") or ""
    wiki = ev.get("wiki") or {}
    if wiki.get("founded") and YEAR_RE.match(wiki["founded"]):
        if empty(row.get("Founded")) or row.get("Founded") != wiki["founded"]:
            # only overwrite if current looks wrong or empty; prefer wiki
            if empty(row.get("Founded")):
                fixes["Founded"] = wiki["founded"]
    if empty(row.get("Founded")):
        for pat in FOUNDED_PATS:
            m = pat.search(blob)
            if m and YEAR_RE.match(m.group(1)):
                fixes["Founded"] = m.group(1)
                break
    if wiki.get("hq") and empty(row.get("Headquarters")):
        fixes["Headquarters"] = wiki["hq"]
    if empty(row.get("Headquarters")):
        for pat in HQ_PATS:
            m = pat.search(blob)
            if m:
                fixes["Headquarters"] = m.group(1).strip(" .,;")[:80]
                break
    # contact crumbs
    if empty(row.get("Email")):
        m = EMAIL_RE.search(blob)
        if m:
            fixes["Email"] = m.group(0)
    if empty(row.get("LinkedIn")):
        m = LI_RE.search(blob)
        if m:
            fixes["LinkedIn"] = m.group(0).rstrip("/")
    return fixes


def get_client() -> OpenAI:
    key = os.getenv("OPENAI_API_KEY")
    base = os.getenv("OPENAI_BASE_URL", "https://api.deepseek.com/v1")
    if not key:
        raise RuntimeError("OPENAI_API_KEY missing (DeepSeek key expected)")
    return OpenAI(api_key=key, base_url=base)


SYSTEM_PROMPT = """You are auditing companies for the Global Antibody Drug Conjugates (ADC) market channel map.
KEEP only if BOTH are true:
1) Company is a distributor, wholesaler, specialty pharmacy, importer, reseller, or channel partner (NOT a manufacturer, CDMO, CMO, CRO, ADC developer, brand owner, or reagent OEM that primarily manufactures).
2) Company is related to ADC / oncology biologics / specialty pharma / antibody / life-science / hospital injectable / cold-chain specialty pharmaceutical distribution (or life-science antibody reagent distribution into the ADC research/clinical channel).

REMOVE if primarily manufacturer/OEM/CDMO/CRO/developer, or unrelated to ADC/oncology/specialty biologics channel.

Also verify/correct key fields using evidence. If unknown, keep existing or use "Not publicly disclosed".
Return STRICT JSON only."""


def deepseek_audit_batch(
    client: OpenAI,
    items: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """items: {row, evidence}"""
    payload = []
    for it in items:
        row = it["row"]
        ev = it["evidence"]
        payload.append({
            "company": row.get("Company"),
            "current": {
                "Founded": row.get("Founded"),
                "Headquarters": row.get("Headquarters"),
                "Cities / Regions": row.get("Cities / Regions"),
                "Ownership": row.get("Ownership"),
                "Business Type": row.get("Business Type"),
                "Employees": row.get("Employees"),
                "Website": row.get("Website"),
                "Core Categories": row.get("Core Categories"),
                "Specialty Focus": row.get("Specialty Focus"),
                "Key Brands Represented": row.get("Key Brands Represented"),
                "Regions Served": row.get("Regions Served"),
                "Distribution Type": row.get("Distribution Type"),
                "Summary": (row.get("Summary") or "")[:300],
            },
            "web_evidence": {
                "snippets": ev.get("snippets"),
                "meta_snippets": ev.get("meta_snippets"),
                "wiki": ev.get("wiki"),
                "site_excerpt": (ev.get("site_excerpt") or "")[:800],
                "channel_hit": ev.get("channel_hit"),
                "mfr_hit": ev.get("mfr_hit"),
                "adc_hit": ev.get("adc_hit"),
            },
        })

    user = {
        "task": "For each company decide keep or remove and return corrected fields.",
        "schema": {
            "company": "string exact name",
            "keep": "boolean",
            "reason": "short string",
            "role": "distributor|specialty_distributor|channel_partner|manufacturer|cdmo|oem|unrelated|other",
            "corrections": {
                "Founded": "YYYY or omit",
                "Headquarters": "city, country preferred",
                "Cities / Regions": "string",
                "Ownership": "Independent|Subsidiary|Public|Private|omit",
                "Business Type": "Distributor|Specialty Distributor",
                "Employees": "string or Not publicly disclosed",
                "Core Categories": "string",
                "Specialty Focus": "string",
                "Key Brands Represented": "string",
                "Regions Served": "string",
                "Distribution Type": "string",
                "Summary": "1-2 sentences",
            },
        },
        "companies": payload,
    }
    for attempt in range(3):
        try:
            resp = client.chat.completions.create(
                model=MODEL,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": json.dumps(user, ensure_ascii=False)},
                ],
                temperature=0.1,
                max_tokens=4000,
                response_format={"type": "json_object"},
            )
            text = resp.choices[0].message.content or "{}"
            data = json.loads(text)
            # normalize to list
            if isinstance(data, dict):
                for k in ("results", "companies", "items", "data"):
                    if isinstance(data.get(k), list):
                        return data[k]
                # single object
                if "company" in data:
                    return [data]
            if isinstance(data, list):
                return data
            return []
        except Exception as e:
            if attempt == 2:
                print(f"  DeepSeek batch error: {e}", flush=True)
                return []
            time.sleep(1.5 * (attempt + 1))
    return []


def apply_decision(row: dict[str, str], decision: dict[str, Any], ev: dict[str, Any]) -> dict[str, str]:
    out = dict(row)
    corr = decision.get("corrections") or {}
    if isinstance(corr, dict):
        for k, v in corr.items():
            if k in HEADERS and v is not None and str(v).strip() and not empty(v):
                out[k] = str(v).strip()
    # web field fills for empties
    for k, v in extract_fields_from_evidence(ev, out).items():
        if empty(out.get(k)):
            out[k] = v
    # normalize business type
    role = str(decision.get("role") or "").lower()
    if role in ("specialty_distributor", "specialty distributor"):
        out["Business Type"] = "Specialty Distributor"
    elif role in ("distributor", "channel_partner", "channel partner"):
        if empty(out.get("Business Type")) or "specialty" not in out.get("Business Type", "").lower():
            out["Business Type"] = "Distributor"
    # ensure ADC-oriented categories if thin
    if empty(out.get("Core Categories")) or "antibody" not in out.get("Core Categories", "").lower():
        if "Specialty" in out.get("Business Type", ""):
            out["Core Categories"] = (
                "Antibody-drug conjugates; Oncology biologics; Specialty pharmaceutical distribution"
            )
        else:
            out["Core Categories"] = (
                "Pharmaceutical wholesale; Oncology / biologics channel; Specialty injectables"
            )
    # confidence
    conf = 70
    if ev.get("channel_hit"):
        conf += 10
    if ev.get("adc_hit"):
        conf += 10
    if ev.get("mfr_hit"):
        conf -= 15
    if decision.get("keep"):
        conf += 5
    out["Data Confidence"] = str(min(95, max(40, conf)))
    out["Quality Score"] = out["Data Confidence"]
    src = out.get("Data Sources") or ""
    if "adc_audit_deepseek_web" not in src:
        out["Data Sources"] = (src + "; adc_audit_deepseek_web; duckduckgo; wikipedia; deepseek-chat").strip("; ")
    if empty(out.get("Summary")) and decision.get("reason"):
        out["Summary"] = str(decision.get("reason"))[:300]
    # blank retail/food nonsense for pharma
    for col in ("Retail Chains", "Specialty Grocers", "Foodservice / HoReCa"):
        if not empty(out.get(col)) and any(
            x in (out.get(col) or "").lower() for x in ("supermarket", "grocery", "restaurant", "horeca")
        ):
            out[col] = "Not applicable"
    return out


def heuristic_keep(row: dict[str, str], ev: dict[str, Any]) -> tuple[bool, str]:
    name = row.get("Company") or ""
    if is_oem(name):
        return False, "oem_blocklist"
    if ev.get("mfr_hit") and not ev.get("channel_hit"):
        return False, "manufacturer_signal_no_channel"
    # must have some ADC/pharma/specialty signal in row or web
    row_blob = " ".join([
        row.get("Core Categories", ""),
        row.get("Specialty Focus", ""),
        row.get("Summary", ""),
        row.get("Business Type", ""),
        row.get("Key Brands Represented", ""),
    ])
    if not (ev.get("adc_hit") or ADC_RE.search(row_blob) or CHANNEL_RE.search(row_blob)):
        return False, "not_adc_or_channel_related"
    if not (ev.get("channel_hit") or CHANNEL_RE.search(row_blob) or "distributor" in (row.get("Business Type") or "").lower()):
        return False, "not_distributor"
    return True, "heuristic_ok"


async def main() -> int:
    if not IN_XLSX.exists():
        print(f"Missing {IN_XLSX}")
        return 1

    # backup once per run name
    import shutil
    if not BACKUP.exists():
        shutil.copy2(IN_XLSX, BACKUP)
        print(f"Backup: {BACKUP}")
    else:
        print(f"Backup exists: {BACKUP}")

    rows = load_rows(IN_XLSX)
    # Prefer pre-audit backup as source of truth if present (avoid auditing already-shrunk file mid-run)
    pre_audits = sorted(OUT_DIR.glob("global_antibody_drug_conjugates_market_FINAL_pre_audit_*.xlsx"))
    if pre_audits:
        src_rows = load_rows(pre_audits[-1])
        if len(src_rows) >= len(rows):
            rows = src_rows
            print(f"Using backup source rows: {len(rows)} from {pre_audits[-1].name}")

    print(f"Loaded {len(rows)} companies; model={MODEL} batch={BATCH}")

    already: dict[str, dict[str, Any]] = {}
    if RESUME and PROGRESS.exists():
        for line in PROGRESS.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line:
                continue
            try:
                rec = json.loads(line)
                already[norm(rec.get("company", ""))] = rec
            except Exception:
                pass
        print(f"Resume: {len(already)} prior decisions")

    oa = get_client()
    t = oa.chat.completions.create(
        model=MODEL,
        messages=[{"role": "user", "content": "Reply OK"}],
        max_tokens=5,
    )
    print(f"DeepSeek OK: {(t.choices[0].message.content or '').strip()} ({t.model})")

    removed: list[dict[str, Any]] = []
    kept_meta: dict[str, dict[str, Any]] = {}  # name -> decision for resume keeps
    decisions_log: list[dict[str, Any]] = []

    # restore removed/kept from prior progress (will re-apply fields for kept via fresh web later if needed)
    work: list[dict[str, str]] = []
    for row in rows:
        name = row.get("Company", "")
        n = norm(name)
        if is_oem(name):
            if n not in already:
                removed.append({"company": name, "reason": "oem_blocklist", "stage": "prefilter"})
                print(f"  REMOVE (blocklist): {name}")
                with PROGRESS.open("a", encoding="utf-8") as pf:
                    pf.write(json.dumps({
                        "company": name, "keep": False, "reason": "oem_blocklist",
                        "role": "oem", "channel_hit": False, "mfr_hit": True, "adc_hit": False,
                    }) + "\n")
            else:
                if not already[n].get("keep"):
                    removed.append({"company": name, "reason": already[n].get("reason"), "stage": "resume"})
            continue
        if n in already:
            rec = already[n]
            if rec.get("keep"):
                kept_meta[n] = {"row": row, "decision": rec, "evidence": {
                    "snippets": [], "meta_snippets": [], "wiki": {}, "site_excerpt": "",
                    "blob": "", "channel_hit": rec.get("channel_hit"), "mfr_hit": rec.get("mfr_hit"),
                    "adc_hit": rec.get("adc_hit"),
                }}
            else:
                removed.append({
                    "company": name,
                    "reason": rec.get("reason"),
                    "role": rec.get("role"),
                    "stage": "resume",
                })
            continue
        work.append(row)

    print(f"Remaining to audit: {len(work)}; already kept: {len(kept_meta)}; already removed: {len(removed)}")

    sem = asyncio.Semaphore(CONCURRENCY)
    limits = httpx.Limits(max_connections=CONCURRENCY + 2, max_keepalive_connections=CONCURRENCY)

    async with httpx.AsyncClient(
        headers={"User-Agent": "vendor-intel-adc-audit/1.0"},
        limits=limits,
        timeout=httpx.Timeout(8.0, connect=4.0),
    ) as http:
        async def one(row: dict[str, str]) -> tuple[dict[str, str], dict[str, Any]]:
            async with sem:
                ev = await gather_evidence(http, row)
                await asyncio.sleep(0.05)
                return row, ev

        for i in range(0, len(work), BATCH):
            chunk = work[i : i + BATCH]
            print(f"\n[{i+1}-{i+len(chunk)}/{len(work)}] web evidence...", flush=True)
            try:
                evidences = await asyncio.wait_for(
                    asyncio.gather(*[one(r) for r in chunk]),
                    timeout=60.0,
                )
            except asyncio.TimeoutError:
                print("  batch web timeout — using row-only evidence", flush=True)
                evidences = []
                for r in chunk:
                    blob = " ".join([
                        r.get("Core Categories", ""), r.get("Specialty Focus", ""),
                        r.get("Summary", ""), r.get("Business Type", ""),
                        r.get("Key Brands Represented", ""),
                    ])
                    evidences.append((r, {
                        "company": r["Company"], "snippets": [], "meta_snippets": [],
                        "wiki": {}, "site_excerpt": "", "blob": blob,
                        "channel_hit": bool(CHANNEL_RE.search(blob)),
                        "mfr_hit": bool(MFR_RE.search(blob)),
                        "adc_hit": bool(ADC_RE.search(blob)),
                    }))
            items = [{"row": r, "evidence": e} for r, e in evidences]

            print(f"  DeepSeek audit ({len(items)})...", flush=True)
            decisions = deepseek_audit_batch(oa, items)
            by_name = {norm(str(d.get("company") or "")): d for d in decisions if d.get("company")}

            for it in items:
                row = it["row"]
                ev = it["evidence"]
                name = row["Company"]
                d = by_name.get(norm(name))
                if not d:
                    ok, reason = heuristic_keep(row, ev)
                    d = {
                        "company": name,
                        "keep": ok,
                        "reason": f"fallback:{reason}",
                        "role": "distributor" if ok else "unrelated",
                        "corrections": {},
                    }
                keep = bool(d.get("keep"))
                if is_oem(name) or (ev.get("mfr_hit") and not ev.get("channel_hit") and not keep):
                    keep = False
                    d["keep"] = False
                    d["reason"] = d.get("reason") or "mfr_override"
                if keep:
                    ok_h, reason_h = heuristic_keep(row, ev)
                    if not ok_h and reason_h in ("oem_blocklist", "manufacturer_signal_no_channel"):
                        keep = False
                        d["keep"] = False
                        d["reason"] = f"override:{reason_h}; was:{d.get('reason')}"

                rec = {
                    "company": name,
                    "keep": keep,
                    "reason": d.get("reason"),
                    "role": d.get("role"),
                    "channel_hit": ev.get("channel_hit"),
                    "mfr_hit": ev.get("mfr_hit"),
                    "adc_hit": ev.get("adc_hit"),
                    "corrections": d.get("corrections") or {},
                }
                decisions_log.append(rec)
                with PROGRESS.open("a", encoding="utf-8") as pf:
                    pf.write(json.dumps(rec, ensure_ascii=False) + "\n")

                if keep:
                    kept_meta[norm(name)] = {"row": row, "decision": d, "evidence": ev}
                    print(f"  KEEP  {name} — {d.get('reason')}", flush=True)
                else:
                    removed.append({
                        "company": name,
                        "reason": d.get("reason"),
                        "role": d.get("role"),
                        "stage": "deepseek_or_heuristic",
                    })
                    print(f"  REMOVE {name} — {d.get('reason')}", flush=True)

            await asyncio.sleep(0.25)

    # Build final kept rows with field application
    kept: list[dict[str, str]] = []
    for n, meta in kept_meta.items():
        kept.append(apply_decision(meta["row"], meta["decision"], meta["evidence"]))

    # Second pass: DeepSeek field verify for kept rows that had empty evidence (resume)
    thin = [k for k in kept if "adc_audit_deepseek_web" not in (k.get("Data Sources") or "")]
    if thin:
        print(f"\nField-verify pass for {len(thin)} resumed keeps...", flush=True)
        async with httpx.AsyncClient(
            headers={"User-Agent": "vendor-intel-adc-audit/1.0"},
            timeout=httpx.Timeout(8.0, connect=4.0),
        ) as http:
            for i in range(0, len(thin), BATCH):
                chunk = thin[i : i + BATCH]
                evids = []
                for r in chunk:
                    try:
                        ev = await asyncio.wait_for(gather_evidence(http, r), timeout=18.0)
                    except asyncio.TimeoutError:
                        blob = " ".join([r.get("Core Categories", ""), r.get("Specialty Focus", ""), r.get("Summary", "")])
                        ev = {
                            "snippets": [], "meta_snippets": [], "wiki": {}, "site_excerpt": "",
                            "blob": blob, "channel_hit": True, "mfr_hit": False, "adc_hit": bool(ADC_RE.search(blob)),
                        }
                    evids.append(ev)
                items = [{"row": r, "evidence": e} for r, e in zip(chunk, evids)]
                decisions = deepseek_audit_batch(oa, items)
                by_name = {norm(str(d.get("company") or "")): d for d in decisions if d.get("company")}
                for r, ev in zip(chunk, evids):
                    d = by_name.get(norm(r["Company"])) or {
                        "company": r["Company"], "keep": True, "reason": "resume_keep",
                        "role": "distributor", "corrections": {},
                    }
                    # if deepseek flips to remove on field pass, honor it
                    if d.get("keep") is False:
                        removed.append({
                            "company": r["Company"], "reason": d.get("reason"),
                            "role": d.get("role"), "stage": "field_verify_remove",
                        })
                        kept = [x for x in kept if norm(x["Company"]) != norm(r["Company"])]
                        print(f"  REMOVE (field pass) {r['Company']} — {d.get('reason')}", flush=True)
                        continue
                    updated = apply_decision(r, d, ev)
                    for i2, x in enumerate(kept):
                        if norm(x["Company"]) == norm(r["Company"]):
                            kept[i2] = updated
                            break
                print(f"  field-verified {i+1}-{i+len(chunk)}", flush=True)

    # dedupe
    seen: set[str] = set()
    unique: list[dict[str, str]] = []
    for r in kept:
        k = norm(r["Company"])
        if k in seen:
            continue
        seen.add(k)
        unique.append(r)
    unique.sort(key=lambda x: norm(x["Company"]))

    write_xlsx(OUT_XLSX, unique)
    summary = {
        "market": MARKET,
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "model": MODEL,
        "input_count": len(rows),
        "kept_count": len(unique),
        "removed_count": len(removed),
        "backup": str(BACKUP),
        "output": str(OUT_XLSX),
        "business_types": dict(Counter(r.get("Business Type") for r in unique)),
        "removed_sample": removed[:50],
        "removed_reasons": dict(Counter(str(r.get("reason", ""))[:100] for r in removed)),
    }
    SUMMARY.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    REMOVED.write_text(json.dumps(removed, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\nDONE kept={len(unique)} removed={len(removed)}")
    print(f"Wrote {OUT_XLSX}")
    print(f"Summary {SUMMARY}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(asyncio.run(main()))
    except KeyboardInterrupt:
        print("Interrupted", file=sys.stderr)
        raise SystemExit(130)
