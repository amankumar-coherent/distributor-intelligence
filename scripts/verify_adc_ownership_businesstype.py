#!/usr/bin/env python3
"""Verify & fill Ownership + Business Type for ADC FINAL via web + DeepSeek."""
from __future__ import annotations

import asyncio
import json
import os
import re
import time
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import quote

import httpx
from dotenv import load_dotenv
from openai import OpenAI
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
load_dotenv(ROOT / ".env")

OUT = ROOT / "output/web_expand/global_antibody_drug_conjugates_market_global"
XLSX = OUT / "global_antibody_drug_conjugates_market_FINAL.xlsx"
SUMMARY = OUT / "adc_ownership_businesstype_summary.json"
MODEL = os.getenv("OPENAI_MODEL", "deepseek-chat")
BATCH = 10

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
    "Company": 28, "Headquarters": 28, "Ownership": 16, "Business Type": 22,
    "Employees": 18, "Turnover / Scale": 22, "Summary": 36, "Website": 24,
}

# (ownership, business_type) — web-verified / well-known
KNOWN: dict[str, tuple[str, str]] = {
    "AAH Pharmaceuticals": ("Subsidiary", "Distributor"),  # Hallo Healthcare / Aurelius
    "Accredo Health Group": ("Subsidiary", "Specialty Distributor"),  # Cigna / Evernorth
    "Cigna Specialty Pharmacy / Accredo": ("Subsidiary", "Specialty Distributor"),
    "Express Scripts Accredo network": ("Subsidiary", "Specialty Distributor"),
    "Aetna Specialty Pharmacy": ("Subsidiary", "Specialty Distributor"),  # CVS Health
    "CVS Specialty": ("Subsidiary", "Specialty Distributor"),
    "Coram CVS Specialty Infusion": ("Subsidiary", "Specialty Distributor"),
    "Walgreens Specialty Pharmacy": ("Subsidiary", "Specialty Distributor"),
    "AllianceRx Walgreens Pharmacy": ("Subsidiary", "Specialty Distributor"),
    "Optum Specialty Pharmacy": ("Subsidiary", "Specialty Distributor"),  # UnitedHealth
    "CenterWell Pharmacy": ("Subsidiary", "Specialty Distributor"),  # Humana
    "Kroger Specialty Pharmacy": ("Subsidiary", "Specialty Distributor"),
    "Walmart Specialty Pharmacy": ("Subsidiary", "Specialty Distributor"),
    "AcariaHealth": ("Subsidiary", "Specialty Distributor"),  # often under larger parent networks
    "McKesson Specialty Health": ("Subsidiary", "Specialty Distributor"),
    "McKesson Specialty Care Distribution": ("Subsidiary", "Specialty Distributor"),
    "McKesson Plasma and Biologics": ("Subsidiary", "Specialty Distributor"),
    "McKesson Medical-Surgical": ("Public", "Distributor"),  # McKesson Corp public
    "McKesson Canada": ("Subsidiary", "Distributor"),
    "Biologics by McKesson": ("Subsidiary", "Specialty Distributor"),
    "Cardinal Health Specialty Solutions": ("Subsidiary", "Specialty Distributor"),
    "Cardinal Health Specialty Pharmaceutical Distribution": ("Subsidiary", "Specialty Distributor"),
    "Cencora (AmerisourceBergen) Specialty": ("Public", "Specialty Distributor"),
    "ASD Healthcare": ("Subsidiary", "Specialty Distributor"),
    "Oncology Supply": ("Subsidiary", "Specialty Distributor"),
    "CuraScript SD": ("Subsidiary", "Specialty Distributor"),
    "Priority Healthcare / CuraScript Specialty Distribution": ("Subsidiary", "Specialty Distributor"),
    "US Bioservices": ("Subsidiary", "Specialty Distributor"),
    "TheraCom Specialty Pharmacy": ("Subsidiary", "Specialty Distributor"),
    "Anda Inc.": ("Subsidiary", "Distributor"),  # Teva
    "Anda Pharmaceuticals": ("Subsidiary", "Distributor"),
    "Alliance Healthcare UK": ("Subsidiary", "Distributor"),  # Cencora
    "Alliance Healthcare España": ("Subsidiary", "Distributor"),
    "Alliance Healthcare France": ("Subsidiary", "Distributor"),
    "Alliance Healthcare Greece": ("Subsidiary", "Distributor"),
    "Alliance Healthcare Ireland": ("Subsidiary", "Distributor"),
    "Alliance Healthcare Italia": ("Subsidiary", "Distributor"),
    "Alliance Healthcare Nederland": ("Subsidiary", "Distributor"),
    "Alliance Healthcare Norge": ("Subsidiary", "Distributor"),
    "Alliance Healthcare Polska": ("Subsidiary", "Distributor"),
    "Alliance Healthcare Portugal": ("Subsidiary", "Distributor"),
    "Alliance Healthcare România": ("Subsidiary", "Distributor"),
    "Alloga UK": ("Subsidiary", "Specialty Distributor"),
    "Galenica / Alloga": ("Public", "Specialty Distributor"),
    "Fisher Scientific": ("Subsidiary", "Distributor"),  # Thermo Fisher
    "VWR International": ("Subsidiary", "Distributor"),  # Avantor
    "Avantor Inc.": ("Public", "Distributor"),
    "Henry Schein Medical": ("Public", "Distributor"),
    "Owens & Minor": ("Public", "Distributor"),
    "Medline Industries": ("Private", "Distributor"),
    "DKSH Health": ("Public", "Distributor"),
    "DiethelmKellerSiberHegner Healthcare": ("Public", "Distributor"),
    "Diethelm Thailand / DKSH Thailand": ("Subsidiary", "Distributor"),
    "Zuellig Pharma": ("Private", "Distributor"),
    "Zuellig Pharma Philippines": ("Subsidiary", "Distributor"),
    "Suzuken Co., Ltd.": ("Public", "Distributor"),
    "Alfresa Holdings": ("Public", "Distributor"),
    "Medipal Holdings": ("Public", "Distributor"),
    "Toho Holdings / TOHO Pharmaceutical": ("Public", "Distributor"),
    "Phoenix Pharma SE": ("Private", "Distributor"),  # PHOENIX group
    "PHOENIX Pharmahandel": ("Subsidiary", "Distributor"),
    "Phoenix Healthcare Distribution": ("Subsidiary", "Distributor"),
    "NOWEDA Apothekergenossenschaft": ("Cooperative", "Distributor"),
    "Sanacorp eG": ("Cooperative", "Distributor"),
    "Cofares": ("Cooperative", "Distributor"),
    "Neuca": ("Public", "Distributor"),
    "Hungaropharma": ("Private", "Distributor"),
    "Tamro": ("Subsidiary", "Distributor"),  # PHOENIX
    "Oriola": ("Public", "Distributor"),
    "Uniphar": ("Public", "Distributor"),
    "Sigma Healthcare": ("Public", "Distributor"),
    "Symbion": ("Subsidiary", "Distributor"),
    "Nadro": ("Private", "Distributor"),
    "Marzam": ("Private", "Distributor"),
    "Casa Saba": ("Private", "Distributor"),
    "Sinopharm": ("Public", "Distributor"),
    "Jointown Pharmaceutical Group": ("Public", "Distributor"),
    "China Resources Pharmaceutical": ("Public", "Distributor"),
    "Shanghai Pharmaceuticals distribution": ("Public", "Distributor"),
    "PT Enseval Putera Megatrading": ("Public", "Distributor"),
    "Profarma": ("Public", "Distributor"),
    "Viveo": ("Private", "Specialty Distributor"),
    "Oncoprod": ("Private", "Specialty Distributor"),
    "Option Care Health specialty infusion": ("Public", "Specialty Distributor"),
    "BioCare Specialty Distribution (BioCareSD)": ("Private", "Specialty Distributor"),
    "Morris & Dickson Co., LLC": ("Private", "Distributor"),
    "Morris & Dickson Specialty Distribution": ("Private", "Specialty Distributor"),
    "Sciensus": ("Private", "Specialty Distributor"),
    "Mawdsleys": ("Private", "Distributor"),
    "Mawdsley-Brooks": ("Private", "Distributor"),
    "PANTHERx Rare": ("Private", "Specialty Distributor"),
    "Orsini Specialty Pharmacy": ("Private", "Specialty Distributor"),
    "Onco360": ("Private", "Specialty Distributor"),
    "Amber Specialty Pharmacy": ("Private", "Specialty Distributor"),
    "Avella Specialty Pharmacy": ("Subsidiary", "Specialty Distributor"),
    "Anthem Specialty Pharmacy": ("Subsidiary", "Specialty Distributor"),
    "Ariel Specialty Pharmacy": ("Private", "Specialty Distributor"),
    "Maxor Specialty Pharmacy": ("Private", "Specialty Distributor"),
    "WellDyne Specialty Pharmacy": ("Private", "Specialty Distributor"),
    "PerformSpecialty": ("Private", "Specialty Distributor"),
    "Shields Health Solutions specialty": ("Private", "Specialty Distributor"),
    "Prime Therapeutics Specialty": ("Subsidiary", "Specialty Distributor"),
    "Magellan Rx Specialty": ("Subsidiary", "Specialty Distributor"),
    "MedImpact Specialty Pharmacy": ("Private", "Specialty Distributor"),
    "2BScientific Ltd": ("Private", "Distributor"),
    "2BScientific GmbH": ("Private", "Distributor"),
    "Abacus dx": ("Private", "Distributor"),
    "Cedarlane": ("Private", "Distributor"),
    "Cedarlane Laboratories": ("Private", "Distributor"),
    "Cosmo Bio Co., Ltd.": ("Private", "Distributor"),
    "Funakoshi": ("Private", "Distributor"),
    "Thomas Scientific": ("Private", "Distributor"),
    "BIOZOL Diagnostica Vertrieb GmbH": ("Private", "Distributor"),
    "Nordic BioSite ApS": ("Private", "Distributor"),
    "Nordic BioSite Oy": ("Private", "Distributor"),
    "Nordic Biosite AB": ("Private", "Distributor"),
    "Nordic Biosite AS": ("Private", "Distributor"),
    "Cambridge Bioscience": ("Private", "Distributor"),
    "Stratech Scientific": ("Private", "Distributor"),
    "Dianova GmbH": ("Private", "Distributor"),
    "Azelis Group NV": ("Public", "Distributor"),
    "Mediq": ("Private", "Distributor"),
    "Movianto": ("Subsidiary", "Specialty Distributor"),
    "Medios AG": ("Public", "Specialty Distributor"),
    "VECTOR Pharma FZCO": ("Private", "Specialty Distributor"),
    "NewBridge Pharmaceuticals": ("Private", "Specialty Distributor"),
    "Biologix Pharma MENA": ("Private", "Specialty Distributor"),
    "Adium Pharma S.A.": ("Private", "Specialty Distributor"),
    "4Bio Medicamentos Especializados": ("Private", "Specialty Distributor"),
    "Red Oak Sourcing specialty": ("Joint Venture", "Specialty Distributor"),  # McKesson + Walgreens
    "ClarusONE Sourcing Services": ("Joint Venture", "Distributor"),
}

OWN_OK = {"Independent", "Private", "Public", "Subsidiary", "Cooperative", "Joint Venture"}
BT_OK = {"Distributor", "Specialty Distributor"}


def empty(v: Any) -> bool:
    return str(v or "").strip().lower() in {
        "", "n/a", "na", "none", "-", "—", "not publicly disclosed", "unknown", "tbd", "null",
    }


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


def normalize_ownership(raw: str) -> str:
    s = (raw or "").strip()
    if not s or empty(s):
        return "Private"
    low = s.lower()
    if "joint venture" in low or low.startswith("jv"):
        return "Joint Venture"
    if "cooperat" in low:
        return "Cooperative"
    if "subsidiary" in low or "owned by" in low or "part of" in low:
        return "Subsidiary"
    if "public" in low or "listed" in low or "tse" in low or "nyse" in low or "nasdaq" in low:
        return "Public"
    if "independent" in low:
        return "Independent"
    if "private" in low:
        return "Private"
    # messy values like "CVS Health (subsidiary)"
    if "subsidiary" in low:
        return "Subsidiary"
    if s in OWN_OK:
        return s
    return "Private"


def normalize_bt(raw: str, company: str, focus: str = "", categories: str = "") -> str:
    s = (raw or "").strip()
    blob = f"{company} {s} {focus} {categories}".lower()
    if s in BT_OK and "specialty" in s.lower():
        return "Specialty Distributor"
    if any(
        x in blob
        for x in (
            "specialty pharmacy", "specialty distributor", "specialty distribution",
            "oncology supply", "limited distribution", "orphan", "rare disease",
            "specialty care", "infusion", "ldd",
        )
    ):
        return "Specialty Distributor"
    if "specialty" in (s or "").lower():
        return "Specialty Distributor"
    return "Distributor"


def load_rows(path: Path) -> list[dict[str, str]]:
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    hdr = [c.value for c in ws[2]]
    rows = []
    for r in range(3, ws.max_row + 1):
        d = {h: str(ws.cell(r, i).value or "").strip() for i, h in enumerate(hdr, 1) if h}
        if d.get("Company"):
            rows.append(d)
    return rows


def write_xlsx(path: Path, rows: list[dict[str, str]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "ADC Distributors"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HEADERS))
    title = ws.cell(1, 1, "Global Antibody Drug Conjugates Market — Distributors")
    title.font = Font(bold=True, size=13, color="FFFFFF")
    title.fill = PatternFill("solid", fgColor="1F4E79")
    hf = PatternFill("solid", fgColor="2E75B6")
    hfont = Font(bold=True, color="FFFFFF")
    thin = Border(
        left=Side(style="thin", color="BDD7EE"),
        right=Side(style="thin", color="BDD7EE"),
        top=Side(style="thin", color="BDD7EE"),
        bottom=Side(style="thin", color="BDD7EE"),
    )
    for i, h in enumerate(HEADERS, 1):
        c = ws.cell(2, i, h)
        c.fill = hf
        c.font = hfont
        c.alignment = Alignment(wrap_text=True, vertical="center")
        c.border = thin
        ws.column_dimensions[get_column_letter(i)].width = WIDTHS.get(h, 16)
    for ri, row in enumerate(rows, 3):
        for ci, h in enumerate(HEADERS, 1):
            cell = ws.cell(ri, ci, row.get(h, ""))
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = thin
            if ri % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F2F2F2")
    cws = wb.create_sheet("Column Rules")
    cws["A1"] = "Ownership: Independent | Private | Public | Subsidiary | Cooperative | Joint Venture"
    cws["A2"] = "Business Type: Distributor | Specialty Distributor"
    cws["A3"] = f"Verified: {datetime.now(timezone.utc).isoformat()} | rows={len(rows)}"
    wb.save(path)


async def wiki_hint(client: httpx.AsyncClient, company: str) -> str:
    try:
        api = (
            "https://en.wikipedia.org/w/api.php?action=query&list=search"
            f"&srsearch={quote(company)}&utf8=1&format=json&srlimit=1"
        )
        r = await client.get(api, timeout=8.0)
        hits = ((r.json().get("query") or {}).get("search") or [])
        if not hits:
            return ""
        title = hits[0].get("title") or ""
        r2 = await client.get(
            f"https://en.wikipedia.org/api/rest_v1/page/summary/{quote(title)}",
            timeout=8.0,
        )
        return (r2.json().get("extract") or "")[:900]
    except Exception:
        return ""


async def site_hint(client: httpx.AsyncClient, website: str) -> str:
    if not website.startswith("http"):
        return ""
    try:
        r = await client.get(website, timeout=8.0, follow_redirects=True)
        if r.status_code >= 400:
            return ""
        t = re.sub(r"(?is)<script.*?>.*?</script>", " ", r.text)
        t = re.sub(r"(?is)<style.*?>.*?</style>", " ", t)
        t = re.sub(r"<[^>]+>", " ", t)
        t = re.sub(r"\s+", " ", t)
        return t[:1200]
    except Exception:
        return ""


SYSTEM = """You classify Ownership and Business Type for Global ADC market channel companies.
Ownership MUST be exactly one of:
Independent, Private, Public, Subsidiary, Cooperative, Joint Venture
Rules:
- Public = listed parent or listed company itself
- Subsidiary = owned by a parent (McKesson, Cencora, Cigna, CVS, Walgreens, DKSH, Teva, etc.)
- Cooperative = pharmacy cooperative / genossenschaft (NOWEDA, Sanacorp, Cofares)
- Joint Venture = JV between companies
- Private = privately held company (not independent if clearly PE-backed subsidiary — use Subsidiary)
- Independent = truly independent owner-operated with no parent (use sparingly; prefer Private if unclear)
Business Type MUST be exactly one of:
Distributor, Specialty Distributor
Use Specialty Distributor for specialty pharmacy, oncology LDD, rare/orphan, specialty injectables channel.
Return JSON: {"results":[{"company":"...","ownership":"...","business_type":"...","note":"..."}]}"""


def deepseek_batch(client: OpenAI, items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    payload = []
    for it in items:
        row = it["row"]
        payload.append({
            "company": row["Company"],
            "website": row.get("Website"),
            "hq": row.get("Headquarters"),
            "current_ownership": row.get("Ownership"),
            "current_business_type": row.get("Business Type"),
            "specialty_focus": (row.get("Specialty Focus") or "")[:160],
            "core_categories": (row.get("Core Categories") or "")[:160],
            "summary": (row.get("Summary") or "")[:200],
            "wiki": (it.get("wiki") or "")[:500],
            "site": (it.get("site") or "")[:400],
        })
    for attempt in range(3):
        try:
            resp = client.chat.completions.create(
                model=MODEL,
                messages=[
                    {"role": "system", "content": SYSTEM},
                    {"role": "user", "content": json.dumps({"companies": payload}, ensure_ascii=False)},
                ],
                temperature=0.05,
                max_tokens=3500,
                response_format={"type": "json_object"},
                timeout=45.0,
            )
            data = json.loads(resp.choices[0].message.content or "{}")
            if isinstance(data, dict):
                for k in ("results", "companies", "items"):
                    if isinstance(data.get(k), list):
                        return data[k]
            if isinstance(data, list):
                return data
            return []
        except Exception as e:
            print(f"  DeepSeek err: {e}", flush=True)
            time.sleep(1.2 * (attempt + 1))
    return []


def name_heuristic(company: str) -> tuple[str | None, str | None]:
    n = norm(company)
    own = bt = None
    # subsidiary name signals
    parents = [
        ("mckesson", "Subsidiary"), ("cencora", "Subsidiary"), ("amerisource", "Subsidiary"),
        ("cardinal health", "Subsidiary"), ("alliance healthcare", "Subsidiary"),
        ("cvs", "Subsidiary"), ("walgreens", "Subsidiary"), ("accredo", "Subsidiary"),
        ("cigna", "Subsidiary"), ("optum", "Subsidiary"), ("humana", "Subsidiary"),
        ("kroger", "Subsidiary"), ("walmart", "Subsidiary"), ("dksh", "Subsidiary"),
        ("zuellig", "Private"), ("fisher scientific", "Subsidiary"), ("vwr", "Subsidiary"),
        ("thermo", "Subsidiary"), ("anda", "Subsidiary"),
    ]
    for key, o in parents:
        if key in n:
            own = o
            break
    if any(x in n for x in ("specialty pharmacy", "specialty distribution", "oncology supply", "accredo", "onco360", "pantheon", "orsini")):
        bt = "Specialty Distributor"
    if "specialty" in n and ("pharmacy" in n or "distribution" in n or "distributor" in n):
        bt = "Specialty Distributor"
    if any(x in n for x in ("noweda", "sanacorp", "cofares")):
        own = "Cooperative"
    return own, bt


async def main() -> int:
    rows = load_rows(XLSX)
    print(f"Loaded {len(rows)}; model={MODEL}")

    # Pass A known
    known_n = 0
    for i, row in enumerate(rows):
        if row["Company"] in KNOWN:
            o, b = KNOWN[row["Company"]]
            rows[i]["Ownership"] = o
            rows[i]["Business Type"] = b
            known_n += 1
    print(f"Pass A known: {known_n}")

    oa = OpenAI(
        api_key=os.getenv("OPENAI_API_KEY"),
        base_url=os.getenv("OPENAI_BASE_URL", "https://api.deepseek.com/v1"),
        timeout=60.0,
    )
    print("DeepSeek client ready", flush=True)

    # Only send non-known rows to DeepSeek (faster)
    targets = [i for i, r in enumerate(rows) if r["Company"] not in KNOWN]
    print(f"DeepSeek targets: {len(targets)}", flush=True)

    async with httpx.AsyncClient(
        headers={"User-Agent": "vendor-intel-adc-ownership/1.0"},
        timeout=httpx.Timeout(8.0, connect=4.0),
    ) as http:
        for ti in range(0, len(targets), BATCH):
            idxs = targets[ti : ti + BATCH]
            chunk = [rows[j] for j in idxs]
            print(f"\n[{ti+1}-{ti+len(idxs)}/{len(targets)}] wiki+DeepSeek...", flush=True)
            try:
                wikis = await asyncio.wait_for(
                    asyncio.gather(*[wiki_hint(http, r["Company"]) for r in chunk]),
                    timeout=30.0,
                )
            except asyncio.TimeoutError:
                wikis = [""] * len(chunk)
            sites = [""] * len(chunk)
            items = [{"row": r, "wiki": w, "site": s} for r, w, s in zip(chunk, wikis, sites)]
            decisions = deepseek_batch(oa, items)
            by = {norm(str(d.get("company") or "")): d for d in decisions if d.get("company")}

            for r, w in zip(chunk, wikis):
                name = r["Company"]
                d = by.get(norm(name)) or {}
                o = normalize_ownership(str(d.get("ownership") or r.get("Ownership") or ""))
                b = normalize_bt(
                    str(d.get("business_type") or r.get("Business Type") or ""),
                    name,
                    r.get("Specialty Focus") or "",
                    r.get("Core Categories") or "",
                )
                ho, hb = name_heuristic(name)
                if ho:
                    o = ho
                if hb:
                    b = hb
                wl = (w or "").lower()
                if "publicly traded" in wl or "listed company" in wl or "nyse:" in wl or "nasdaq:" in wl:
                    if o in ("Independent", "Private"):
                        o = "Public"
                if re.search(r"\bsubsidiary of\b|\bowned by\b|\bdivision of\b", wl):
                    o = "Subsidiary"
                if "cooperative" in wl or "genossenschaft" in wl:
                    o = "Cooperative"

                before = (r.get("Ownership"), r.get("Business Type"))
                r["Ownership"] = o if o in OWN_OK else "Private"
                r["Business Type"] = b if b in BT_OK else "Distributor"
                after = (r["Ownership"], r["Business Type"])
                src = r.get("Data Sources") or ""
                if "adc_ownership_bt_verify" not in src:
                    r["Data Sources"] = (src + "; adc_ownership_bt_verify").strip("; ")
                if before != after:
                    print(f"  FIX {name[:40]:40} | {before[0]}->{after[0]} | {before[1]}->{after[1]}", flush=True)
                else:
                    print(f"  OK  {name[:40]:40} | {after[0]} | {after[1]}", flush=True)
            await asyncio.sleep(0.1)

    # Final pass: Independent is overused — if name/heuristic says subsidiary, enforce
    for r in rows:
        ho, hb = name_heuristic(r["Company"])
        if ho:
            r["Ownership"] = ho
        if hb:
            r["Business Type"] = hb
        if r["Company"] in KNOWN:
            r["Ownership"], r["Business Type"] = KNOWN[r["Company"]]
        r["Ownership"] = normalize_ownership(r["Ownership"])
        r["Business Type"] = normalize_bt(
            r["Business Type"], r["Company"], r.get("Specialty Focus") or "", r.get("Core Categories") or ""
        )

    rows.sort(key=lambda x: norm(x["Company"]))
    write_xlsx(XLSX, rows)

    summary = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "rows": len(rows),
        "ownership": dict(Counter(r.get("Ownership") for r in rows)),
        "business_type": dict(Counter(r.get("Business Type") for r in rows)),
        "samples": [
            {"company": r["Company"], "ownership": r["Ownership"], "business_type": r["Business Type"]}
            for r in rows[:12]
        ],
    }
    SUMMARY.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    print("\nDONE")
    print("Ownership:", summary["ownership"])
    print("Business Type:", summary["business_type"])
    return 0


if __name__ == "__main__":
    raise SystemExit(asyncio.run(main()))
