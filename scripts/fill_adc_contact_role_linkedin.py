#!/usr/bin/env python3
"""Fill Contact Person, Role, LinkedIn (person) for ADC FINAL via web search.

Sources: company about/leadership pages, DuckDuckGo HTML, Wikipedia,
public leadership directories. Prefer linkedin.com/in/ person URLs.
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
OUT = ROOT / "output/web_expand/global_antibody_drug_conjugates_market_global"
XLSX = OUT / "global_antibody_drug_conjugates_market_FINAL.xlsx"
SUMMARY = OUT / "adc_contact_role_linkedin_summary.json"
PROGRESS = OUT / "adc_contact_progress.jsonl"

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
    "Company": 28, "Contact Person": 22, "Role": 28, "LinkedIn": 36,
    "Headquarters": 24, "Website": 24, "Summary": 36,
}

NPD = {"", "n/a", "na", "none", "-", "—", "not publicly disclosed", "unknown", "tbd", "null"}

# (person, role, person_linkedin_or_empty)
KNOWN: dict[str, tuple[str, str, str]] = {
    "Suzuken Co., Ltd.": ("Shigeru Asano", "President and CEO", ""),
    "Alfresa Holdings": ("Taiji Ito", "President", ""),
    "Medipal Holdings": ("Shuichi Watanabe", "President and CEO", ""),
    "Toho Holdings / TOHO Pharmaceutical": ("Atsushi Udoh", "President and CEO", ""),
    "Oncology Supply": ("Annette Bentley", "VP and General Manager", "https://www.linkedin.com/in/annette-bentley-9524802a"),
    "Cencora (AmerisourceBergen) Specialty": ("Lisa Harrison", "SVP & President, Specialty Distribution & Solutions", ""),
    "ASD Healthcare": ("Lisa Harrison", "SVP & President, Specialty Distribution & Solutions (Cencora)", ""),
    "Besse Medical": ("Lisa Harrison", "SVP & President, Specialty Distribution & Solutions (Cencora)", ""),
    "Avantor Inc.": ("Michael Stubblefield", "President and CEO", ""),
    "Henry Schein Medical": ("Stanley Bergman", "Chairman and CEO", ""),
    "Owens & Minor": ("Edward A. Pesicka", "President and CEO", ""),
    "Medline Industries": ("Jim Boyle", "CEO", ""),
    "Option Care Health specialty infusion": ("John C. Rademacher", "President and CEO", ""),
    "DKSH Health": ("Stefan P. Butz", "CEO", ""),
    "DiethelmKellerSiberHegner Healthcare": ("Stefan P. Butz", "CEO (DKSH)", ""),
    "Diethelm Thailand / DKSH Thailand": ("Stefan P. Butz", "CEO (DKSH)", ""),
    "Zuellig Pharma": ("John Graham", "Group CEO", "https://www.linkedin.com/in/john-a-graham-5a6b471"),
    "Zuellig Pharma Philippines": ("John Graham", "Group CEO (Zuellig Pharma)", "https://www.linkedin.com/in/john-a-graham-5a6b471"),
    "Neuca": ("Piotr Sucharski", "President of the Management Board", ""),
    "Cofares": ("Eduardo Pastor", "President", ""),
    "Uniphar": ("Gerard Rabbette", "CEO", ""),
    "Oriola": ("Katarina Gabrielsson", "CEO", ""),
    "Sigma Healthcare": ("Vikesh Ramsunder", "CEO", ""),
    "Galenica / Alloga": ("Marc Werner", "CEO (Galenica)", ""),
    "Medios AG": ("Matthias Gärtner", "CEO", ""),
    "Azelis Group NV": ("Anna Bertona", "CEO", ""),
    "Profarma": ("Samir Santos de Souza", "CEO", ""),
    "PT Enseval Putera Megatrading": ("Maria Leovina", "President Director", ""),
    "Sinopharm": ("Liu Jingzhen", "Chairman", ""),
    "Jointown Pharmaceutical Group": ("Liu Baolin", "Chairman", ""),
    "China Resources Pharmaceutical": ("Han Yuewei", "Chairman", ""),
    "McKesson Specialty Health": ("Brian Tyler", "CEO (McKesson)", "https://www.linkedin.com/in/brian-tyler"),
    "McKesson Specialty Care Distribution": ("Brian Tyler", "CEO (McKesson)", "https://www.linkedin.com/in/brian-tyler"),
    "McKesson Plasma and Biologics": ("Brian Tyler", "CEO (McKesson)", "https://www.linkedin.com/in/brian-tyler"),
    "McKesson Medical-Surgical": ("Brian Tyler", "CEO (McKesson)", "https://www.linkedin.com/in/brian-tyler"),
    "McKesson Canada": ("Brian Tyler", "CEO (McKesson)", "https://www.linkedin.com/in/brian-tyler"),
    "Biologics by McKesson": ("Brian Tyler", "CEO (McKesson)", "https://www.linkedin.com/in/brian-tyler"),
    "Cardinal Health Specialty Solutions": ("Jason Hollar", "CEO (Cardinal Health)", ""),
    "Cardinal Health Specialty Pharmaceutical Distribution": ("Jason Hollar", "CEO (Cardinal Health)", ""),
    "CVS Specialty": ("David Joyner", "CEO (CVS Health)", ""),
    "Aetna Specialty Pharmacy": ("David Joyner", "CEO (CVS Health)", ""),
    "Coram CVS Specialty Infusion": ("David Joyner", "CEO (CVS Health)", ""),
    "Walgreens Specialty Pharmacy": ("Tim Wentworth", "CEO (Walgreens Boots Alliance)", ""),
    "AllianceRx Walgreens Pharmacy": ("Tim Wentworth", "CEO (Walgreens Boots Alliance)", ""),
    "Accredo Health Group": ("David Cordani", "CEO (The Cigna Group)", ""),
    "Cigna Specialty Pharmacy / Accredo": ("David Cordani", "CEO (The Cigna Group)", ""),
    "Express Scripts Accredo network": ("David Cordani", "CEO (The Cigna Group)", ""),
    "Optum Specialty Pharmacy": ("Andrew Witty", "CEO (UnitedHealth Group)", ""),
    "Walmart Specialty Pharmacy": ("Doug McMillon", "CEO (Walmart)", ""),
    "Kroger Specialty Pharmacy": ("Ron Vail", "CEO (Kroger)", ""),
    "CenterWell Pharmacy": ("Jim Rechtin", "CEO (Humana)", ""),
    "Fisher Scientific": ("Marc N. Casper", "CEO (Thermo Fisher Scientific)", ""),
    "VWR International": ("Michael Stubblefield", "CEO (Avantor)", ""),
    "Anda Inc.": ("Richard Francis", "CEO (Teva)", ""),
    "Anda Pharmaceuticals": ("Richard Francis", "CEO (Teva)", ""),
    "Phoenix Pharma SE": ("Sven Fenchel", "CEO (PHOENIX group)", ""),
    "PHOENIX Pharmahandel": ("Sven Fenchel", "CEO (PHOENIX group)", ""),
    "Phoenix Healthcare Distribution": ("Sven Fenchel", "CEO (PHOENIX group)", ""),
    "Tamro": ("Sven Fenchel", "CEO (PHOENIX / Tamro)", ""),
    "NOWEDA Apothekergenossenschaft": ("Dr. Michael Kuck", "Vorstandsvorsitzender (CEO)", ""),
    "Hungaropharma": ("Dr. Feller Antal", "CEO", ""),
    "Nadro": ("Fernando Motilla", "CEO", ""),
    "Morris & Dickson Co., LLC": ("Paul Dickson", "CEO", ""),
    "Morris & Dickson Specialty Distribution": ("Paul Dickson", "CEO", ""),
    "BioCare Specialty Distribution (BioCareSD)": ("James Frary", "Board Chair", ""),
    "Onco360": ("Paul Jardina", "CEO", ""),
    "PANTHERx Rare": ("Curtis Bachmann", "CEO", ""),
    "Orsini Specialty Pharmacy": ("Dave Cervenak", "CEO", ""),
    "Sciensus": ("Russell Auty", "CEO", ""),
    "NewBridge Pharmaceuticals": ("Joe Henein", "CEO", ""),
    "Specialised Therapeutics Asia": ("Carlo Montagner", "CEO", ""),
    "Asia Biotech Group": ("Arun Prakash", "Founder", ""),
    "Cedarlane": ("Bill Clark", "President", ""),
    "Cedarlane Laboratories": ("Bill Clark", "President", ""),
    "Marken (UPS Healthcare)": ("Wes Wheeler", "CEO (Marken)", ""),
    "Mediq": ("Erik-Jan Bijvank", "CEO", ""),
    "CuraScript SD": ("Lisa Harrison", "SVP & President, Specialty Distribution & Solutions (Cencora)", ""),
    "US Bioservices": ("Lisa Harrison", "SVP & President, Specialty Distribution & Solutions (Cencora)", ""),
    "TheraCom Specialty Pharmacy": ("Lisa Harrison", "SVP & President, Specialty Distribution & Solutions (Cencora)", ""),
    "Alliance Healthcare UK": ("Robert Maione", "Cencora Alliance Healthcare leadership", ""),
    "Alloga UK": ("Lisa Harrison", "Cencora specialty / Alloga leadership", ""),
    "AAH Pharmaceuticals": ("Not publicly disclosed", "Joint MD / Executive leadership (post-2025 CEO transition)", ""),
}

PERSON_LI_RE = re.compile(
    r"https?://(?:[a-z]{2,3}\.)?linkedin\.com/in/[a-zA-Z0-9\-_%]+/?",
    re.I,
)
COMPANY_LI_RE = re.compile(
    r"https?://(?:[a-z]{2,3}\.)?linkedin\.com/company/[a-zA-Z0-9\-_%]+/?",
    re.I,
)

CEO_PATS = [
    re.compile(
        r"(?:CEO|Chief Executive Officer|Chief Executive|President(?:\s+and\s+CEO)?|"
        r"Managing Director|Group CEO|Vorstandsvorsitzender|"
        r"President and Chief Executive Officer)\s*[:\-–]?\s*"
        r"([A-Z][a-zA-Z.''\-]+(?:\s+[A-Z][a-zA-Z.''\-]+){1,3})",
        re.I,
    ),
    re.compile(
        r"([A-Z][a-zA-Z.''\-]+(?:\s+[A-Z][a-zA-Z.''\-]+){1,3})\s*,?\s*"
        r"(?:CEO|Chief Executive Officer|President(?:\s+and\s+CEO)?|Managing Director|Group CEO)\b",
        re.I,
    ),
    re.compile(
        r"([A-Z][a-zA-Z.''\-]+(?:\s+[A-Z][a-zA-Z.''\-]+){1,3})\s+has\s+been\s+(?:appointed|named)\s+"
        r"(?:as\s+)?(?:CEO|President|Managing Director)",
        re.I,
    ),
]


def empty(v: Any) -> bool:
    return str(v or "").strip().lower() in NPD


def looks_person(name: str, company: str = "") -> bool:
    n = (name or "").strip()
    if len(n) < 4 or len(n) > 60 or "\n" in n:
        return False
    low = n.lower()
    if low in NPD or "not publicly" in low or "disclosed" in low:
        return False
    bad = {
        "ceo", "cto", "inc", "ltd", "llc", "gmbh", "team", "the", "and", "leadership",
        "board", "group", "company", "president", "director", "officer", "our", "about",
        "contact", "privacy", "cookie", "home", "news", "careers", "products", "publicly",
        "disclosed", "managing", "joint",
    }
    if low in bad:
        return False
    if any(x in low for x in ("leadership", "joint managing", "jv leadership", "curriculum", "not in your")):
        return False
    parts = n.replace("-", " ").split()
    if not (2 <= len(parts) <= 4):
        return False
    if any(p.lower().strip(".,") in bad for p in parts):
        return False
    # require Title Case-ish: at least 2 capitalized tokens
    caps = sum(1 for p in parts if p[:1].isupper())
    if caps < 2:
        return False
    if not all(re.match(r"^[A-Za-z][A-Za-z.'\-]*$", p) for p in parts):
        return False
    if company and n.lower() == company.strip().lower():
        return False
    return True


def role_from_match(blob: str, name: str) -> str:
    # find role near name
    for role in (
        "Chief Executive Officer", "President and CEO", "Group CEO", "CEO",
        "Managing Director", "President", "Chairman", "Founder",
        "VP and General Manager", "General Manager", "Vorstandsvorsitzender",
    ):
        if re.search(rf"{re.escape(name)}.{{0,40}}{re.escape(role)}|{re.escape(role)}.{{0,40}}{re.escape(name)}", blob, re.I):
            return role
    if re.search(r"\bCEO\b", blob):
        return "CEO"
    if re.search(r"\bManaging Director\b", blob, re.I):
        return "Managing Director"
    if re.search(r"\bPresident\b", blob, re.I):
        return "President"
    return "CEO"


def strip_html(html: str) -> str:
    t = re.sub(r"(?is)<script.*?>.*?</script>", " ", html)
    t = re.sub(r"(?is)<style.*?>.*?</style>", " ", t)
    t = re.sub(r"<[^>]+>", " ", t)
    t = re.sub(r"&nbsp;|&amp;|&quot;|&#39;", " ", t)
    return re.sub(r"\s+", " ", t)


async def fetch(client: httpx.AsyncClient, url: str, timeout: float = 8.0) -> str:
    try:
        r = await client.get(url, timeout=timeout, follow_redirects=True)
        if r.status_code >= 400:
            return ""
        return r.text or ""
    except Exception:
        return ""


async def ddg_search(client: httpx.AsyncClient, query: str) -> str:
    html = await fetch(client, f"https://html.duckduckgo.com/html/?q={quote(query)}", timeout=10.0)
    return strip_html(html)[:5000] if html else ""


async def company_pages(client: httpx.AsyncClient, website: str) -> str:
    if not website.startswith("http"):
        return ""
    blobs = []
    base = website.rstrip("/")
    for path in ("", "/about", "/about-us", "/company", "/leadership", "/our-team", "/management", "/en/company", "/en/about"):
        url = base if not path else urljoin(base + "/", path.lstrip("/"))
        html = await fetch(client, url, timeout=6.0)
        if html:
            blobs.append(strip_html(html)[:2500])
        if len(blobs) >= 2:
            break
    return " ".join(blobs)[:6000]


def extract_contact(blob: str, company: str) -> dict[str, str]:
    out: dict[str, str] = {}
    if not blob:
        return out
    # person linkedin first
    for m in PERSON_LI_RE.finditer(blob):
        out["linkedin"] = m.group(0).rstrip("/")
        break
    for pat in CEO_PATS:
        m = pat.search(blob)
        if m:
            name = m.group(1).strip(" .,;")
            if looks_person(name, company):
                out["person"] = name
                out["role"] = role_from_match(blob[max(0, m.start() - 40) : m.end() + 60], name)
                break
    return out


async def enrich_one(client: httpx.AsyncClient, row: dict[str, str]) -> dict[str, str]:
    company = row["Company"]
    if company in KNOWN:
        p, role, li = KNOWN[company]
        return {"person": p, "role": role, "linkedin": li, "source": "known_web"}

    website = row.get("Website") or ""
    q1 = f'"{company}" (CEO OR "Managing Director" OR President OR "Chief Executive")'
    q2 = f'"{company}" CEO site:linkedin.com/in'
    site_blob, ddg1, ddg2 = await asyncio.gather(
        company_pages(client, website),
        ddg_search(client, q1),
        ddg_search(client, q2),
    )
    blob = " ".join([site_blob, ddg1, ddg2])
    found = extract_contact(blob, company)
    # if person LI missing, try targeted query with person name
    if found.get("person") and not found.get("linkedin"):
        ddg3 = await ddg_search(
            client,
            f'"{found["person"]}" "{company}" site:linkedin.com/in',
        )
        m = PERSON_LI_RE.search(ddg3)
        if m:
            found["linkedin"] = m.group(0).rstrip("/")
    found["source"] = "web_ddg_site"
    return found


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
    cws["A1"] = "Contact Person = named executive (CEO/MD/President preferred)"
    cws["A2"] = "Role = job title"
    cws["A3"] = "LinkedIn = person profile (linkedin.com/in/...), not company page when possible"
    cws["A4"] = f"Updated: {datetime.now(timezone.utc).isoformat()} | rows={len(rows)}"
    wb.save(path)


async def main() -> int:
    rows = load_rows(XLSX)
    print(f"Loaded {len(rows)}")
    if PROGRESS.exists():
        PROGRESS.unlink()

    # Apply known first
    for r in rows:
        if r["Company"] in KNOWN:
            p, role, li = KNOWN[r["Company"]]
            r["Contact Person"] = p
            r["Role"] = role
            # only set person LI; clear company LI if present without person
            if li:
                r["LinkedIn"] = li
            elif (r.get("LinkedIn") or "").find("/company/") >= 0:
                # keep company LI only if no person — user wants person; clear to NPD
                r["LinkedIn"] = "Not publicly disclosed"
            else:
                if empty(r.get("LinkedIn")):
                    r["LinkedIn"] = "Not publicly disclosed"

    sem = asyncio.Semaphore(6)
    limits = httpx.Limits(max_connections=8, max_keepalive_connections=6)
    filled_p = filled_r = filled_l = 0

    async with httpx.AsyncClient(
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            )
        },
        limits=limits,
        timeout=httpx.Timeout(12.0, connect=5.0),
    ) as http:

        async def one(i: int, row: dict[str, str]) -> tuple[int, dict[str, str]]:
            async with sem:
                if row["Company"] in KNOWN and looks_person(KNOWN[row["Company"]][0]):
                    p, role, li = KNOWN[row["Company"]]
                    return i, {"person": p, "role": role, "linkedin": li, "source": "known_web"}
                try:
                    ev = await asyncio.wait_for(enrich_one(http, row), timeout=25.0)
                except asyncio.TimeoutError:
                    ev = {}
                await asyncio.sleep(0.12)
                return i, ev

        # process in batches for progress logging
        for start in range(0, len(rows), 12):
            chunk_idx = list(range(start, min(start + 12, len(rows))))
            print(f"[{start+1}-{chunk_idx[-1]+1}/{len(rows)}] contacts...", flush=True)
            results = await asyncio.gather(*[one(i, rows[i]) for i in chunk_idx])
            for i, ev in results:
                row = rows[i]
                person = (ev.get("person") or "").strip()
                role = (ev.get("role") or "").strip()
                li = (ev.get("linkedin") or "").strip()

                # Prefer real person names
                if looks_person(person, row["Company"]):
                    row["Contact Person"] = person
                    row["Role"] = role or "CEO"
                elif row["Company"] in KNOWN:
                    p, role_k, li_k = KNOWN[row["Company"]]
                    row["Contact Person"] = p
                    row["Role"] = role_k
                    if li_k:
                        li = li_k
                else:
                    if empty(row.get("Contact Person")):
                        row["Contact Person"] = "Not publicly disclosed"
                    if empty(row.get("Role")):
                        row["Role"] = "Not publicly disclosed"

                # LinkedIn: person profile only
                if li and "/in/" in li:
                    row["LinkedIn"] = li.rstrip("/")
                else:
                    # if existing is company page, clear for person-requirement
                    cur = row.get("LinkedIn") or ""
                    if "/company/" in cur or empty(cur):
                        row["LinkedIn"] = "Not publicly disclosed"

                src = row.get("Data Sources") or ""
                if "adc_contact_web" not in src:
                    row["Data Sources"] = (src + "; adc_contact_web").strip("; ")

                if not empty(row.get("Contact Person")) and looks_person(row["Contact Person"]):
                    filled_p += 1
                if not empty(row.get("Role")) and row.get("Role", "").lower() != "not publicly disclosed":
                    filled_r += 1
                if not empty(row.get("LinkedIn")) and "/in/" in (row.get("LinkedIn") or ""):
                    filled_l += 1

                rec = {
                    "company": row["Company"],
                    "person": row.get("Contact Person"),
                    "role": row.get("Role"),
                    "linkedin": row.get("LinkedIn"),
                    "source": ev.get("source"),
                }
                with PROGRESS.open("a", encoding="utf-8") as pf:
                    pf.write(json.dumps(rec, ensure_ascii=False) + "\n")
                print(
                    f"  {row['Company'][:36]:36} | {str(row.get('Contact Person'))[:22]:22} | "
                    f"{str(row.get('Role'))[:24]:24} | LI={'yes' if '/in/' in (row.get('LinkedIn') or '') else 'no'}",
                    flush=True,
                )

    rows.sort(key=lambda x: (x.get("Company") or "").lower())
    write_xlsx(XLSX, rows)

    person_n = sum(1 for r in rows if looks_person(r.get("Contact Person") or ""))
    role_n = sum(1 for r in rows if not empty(r.get("Role")) and r.get("Role", "").lower() != "not publicly disclosed")
    li_n = sum(1 for r in rows if "/in/" in (r.get("LinkedIn") or ""))
    summary = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "rows": len(rows),
        "contact_person_named": person_n,
        "role_filled": role_n,
        "person_linkedin": li_n,
        "note": "LinkedIn column prefers person profiles (linkedin.com/in/). Company pages cleared when no person URL found.",
    }
    SUMMARY.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\nDONE named_contacts={person_n} roles={role_n} person_linkedin={li_n}")
    return 0


if __name__ == "__main__":
    raise SystemExit(asyncio.run(main()))
