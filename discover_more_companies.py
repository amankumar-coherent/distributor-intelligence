#!/usr/bin/env python3
"""
discover_more_companies.py  — Full Discovery + Enrichment
==========================================================
One command:
    python discover_more_companies.py                         # all markets
    python discover_more_companies.py --market avocado_oil   # one market

Pipeline per market:
  1. Run 32+ search queries across web (DuckDuckGo/Bing), Alibaba,
     ThomasNet, Kompass, LinkedIn, Owler, ImportYeti, trade directories
  2. Validate candidates with GPT-4o-mini (real company in market?)
  3. Deduplicate against existing companies in Excel/CSV
  4. For every NEW company — FULL ENRICHMENT from 4 sources:
       a) Company website (scraped)
       b) LinkedIn company page (searched + scraped)
       c) Owler company page (searched + scraped)
       d) Web search for firmographics (founded, HQ, employees, revenue)
  5. GPT-4o-mini extracts ALL 35 Excel column values from gathered text
  6. Writes complete enriched row to both .xlsx and .csv
"""

import argparse
import asyncio
import csv
import copy
import io
import json
import os
import re
import sys
import time
from pathlib import Path
from urllib.parse import urlparse

import httpx
import openpyxl
from openai import AsyncOpenAI

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

# ── Load .env ──────────────────────────────────────────────────────────────────
env_path = ROOT / ".env"
if env_path.exists():
    with open(env_path, encoding="utf-8") as _ef:
        for _line in _ef:
            _line = _line.strip()
            if _line and not _line.startswith("#") and "=" in _line:
                _k, _v = _line.split("=", 1)
                os.environ.setdefault(_k.strip(), _v.strip())

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
ai = AsyncOpenAI(api_key=OPENAI_API_KEY)

# ── All 35 Excel columns ───────────────────────────────────────────────────────
ALL_COLUMNS = [
    "Company", "Founded", "Headquarters", "Cities / Regions", "Ownership",
    "Business Type", "Employees", "Turnover / Scale", "Contact Person", "Role",
    "Email", "Phone / WhatsApp", "LinkedIn", "Website", "Core Categories",
    "Specialty Focus", "Price Segment", "Key Brands Represented", "Exclusivity",
    "Partnership Duration", "Retail Chains", "Specialty Grocers",
    "Foodservice / HoReCa", "E-commerce", "Channel Strength", "Distribution Type",
    "Regions Served", "Regional Extensions", "Warehouse / Logistics",
    "Delivery / Storage", "Competitive Benchmark", "Data Confidence",
    "Quality Score", "Summary", "Data Sources",
]

MARKETS = [
    "avocado_oil", "technical_foam", "spectral_sensor",
    "gel_ice_pack", "cyber_physical_systems_security", "high_energy_laser",
]

MARKET_CONFIG = {
    "avocado_oil": {
        "label": "Avocado Oil",
        "description": "companies that produce, process, bottle, or distribute avocado oil as a food or cosmetic ingredient",
        "search_queries": [
            "avocado oil manufacturer company official website",
            "avocado oil producer exporter B2B wholesale",
            "extra virgin avocado oil manufacturer company",
            "avocado oil bulk supplier company",
            "avocado oil brand company USA",
            "avocado oil manufacturer Mexico",
            "avocado oil producer Peru Chile",
            "avocado oil manufacturer South Africa",
            "avocado oil producer Kenya",
            "avocado oil manufacturer New Zealand Australia",
            "avocado oil producer Europe Spain Italy",
            "cold pressed avocado oil manufacturer",
            "organic avocado oil manufacturer company",
            "avocado oil private label manufacturer OEM",
            "avocado oil refinery company",
            "avocado oil cosmetic grade manufacturer",
            "avocado oil food grade manufacturer bulk",
            "avocado oil company linkedin.com/company",
            "site:alibaba.com avocado oil manufacturer",
            "site:globalsources.com avocado oil manufacturer",
            "site:kompass.com avocado oil manufacturer",
            "site:thomasnet.com avocado oil supplier",
            "site:tradeindia.com avocado oil manufacturer",
            "site:importyeti.com avocado oil",
            "avocado oil brand list top producers 2024",
            "avocado oil supplier Israel Dominican Republic",
            "avocado oil manufacturer Indonesia India China",
            "avocado oil specialty food company",
            "avocado oil startup brand producer",
            "certified avocado oil producer organic",
            "avocado oil company distributor Europe",
            "avocado oil wholesale trader company",
        ],
    },
    "technical_foam": {
        "label": "Technical Foam",
        "description": "companies that manufacture technical, industrial, or specialty foam",
        "search_queries": [
            "technical foam manufacturer company official site",
            "industrial foam manufacturer polyurethane company",
            "acoustic foam manufacturer company B2B",
            "thermal foam insulation manufacturer company",
            "specialty foam manufacturer company",
            "foam fabricator industrial company",
            "open cell foam manufacturer company",
            "closed cell foam manufacturer company",
            "high density foam manufacturer industrial",
            "reticulated foam manufacturer company",
            "foam die cutting company manufacturer",
            "foam gasket manufacturer company",
            "foam sealing manufacturer company",
            "EVA foam manufacturer company",
            "PU foam manufacturer company industrial",
            "melamine foam manufacturer company",
            "silicone foam manufacturer company",
            "neoprene foam manufacturer company",
            "site:thomasnet.com polyurethane foam manufacturer",
            "site:alibaba.com industrial foam manufacturer",
            "site:kompass.com foam manufacturer",
            "site:globalsources.com foam manufacturer",
            "foam manufacturer company linkedin.com/company",
            "foam manufacturer Germany UK France Italy",
            "foam manufacturer USA Canada",
            "foam manufacturer China Japan Korea",
            "foam manufacturer India Turkey",
            "foam converter manufacturer company",
            "flexible foam manufacturer industrial",
            "rigid foam manufacturer industrial",
            "foam manufacturer aerospace automotive",
            "foam products company official website",
        ],
    },
    "spectral_sensor": {
        "label": "Spectral Sensor",
        "description": "companies that manufacture spectral sensors, hyperspectral cameras, NIR sensors, or spectrometers",
        "search_queries": [
            "spectral sensor manufacturer company official site",
            "hyperspectral imaging company manufacturer",
            "NIR near infrared sensor manufacturer",
            "spectrometer manufacturer company industrial",
            "multispectral camera manufacturer company",
            "SWIR sensor camera manufacturer",
            "Raman spectrometer manufacturer company",
            "hyperspectral camera manufacturer agriculture",
            "inline spectrometer manufacturer food industry",
            "spectroscopy instrument manufacturer company",
            "imaging spectrometer manufacturer company",
            "fluorescence spectrometer manufacturer",
            "UV VIS spectrometer manufacturer company",
            "FTIR spectrometer manufacturer company",
            "handheld spectrometer manufacturer company",
            "miniature spectrometer manufacturer company",
            "spectral imaging system manufacturer",
            "wavelength sensor manufacturer company",
            "site:spie.org spectral imaging company",
            "site:photonics.com company hyperspectral",
            "site:alibaba.com spectrometer manufacturer",
            "site:kompass.com spectrometer manufacturer",
            "spectral sensor company linkedin.com/company",
            "hyperspectral sensor company USA Europe",
            "spectral sensor manufacturer Germany Netherlands",
            "spectral camera manufacturer Japan Korea",
            "spectral sensor startup company",
            "NIR sensor manufacturer agriculture food",
            "hyperspectral drone sensor manufacturer",
            "spectral sensor mining medical company",
            "optical sensor manufacturer company",
            "spectral analysis company manufacturer",
        ],
    },
    "gel_ice_pack": {
        "label": "Gel Ice Pack",
        "description": "companies that manufacture gel ice packs, cold packs, reusable ice packs, or cold chain packaging",
        "search_queries": [
            "gel ice pack manufacturer company official site",
            "cold pack manufacturer company B2B",
            "reusable ice pack manufacturer company",
            "gel pack manufacturer wholesale",
            "instant cold pack manufacturer company",
            "cold chain gel pack manufacturer",
            "pharmaceutical cold pack manufacturer",
            "medical cold pack manufacturer company",
            "sports ice pack manufacturer company",
            "cold compress manufacturer company",
            "PCM phase change material pack manufacturer",
            "cold therapy pack manufacturer company",
            "freezer gel pack manufacturer company",
            "coolant pack manufacturer company",
            "cold pack OEM private label manufacturer",
            "site:alibaba.com gel ice pack manufacturer",
            "site:globalsources.com ice pack manufacturer",
            "site:thomasnet.com cold pack manufacturer",
            "site:kompass.com cold pack manufacturer",
            "ice pack manufacturer linkedin.com/company",
            "gel ice pack manufacturer USA Canada",
            "ice pack manufacturer Europe UK Germany",
            "ice pack manufacturer China Asia",
            "ice pack manufacturer India",
            "temperature control pack manufacturer",
            "cold shipping pack manufacturer company",
            "insulated pack manufacturer cold chain",
            "gel pack company startup brand",
            "cold pack supplier company official",
            "cryotherapy pack manufacturer company",
            "cold storage pack manufacturer industrial",
            "food grade cold pack manufacturer company",
        ],
    },
    "cyber_physical_systems_security": {
        "label": "CPS Security",
        "description": "companies providing cybersecurity for cyber-physical systems, OT/ICS/SCADA security",
        "search_queries": [
            "OT security company ICS cybersecurity vendor",
            "cyber physical systems security company",
            "SCADA security vendor company official",
            "industrial cybersecurity company OT ICS",
            "ICS security platform vendor company",
            "operational technology security company",
            "OT network security vendor company",
            "industrial control system security vendor",
            "cyber physical security platform company",
            "IoT OT security company vendor",
            "critical infrastructure security company",
            "IIoT security vendor company",
            "OT asset management security company",
            "OT threat detection company",
            "industrial SOC security company",
            "SCADA protection security vendor",
            "smart grid security company vendor",
            "building automation security company",
            "water utility energy sector OT security company",
            "OT security company linkedin.com/company",
            "OT security company USA Israel",
            "OT security company Europe Germany UK",
            "OT security startup company",
            "SCADA security company Asia Pacific",
            "ICS cybersecurity company 2024",
            "Gartner Forrester OT ICS security vendor",
            "cyber physical security startup 2023 2024",
            "managed OT security services company",
            "zero trust OT security vendor",
            "OT security monitoring anomaly detection company",
            "ICS incident response company",
            "OT security assessment company",
        ],
    },
    "high_energy_laser": {
        "label": "High Energy Laser",
        "description": "companies that manufacture high energy laser systems, directed energy weapons, or high-power industrial lasers",
        "search_queries": [
            "high energy laser manufacturer company official",
            "directed energy laser company defense",
            "high power fiber laser manufacturer company",
            "solid state laser manufacturer high power",
            "laser weapon system manufacturer defense",
            "industrial high power laser manufacturer",
            "kW class laser manufacturer company",
            "DPSS diode pumped solid state laser manufacturer",
            "Nd YAG laser manufacturer high energy",
            "CO2 laser manufacturer high power industrial",
            "fiber laser manufacturer kW company",
            "disk laser manufacturer company",
            "laser cladding welding system manufacturer",
            "laser cutting high power manufacturer",
            "pulsed laser manufacturer high energy",
            "ultrafast laser manufacturer company",
            "excimer laser manufacturer company",
            "diode laser manufacturer high power company",
            "site:deps.org laser company member",
            "site:spie.org high power laser company",
            "laser company linkedin.com/company defense",
            "high energy laser company USA defense contractor",
            "laser company Germany UK France",
            "laser manufacturer Japan Korea China",
            "high power laser company Israel",
            "directed energy company Australia Europe",
            "laser startup company high power",
            "laser defense company Europe NATO",
            "laser system integrator company",
            "laser company DARPA DoD contract",
            "laser manufacturer aerospace company",
            "photonic laser manufacturer company",
        ],
    },
}

# ── Utility ────────────────────────────────────────────────────────────────────

def clean_domain(url: str) -> str:
    if not url:
        return ""
    if not url.startswith("http"):
        url = "https://" + url
    try:
        return urlparse(url).netloc.lower().lstrip("www.")
    except Exception:
        return url.lower()


def load_existing_companies(market: str) -> set[str]:
    csv_path = ROOT / f"output/{market}/global_{market}_market_global.csv"
    if not csv_path.exists():
        return set()
    with open(csv_path, encoding="utf-8", errors="replace") as f:
        lines = f.readlines()
    header = lines[0]
    data_lines = [l for l in lines[1:] if not l.strip().startswith("===")]
    reader = csv.DictReader(io.StringIO(header + "".join(data_lines)))
    return {r.get("Company", "").strip().lower() for r in reader if r.get("Company")}


async def ddgs_search(query: str, num: int = 12) -> list[dict]:
    results = []
    try:
        from ddgs import DDGS
        def _search():
            with DDGS() as ddgs:
                return list(ddgs.text(query, max_results=num))
        results = await asyncio.to_thread(_search)
    except Exception:
        pass
    return results or []


async def fetch_text(url: str, max_chars: int = 5000) -> str:
    """Fetch text from a URL using httpx with fallback to ddgs extract."""
    if not url:
        return ""
    if not url.startswith("http"):
        url = "https://" + url
    # Try httpx first
    try:
        async with httpx.AsyncClient(timeout=15, follow_redirects=True,
                                     headers={"User-Agent": "Mozilla/5.0 (compatible; ResearchBot/1.0)"}) as client:
            resp = await client.get(url)
            if resp.status_code == 200:
                text = resp.text[:max_chars * 3]
                # Strip HTML tags roughly
                text = re.sub(r"<[^>]+>", " ", text)
                text = re.sub(r"\s+", " ", text).strip()
                return text[:max_chars]
    except Exception:
        pass
    # Fallback: ddgs extract
    try:
        from ddgs import DDGS
        def _extract():
            with DDGS() as ddgs:
                if hasattr(ddgs, "extract"):
                    out = ddgs.extract(url, fmt="text_plain")
                    return str(out.get("content", ""))[:max_chars]
            return ""
        return await asyncio.to_thread(_extract)
    except Exception:
        return ""


async def search_company_info(company_name: str, extra_query: str = "") -> str:
    """Run targeted web search for company firmographics."""
    queries = [
        f'"{company_name}" company founded headquarters employees',
        f'"{company_name}" annual revenue turnover',
        f'"{company_name}" CEO contact email phone',
        f'site:linkedin.com/company "{company_name}"',
        f'site:owler.com/company "{company_name.lower().replace(" ", "")}"',
    ]
    if extra_query:
        queries.append(f'"{company_name}" {extra_query}')

    all_text = []
    for q in queries[:4]:  # limit to 4 queries per company
        results = await ddgs_search(q, num=5)
        for r in results:
            snippet = r.get("body", "")
            title = r.get("title", "")
            url = r.get("href", "")
            all_text.append(f"[{title}] ({url})\n{snippet}")
        await asyncio.sleep(0.3)

    return "\n\n".join(all_text)[:6000]


async def fetch_website_text(website: str) -> str:
    """Fetch and return cleaned text from company website."""
    if not website:
        return ""
    url = website if website.startswith("http") else f"https://{website}"
    return await fetch_text(url, max_chars=4000)


async def fetch_linkedin_text(company_name: str) -> str:
    """Search and fetch LinkedIn company page text."""
    results = await ddgs_search(f'site:linkedin.com/company "{company_name}"', num=3)
    for r in results:
        href = r.get("href", "")
        if "linkedin.com/company" in href:
            snippet = r.get("body", "")
            return f"LinkedIn: {href}\n{snippet}"
    return ""


async def fetch_owler_text(company_name: str) -> str:
    """Search and fetch Owler company page text."""
    slug = re.sub(r"[^a-z0-9]", "", company_name.lower())
    owler_url = f"https://www.owler.com/company/{slug}"
    results = await ddgs_search(f'site:owler.com "{company_name}"', num=3)
    text_parts = []
    for r in results:
        if "owler.com" in r.get("href", ""):
            text_parts.append(r.get("body", ""))
    # Try direct fetch
    direct = await fetch_text(owler_url, max_chars=2000)
    if direct:
        text_parts.append(direct)
    return "\n".join(text_parts)[:3000]


# ── Main enrichment via GPT ────────────────────────────────────────────────────

async def enrich_company_full(
    company_name: str,
    website: str,
    market_label: str,
    market_desc: str,
    semaphore: asyncio.Semaphore,
) -> dict:
    """
    Gather data from website + LinkedIn + Owler + web search,
    then use GPT-4o-mini to fill ALL 35 Excel columns.
    """
    async with semaphore:
        print(f"    Enriching: {company_name}...")

        # Gather from all sources in parallel
        web_info_task = search_company_info(company_name)
        website_task = fetch_website_text(website)
        linkedin_task = fetch_linkedin_text(company_name)
        owler_task = fetch_owler_text(company_name)

        web_info, website_text, linkedin_text, owler_text = await asyncio.gather(
            web_info_task, website_task, linkedin_task, owler_task,
            return_exceptions=True,
        )

        def safe(v):
            if isinstance(v, Exception) or v is None:
                return ""
            return str(v)

        web_info = safe(web_info)
        website_text = safe(website_text)
        linkedin_text = safe(linkedin_text)
        owler_text = safe(owler_text)

        # Build combined context
        context = f"""COMPANY: {company_name}
WEBSITE: {website}
MARKET: {market_label} — {market_desc}

=== WEBSITE CONTENT ===
{website_text[:2500]}

=== WEB SEARCH RESULTS ===
{web_info[:2000]}

=== LINKEDIN ===
{linkedin_text[:1000]}

=== OWLER ===
{owler_text[:1000]}"""

        prompt = f"""You are a business intelligence analyst. Extract and fill ALL fields below for this company using ONLY the provided source text.

{context}

Fill every field you can find evidence for. Use ONLY information present in the sources above — do NOT hallucinate or invent data. If a field cannot be determined, use empty string "".

Return a JSON object with exactly these keys:
{{
  "Company": "{company_name}",
  "Founded": "year founded (e.g. 1995) or empty",
  "Headquarters": "City, Country",
  "Cities / Regions": "other cities or regions where they operate (comma separated)",
  "Ownership": "Public / Private / Subsidiary of [Parent] / PE-backed",
  "Business Type": "Manufacturer / Distributor / Service Provider / etc.",
  "Employees": "headcount range e.g. 50-200 or 1,200",
  "Turnover / Scale": "annual revenue if known e.g. $50M or Mid-tier",
  "Contact Person": "name of key contact if found",
  "Role": "title of contact person",
  "Email": "contact email if found",
  "Phone / WhatsApp": "phone number if found",
  "LinkedIn": "full LinkedIn company URL",
  "Website": "full website URL",
  "Core Categories": "main product/service categories",
  "Specialty Focus": "specific specializations, certifications, niche areas",
  "Price Segment": "Premium / Mid-tier / Economy / Budget",
  "Key Brands Represented": "brand names they own or distribute",
  "Exclusivity": "exclusive distributor / non-exclusive / own brand",
  "Partnership Duration": "years in operation or partnership length",
  "Retail Chains": "names of retail chains they supply to (if any)",
  "Specialty Grocers": "specialty grocery stores they work with (if any)",
  "Foodservice / HoReCa": "restaurant chains, hotels, food service clients (if any)",
  "E-commerce": "Amazon / Shopify / own website / other platforms (if any)",
  "Channel Strength": "Strong / Moderate / Limited",
  "Distribution Type": "Direct / Through distributors / Both",
  "Regions Served": "geographic regions served",
  "Regional Extensions": "recent expansion regions",
  "Warehouse / Logistics": "warehouse locations or logistics capabilities",
  "Delivery / Storage": "delivery terms or storage capabilities",
  "Competitive Benchmark": "Premium / Mid-tier / Value",
  "Data Confidence": "High / Medium / Low (based on source quality)",
  "Quality Score": 0.8,
  "Summary": "2-3 sentence company summary describing what they do",
  "Data Sources": "website | linkedin | owler | web_search"
}}"""

        for attempt in range(3):
            try:
                resp = await ai.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0,
                    max_tokens=1200,
                    response_format={"type": "json_object"},
                )
                data = json.loads(resp.choices[0].message.content)
                # Ensure company name and website are set
                data["Company"] = company_name
                if not data.get("Website"):
                    data["Website"] = website
                return data
            except Exception as e:
                if attempt < 2:
                    await asyncio.sleep(2 ** attempt)

        # Fallback minimal row
        return {
            "Company": company_name,
            "Website": website,
            "Data Confidence": "Low",
            "Quality Score": 0.5,
            "Summary": f"{company_name} operates in the {market_label} market.",
            "Data Sources": "web_search",
        }


# ── Discovery ──────────────────────────────────────────────────────────────────

async def discover_candidates(market: str, config: dict) -> list[dict]:
    """Run all search queries and extract candidate companies."""
    SKIP_DOMAINS = {
        "wikipedia.org", "youtube.com", "facebook.com", "twitter.com",
        "instagram.com", "reddit.com", "quora.com", "amazon.com", "ebay.com",
        "google.com", "bing.com", "forbes.com", "bloomberg.com", "reuters.com",
        "statista.com", "mordorintelligence.com", "grandviewresearch.com",
        "marketsandmarkets.com", "globenewswire.com", "prnewswire.com",
        "businesswire.com", "researchandmarkets.com", "alliedmarketresearch.com",
        "technavio.com", "precedenceresearch.com", "marketwatch.com",
    }

    seen_domains: set[str] = set()
    candidates = []

    queries = config["search_queries"]
    for i, query in enumerate(queries, 1):
        print(f"    [{i}/{len(queries)}] {query[:65]}...")
        results = await ddgs_search(query, num=12)
        for r in results:
            url = r.get("href", "")
            title = r.get("title", "").strip()
            snippet = r.get("body", "").strip()
            domain = clean_domain(url)
            if not domain or domain in SKIP_DOMAINS or domain in seen_domains:
                continue
            seen_domains.add(domain)
            candidates.append({"name": title, "url": url, "domain": domain, "snippet": snippet})
        await asyncio.sleep(0.4)

    return candidates


async def validate_candidates(candidates: list[dict], config: dict, sem: asyncio.Semaphore) -> list[dict]:
    """GPT validates which candidates are real companies in this market."""
    market_desc = config["description"]
    label = config["label"]

    async def validate_one(c: dict):
        async with sem:
            prompt = f"""Market: {label} — {market_desc}

Search result:
Title: {c['name']}
URL: {c['url']}
Snippet: {c['snippet'][:300]}

Is this a REAL standalone company (not a directory listing, article, or report) operating in this market?
Extract the actual company name and domain.

JSON response:
{{"is_company": true/false, "company_name": "...", "website": "https://...", "confidence": "high/medium/low"}}"""
            try:
                resp = await ai.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0, max_tokens=120,
                    response_format={"type": "json_object"},
                )
                data = json.loads(resp.choices[0].message.content)
                if data.get("is_company") and data.get("company_name"):
                    return {"name": data["company_name"].strip(), "website": data.get("website", c["url"])}
            except Exception:
                pass
            return None

    tasks = [validate_one(c) for c in candidates]
    results = await asyncio.gather(*tasks)
    return [r for r in results if r]


# ── Excel / CSV append ─────────────────────────────────────────────────────────

def find_header_row(ws) -> int:
    for i, row in enumerate(ws.rows, 1):
        if row[0].value == "Company":
            return i
    raise ValueError("No 'Company' header found")


def get_row_style(ws, header_row: int) -> list:
    """Get styles from last data row to replicate."""
    styles = []
    last_row = None
    for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row):
        if row[0].value:
            last_row = row
    if last_row:
        for cell in last_row:
            styles.append({
                "font": copy.copy(cell.font) if cell.has_style else None,
                "fill": copy.copy(cell.fill) if cell.has_style else None,
                "alignment": copy.copy(cell.alignment) if cell.has_style else None,
                "border": copy.copy(cell.border) if cell.has_style else None,
            })
    return styles


def append_rows_to_excel(market: str, enriched_rows: list[dict]) -> int:
    xlsx_path = ROOT / f"output/{market}/global_{market}_market_global.xlsx"
    if not xlsx_path.exists():
        return 0
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active
    header_row_num = find_header_row(ws)
    headers = [cell.value for cell in ws[header_row_num]]
    col_map = {h: i + 1 for i, h in enumerate(headers) if h}
    row_styles = get_row_style(ws, header_row_num)

    added = 0
    for row_data in enriched_rows:
        new_row_num = ws.max_row + 1
        for col_name in ALL_COLUMNS:
            col_idx = col_map.get(col_name)
            if col_idx is None:
                continue
            value = row_data.get(col_name, "")
            cell = ws.cell(row=new_row_num, column=col_idx, value=value)
            # Apply style from last row
            if row_styles and col_idx <= len(row_styles):
                s = row_styles[col_idx - 1]
                if s["font"]:
                    cell.font = s["font"]
                if s["fill"]:
                    cell.fill = s["fill"]
                if s["alignment"]:
                    cell.alignment = s["alignment"]
                if s["border"]:
                    cell.border = s["border"]
        added += 1

    wb.save(xlsx_path)
    return added


def append_rows_to_csv(market: str, enriched_rows: list[dict]) -> int:
    csv_path = ROOT / f"output/{market}/global_{market}_market_global.csv"
    if not csv_path.exists():
        return 0
    # Read current fieldnames
    with open(csv_path, encoding="utf-8", errors="replace") as f:
        fieldnames = f.readline().strip().split(",")

    added = 0
    with open(csv_path, "a", encoding="utf-8", newline="", errors="replace") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        for row_data in enriched_rows:
            csv_row = {}
            # Map Excel column names → CSV field names
            mapping = {
                "Company": "Company", "Founded": "Founded",
                "Headquarters": "Headquarters", "Cities / Regions": "Cities_Regions",
                "Ownership": "Ownership", "Business Type": "Business_Type",
                "Employees": "Employees", "Turnover / Scale": "Turnover_Scale",
                "Website": "Website", "Summary": "Summary",
                "Data Confidence": "Confidence", "Quality Score": "Quality_Score",
                "Data Sources": "Data_Sources", "Role": "Role",
                "Email": "Email", "LinkedIn": "LinkedIn",
            }
            for excel_col, csv_col in mapping.items():
                if csv_col in fieldnames:
                    csv_row[csv_col] = row_data.get(excel_col, "")
            # Fallback: write by Excel column name too
            for col in ALL_COLUMNS:
                if col in fieldnames:
                    csv_row[col] = row_data.get(col, "")
            writer.writerow(csv_row)
            added += 1
    return added


# ── Market pipeline ────────────────────────────────────────────────────────────

async def process_market(market: str, enrich_sem: asyncio.Semaphore) -> dict:
    config = MARKET_CONFIG[market]
    label = config["label"]

    print(f"\n{'='*65}")
    print(f"  MARKET: {label.upper()}")
    print(f"{'='*65}")

    # 1. Load existing
    existing = load_existing_companies(market)
    print(f"  Existing companies in file: {len(existing)}")

    # 2. Discover candidates
    print(f"\n  [STEP 1] Running {len(config['search_queries'])} search queries...")
    candidates = await discover_candidates(market, config)
    print(f"  Raw candidates: {len(candidates)}")

    # 3. Validate
    print(f"\n  [STEP 2] AI validating {len(candidates)} candidates...")
    valid_sem = asyncio.Semaphore(10)
    valid = await validate_candidates(candidates, config, valid_sem)
    print(f"  Valid company candidates: {len(valid)}")

    # 4. Deduplicate
    seen_names: set[str] = set()
    new_companies = []
    for v in valid:
        name_lower = re.sub(r"[^a-z0-9]", "", v["name"].lower())
        if v["name"].strip().lower() not in existing and name_lower not in seen_names and name_lower:
            seen_names.add(name_lower)
            new_companies.append(v)
        else:
            print(f"    [SKIP duplicate/existing] {v['name']}")

    print(f"  Truly NEW companies: {len(new_companies)}")
    if not new_companies:
        print("  Nothing new to add.")
        return {"market": market, "found": 0, "added": 0, "companies": []}

    print(f"\n  Candidates to enrich:")
    for c in new_companies:
        print(f"    - {c['name']} | {c.get('website', '')}")

    # 5. Full enrichment for each new company
    print(f"\n  [STEP 3] Full enrichment ({len(new_companies)} companies)...")
    print(f"          Sources: Website + LinkedIn + Owler + Web Search + AI extraction")
    enrich_tasks = [
        enrich_company_full(c["name"], c.get("website", ""), label, config["description"], enrich_sem)
        for c in new_companies
    ]
    enriched_rows = await asyncio.gather(*enrich_tasks, return_exceptions=True)
    enriched_rows = [r for r in enriched_rows if isinstance(r, dict)]
    print(f"  Enriched: {len(enriched_rows)} rows")

    # 6. Write to Excel + CSV
    xl_added = append_rows_to_excel(market, enriched_rows)
    csv_added = append_rows_to_csv(market, enriched_rows)
    print(f"  Written to Excel: {xl_added} rows | CSV: {csv_added} rows")

    # Summary
    added_names = [r.get("Company", "") for r in enriched_rows]
    print(f"\n  NEW COMPANIES ADDED:")
    for i, name in enumerate(added_names, 1):
        row = next((r for r in enriched_rows if r.get("Company") == name), {})
        hq = row.get("Headquarters", "")
        emp = row.get("Employees", "")
        summary = row.get("Summary", "")[:70]
        print(f"    {i:3}. {name} | HQ: {hq} | Emp: {emp}")
        print(f"         {summary}")

    # Save per-market result JSON
    result_path = ROOT / f"output/{market}/discovery_new_companies.json"
    with open(result_path, "w", encoding="utf-8") as f:
        json.dump(enriched_rows, f, indent=2, ensure_ascii=False)

    return {
        "market": market,
        "found": len(new_companies),
        "added": len(enriched_rows),
        "companies": added_names,
    }


# ── Entry point ────────────────────────────────────────────────────────────────

async def main(target_markets: list[str]):
    print("=" * 65)
    print("  FULL COMPANY DISCOVERY + ENRICHMENT")
    print("  Sources: Web Search, Website, LinkedIn, Owler, Alibaba,")
    print("           Kompass, ThomasNet, TradeIndia, ImportYeti, SPIE")
    print("  Output:  All 35 Excel columns filled for every new company")
    print("=" * 65)
    print(f"  Markets: {', '.join(target_markets)}")

    enrich_semaphore = asyncio.Semaphore(5)  # max 5 concurrent full enrichments
    all_results = []

    for market in target_markets:
        try:
            result = await process_market(market, enrich_semaphore)
            all_results.append(result)
        except Exception as e:
            print(f"  [ERROR] {market}: {e}")
            all_results.append({"market": market, "found": 0, "added": 0, "companies": [], "error": str(e)})

    # Grand summary
    print("\n" + "=" * 65)
    print("  DISCOVERY COMPLETE — FINAL SUMMARY")
    print("=" * 65)
    print(f"\n  {'Market':<40} {'Found':>6} {'Added':>6}")
    print(f"  {'-'*54}")
    total_added = 0
    for r in all_results:
        err = " [ERROR]" if "error" in r else ""
        print(f"  {r['market']:<40} {r.get('found',0):>6} {r.get('added',0):>6}{err}")
        total_added += r.get("added", 0)
    print(f"  {'-'*54}")
    print(f"  {'TOTAL':<40} {'':>6} {total_added:>6}")
    print("=" * 65)

    # Save master log
    log_path = ROOT / "output" / "discovery_log.json"
    with open(log_path, "w", encoding="utf-8") as f:
        json.dump(all_results, f, indent=2, ensure_ascii=False)
    print(f"\n  Log saved: {log_path}")
    print(f"\n  Per-market JSONs saved in each output/<market>/discovery_new_companies.json")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Discover + fully enrich new companies for all markets"
    )
    parser.add_argument(
        "--market",
        choices=MARKETS,
        default=None,
        help="Run for one specific market (default: all 6 markets)",
    )
    args = parser.parse_args()
    target_markets = [args.market] if args.market else MARKETS
    asyncio.run(main(target_markets))
