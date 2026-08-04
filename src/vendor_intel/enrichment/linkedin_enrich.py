"""LinkedIn /about page fetch + LLM extraction for company firmographics."""
from __future__ import annotations

import json
import os
import re
from typing import Any

_SUFFIXES = re.compile(
    r"\b(corp(oration)?|inc(orporated)?|ltd|limited|co(mpany)?|gmbh|plc|s\.?a\.?|a\.?g\.?)\b",
    re.IGNORECASE,
)


def slugify_company_name(name: str) -> str:
    s = name.strip().lower()
    s = re.sub(r"[^a-z0-9\s-]", "", s)
    s = re.sub(r"[\s-]+", "-", s)
    return s


def get_linkedin_urls(name: str) -> list[str]:
    slugs: list[str] = []
    s1 = slugify_company_name(name)
    if s1:
        slugs.append(s1)
    cleaned_name = _SUFFIXES.sub("", name).strip()
    s2 = slugify_company_name(cleaned_name)
    if s2 and s2 != s1:
        slugs.append(s2)
    return [f"https://linkedin.com/company/{slug}/about" for slug in slugs]


async def fetch_linkedin_page(url: str) -> str:
    try:
        from crawler.smart_crawl import stealthy_get

        html = await stealthy_get(url)
        if html and len(html) > 500 and "signup" not in html.lower() and "join-form" not in html.lower():
            return html
    except Exception as e:
        print(f"  [linkedin] stealthy_get error: {e}", flush=True)

    try:
        from crawler.smart_crawl import playwright_get

        html = await playwright_get(url)
        if html and len(html) > 500:
            return html
    except Exception as e:
        print(f"  [linkedin] playwright_get error: {e}", flush=True)

    # Agent Reach Jina Reader — public LinkedIn company /about pages (no login MCP required)
    try:
        from vendor_intel.integrations.page_fetch import fetch_page_text

        text, _ = await fetch_page_text(url, max_chars=15000)
        if text and len(text) > 200 and "signup" not in text.lower():
            print(f"  [linkedin] Jina Reader OK for {url[:60]}", flush=True)
            return text
    except Exception as e:
        print(f"  [linkedin] Jina fallback error: {e}", flush=True)
    return ""


async def extract_linkedin_info(company_name: str, html: str) -> dict[str, Any]:
    from backend.config import get_settings, sync_opencode_to_openai_env
    from openai import AsyncOpenAI

    sync_opencode_to_openai_env()
    cfg = get_settings()
    if not cfg.openai_api_key or not html.strip():
        return {}
    client = AsyncOpenAI(api_key=cfg.openai_api_key)
    model = (os.getenv("OPENAI_MODEL") or cfg.model_extract or "gpt-4o-mini").strip()

    try:
        from bs4 import BeautifulSoup

        soup = BeautifulSoup(html, "html.parser")
        for script in soup(["script", "style"]):
            script.decompose()
        text = soup.get_text(separator="\n")
        text = "\n".join(line.strip() for line in text.splitlines() if line.strip())
    except Exception:
        text = html[:30000]

    prompt = (
        f"Company: {company_name}\n\n"
        "From the LinkedIn /about page text below, extract the following information:\n"
        "1. Founded Year (as string, e.g. '2015')\n"
        "2. Headquarters Location (as string, e.g. 'San Francisco, CA')\n"
        "3. Total Employees (as string, e.g. '11-50 employees' or '250')\n"
        "4. About/Overview (a detailed description/summary of the company, e.g. 1-2 paragraphs)\n\n"
        "Return a JSON object with keys: founded_year, headquarters, employee_count, about_description.\n"
        "Use empty string for unknown fields. Explicitly do not guess or hallucinate.\n\n"
        f"{text[:15000]}"
    )
    try:
        resp = await client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"},
            temperature=0,
        )
        raw_content = resp.choices[0].message.content or ""
        return json.loads(raw_content)
    except Exception as e:
        print(f"  [linkedin] LLM extraction error: {e}", flush=True)
        return {}


async def generate_polished_summary(
    company_name: str, website_text: str, linkedin_text: str
) -> str:
    from backend.config import get_settings, sync_opencode_to_openai_env
    from openai import AsyncOpenAI

    sync_opencode_to_openai_env()
    cfg = get_settings()
    if not cfg.openai_api_key:
        return ""

    if not website_text.strip() and not linkedin_text.strip():
        return ""

    base_url = (os.getenv("OPENAI_BASE_URL") or "").strip() or None
    client = AsyncOpenAI(api_key=cfg.openai_api_key, base_url=base_url)
    model = (os.getenv("OPENAI_MODEL") or cfg.model_extract or "gpt-4o-mini").strip()

    prompt = (
        f"Company Name: {company_name}\n"
        f"Website Scraped Data: {website_text[:8000]}\n"
        f"LinkedIn About Data: {linkedin_text[:8000]}\n\n"
        "Using the information provided above, write a professional, coherent 2-3 sentence summary of the company. "
        "Focus on what the company does, its core products/services, and its target markets. "
        "Make the sentences polished, clear, and flow naturally. Do not mention the source of the data "
        "(e.g. 'scraped website' or 'LinkedIn page'). "
        "If the information is highly limited, write a brief, professional summary using only the available facts. "
        "Do not invent any facts."
    )
    try:
        resp = await client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3,
        )
        return (resp.choices[0].message.content or "").strip()
    except Exception as e:
        print(f"  [summary] OpenAI polish summary error: {e}", flush=True)
        parts = [website_text, linkedin_text]
        return " ".join([p.strip() for p in parts if p.strip()])[:500]


def apply_linkedin_and_summary(
    row: dict,
    linkedin_data: dict,
    linkedin_url: str,
    polished_summary: str,
) -> None:
    raw = row.get("_crawl_raw") or {"data": {}}
    if not isinstance(raw, dict):
        raw = {"data": {}}
    data = raw.get("data") or {}
    if not isinstance(data, dict):
        data = {}

    company = dict(data.get("company") or {})
    location = dict(data.get("location") or {})
    financials = dict(data.get("financials") or {})
    contact = dict(data.get("contact") or {})
    crawl = dict(row.get("_crawl") or {})

    if linkedin_data.get("founded_year") and not str(crawl.get("founded_year") or "").strip():
        company["founded_year"] = linkedin_data["founded_year"]
        crawl["founded_year"] = linkedin_data["founded_year"]
    if linkedin_data.get("headquarters") and not str(crawl.get("headquarters") or "").strip():
        location["headquarters"] = linkedin_data["headquarters"]
        crawl["headquarters"] = linkedin_data["headquarters"]
    if linkedin_data.get("employee_count") and not str(crawl.get("employee_count") or "").strip():
        financials["employee_count"] = linkedin_data["employee_count"]
        crawl["employee_count"] = linkedin_data["employee_count"]
    if linkedin_url and not str(crawl.get("linkedin") or row.get("linkedin") or "").strip():
        contact["linkedin"] = linkedin_url
        crawl["linkedin"] = linkedin_url
        row["linkedin"] = linkedin_url

    data["company"] = company
    data["location"] = location
    data["financials"] = financials
    data["contact"] = contact
    raw["data"] = data
    row["_crawl_raw"] = raw

    from vendor_intel.export.fmcg_landscape import crawl_snapshot

    row["_crawl"] = {**crawl_snapshot(raw), **{k: v for k, v in crawl.items() if v}}

    if polished_summary:
        row["polished_summary"] = polished_summary
        row["company_summary"] = polished_summary

    src = str(row.get("enrichment_source") or "")
    if "linkedin" not in src:
        row["enrichment_source"] = f"{src}+linkedin".strip("+")
    row["linkedin_enriched"] = True


async def enrich_company_linkedin(
    row: dict,
    *,
    smart_data: dict | None = None,
    idx: int = 0,
    total: int = 0,
    log_fn=None,
) -> None:
    """Fetch LinkedIn /about, extract firmographics, apply gap-fill to row."""
    name = str(row.get("company") or row.get("brand") or "")
    if not name:
        return

    async def _log(msg: str, level: str = "INFO") -> None:
        if log_fn:
            await log_fn(msg, level=level)
        else:
            print(msg, flush=True)

    existing_linkedin = str(
        row.get("linkedin") or (row.get("_crawl") or {}).get("linkedin") or ""
    ).strip()
    if existing_linkedin and "linkedin.com/company/" in existing_linkedin:
        base = existing_linkedin.split("?")[0].rstrip("/")
        linkedin_url = base + "/about" if not base.endswith("/about") else base
        urls_to_try = [linkedin_url]
    else:
        urls_to_try = get_linkedin_urls(name)

    linkedin_html = ""
    fetched_url = ""
    linkedin_data: dict[str, Any] = {}
    linkedin_about_text = ""

    # Agent Reach LinkedIn MCP — structured company profile + URL resolution
    try:
        from vendor_intel.integrations.linkedin_mcp import enrich_company_via_mcp, linkedin_mcp_available

        if linkedin_mcp_available():
            existing_li = str(
                row.get("linkedin") or (row.get("_crawl") or {}).get("linkedin") or ""
            ).strip()
            mcp_profile, mcp_url = await enrich_company_via_mcp(name, linkedin_url=existing_li)
            if mcp_profile:
                linkedin_data = {
                    k: v
                    for k, v in mcp_profile.items()
                    if k in ("founded_year", "headquarters", "employee_count", "about_description")
                    and v
                }
                linkedin_about_text = str(mcp_profile.get("about_description") or "")
                if mcp_url:
                    fetched_url = mcp_url.replace("/about", "").rstrip("/")
                if idx:
                    await _log(f"[{idx}/{total}] LinkedIn MCP profile OK for {name[:40]}")
    except Exception as e:
        if idx:
            await _log(f"[{idx}/{total}] LinkedIn MCP skipped: {e}", level="WARN")

    if not linkedin_data or not linkedin_about_text:
        for url in urls_to_try:
            if idx:
                await _log(f"[{idx}/{total}] Fetching LinkedIn: {url}")
            linkedin_html = await fetch_linkedin_page(url)
            if linkedin_html:
                if not fetched_url:
                    fetched_url = url.replace("/about", "")
                break

        if linkedin_html:
            if idx:
                await _log(f"[{idx}/{total}] Extracting LinkedIn details...")
            extracted = await extract_linkedin_info(name, linkedin_html)
            for key, val in extracted.items():
                if val and not linkedin_data.get(key):
                    linkedin_data[key] = val
            if not linkedin_about_text:
                linkedin_about_text = linkedin_data.get("about_description") or ""

    website_summary_text = ""
    if smart_data:
        from vendor_intel.pipeline.csv_fields import extract_company_summary

        website_summary_text = extract_company_summary(smart_data, max_len=4000)

    if idx:
        await _log(f"[{idx}/{total}] Polishing summary...")
    polished_summary = await generate_polished_summary(name, website_summary_text, linkedin_about_text)
    apply_linkedin_and_summary(row, linkedin_data, fetched_url, polished_summary)
