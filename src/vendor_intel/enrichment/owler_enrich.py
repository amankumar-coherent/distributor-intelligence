"""Owler firmographic enrichment via LLM + search snippets (gap-fill only)."""
from __future__ import annotations

import asyncio
import json
import time
from typing import Any

from vendor_intel.export.fmcg_landscape import is_plausible_location


def _field_empty(crawl: dict, field: str) -> bool:
    return not str(crawl.get(field) or "").strip()


def apply_owler_to_row(row: dict, extracted: dict[str, Any]) -> list[str]:
    """Merge Owler fields into row only where _crawl fields are empty."""
    if "_crawl" not in row or not isinstance(row["_crawl"], dict):
        row["_crawl"] = {}
    crawl = row["_crawl"]
    updates: list[str] = []

    hq = extracted.get("headquarters")
    if hq and str(hq) != "null" and _field_empty(crawl, "headquarters"):
        if is_plausible_location(str(hq), allow_country_only=True):
            crawl["headquarters"] = str(hq)
            updates.append(f"HQ={hq}")

    founded = extracted.get("founded_year")
    if founded and str(founded) != "null" and _field_empty(crawl, "founded_year"):
        try:
            crawl["founded_year"] = int(founded)
            updates.append(f"Founded={founded}")
        except (TypeError, ValueError):
            pass

    emp = extracted.get("employee_count")
    if emp and str(emp) != "null" and _field_empty(crawl, "employee_count"):
        crawl["employee_count"] = str(emp)
        updates.append(f"Employees={emp}")

    rev = extracted.get("revenue")
    if rev and str(rev) != "null" and _field_empty(crawl, "revenue"):
        crawl["revenue"] = str(rev)
        updates.append(f"Turnover={rev}")

    offices = extracted.get("offices")
    if offices and str(offices) != "null" and _field_empty(crawl, "offices"):
        crawl["offices"] = str(offices)
        updates.append(f"Cities/Regions={offices}")

    if not updates:
        return updates

    raw = dict(row.get("_crawl_raw") or {})
    data = dict(raw.get("data") or {})
    company = dict(data.get("company") or {})
    location = dict(data.get("location") or {})
    financials = dict(data.get("financials") or {})

    if crawl.get("headquarters"):
        location["headquarters"] = crawl["headquarters"]
    if crawl.get("founded_year"):
        company["founded_year"] = crawl["founded_year"]
    if crawl.get("employee_count"):
        financials["employee_count"] = crawl["employee_count"]
    if crawl.get("revenue"):
        financials["revenue"] = crawl["revenue"]

    data["company"] = company
    data["location"] = location
    data["financials"] = financials
    raw["data"] = data
    row["_crawl_raw"] = raw

    src = str(row.get("enrichment_source") or "")
    if "owler" not in src:
        row["enrichment_source"] = f"{src}+owler".strip("+")

    return updates


async def _fetch_owler_fields(
    name: str,
    domain: str,
    *,
    client: Any,
    router: Any,
    model_name: str,
) -> dict[str, Any]:
    import os

    use_llm = (os.getenv("GAP_FILL_USE_LLM") or "true").strip().lower() not in (
        "0",
        "false",
        "no",
        "off",
    )
    res: dict[str, Any] = {}

    if use_llm and client is not None:
        system_prompt_direct = (
            "You are an expert data retrieval assistant. Your goal is to retrieve the firmographics "
            "for the given company as listed on their official Owler profile.\n"
            "CRITICAL: If you do not know the company, if it is a small/local business, or if you do not "
            "have highly certain, exact factual knowledge of its Owler profile details, you MUST return null "
            "for those fields.\n"
            "Do NOT guess, assume, or estimate. Zero hallucination is allowed. If not 100% factually certain, output null."
        )
        user_prompt_direct = f"""Company Name: {name}
Website Domain: {domain}

JSON Output Format (Strictly return ONLY this JSON object):
{{
  "headquarters": "headquarters location (City, State, Country or City, Country) from their Owler profile, or null if unknown",
  "founded_year": integer representing year founded from their Owler profile, or null if unknown,
  "employee_count": "number of employees or range (e.g. '100', '100-500', '1000+') from their Owler profile, or null if unknown",
  "revenue": "turnover / annual revenue (e.g. '$10M', '$1.2B') from their Owler profile, or null if unknown",
  "offices": "comma-separated list of other cities/regions or office locations from their Owler profile, or null if unknown"
}}
"""
        try:
            response = await client.chat.completions.create(
                model=model_name,
                messages=[
                    {"role": "system", "content": system_prompt_direct},
                    {"role": "user", "content": user_prompt_direct},
                ],
                response_format={"type": "json_object"},
                temperature=0.0,
                max_tokens=400,
            )
            content = response.choices[0].message.content or "{}"
            res = json.loads(content)
        except Exception as e:
            print(f"  [owler] Direct LLM lookup failed for {name}: {e}", flush=True)

    has_revenue = bool(res.get("revenue"))
    has_employees = bool(res.get("employee_count"))

    if not (has_revenue and has_employees):
        clean_domain = (
            domain.removeprefix("https://")
            .removeprefix("http://")
            .removeprefix("www.")
            .split("/")[0]
            if domain
            else ""
        )
        query = f"site:owler.com/company/ {name}"
        query_alt = (
            f"site:owler.com/company/ {name} {clean_domain}"
            if clean_domain
            else f"owler.com company {name}"
        )
        hits = []
        try:
            hits = await router.search(query, discovery_mode=True)
            if not hits and clean_domain:
                await asyncio.sleep(1.0)
                hits = await router.search(query_alt, discovery_mode=True)
        except Exception:
            pass

        if hits:
            lines = [f"- {h.title}: {h.snippet} ({h.link})" for h in hits[:5]]
            snippets_text = "\n".join(lines)

            if use_llm and client is not None:
                system_prompt_search = (
                    "You are an expert data extractor. Your task is to extract firmographics for a company "
                    "from the provided Owler search snippets.\n"
                    "CRITICAL: You must extract information ONLY if it is explicitly stated in the text of the snippets. "
                    "Do NOT use general knowledge, guess, or extrapolate. If a field is not explicitly mentioned "
                    "in the snippets, you must return null. Zero hallucination is allowed."
                )
                user_prompt_search = f"""Company Name: {name}
Website: {domain}

Search Snippets:
{snippets_text}

JSON Output Format (Strictly return ONLY this JSON object):
{{
  "headquarters": "headquarters location (City, State, Country or City, Country) if found, else null",
  "founded_year": integer representing year founded if found, else null,
  "employee_count": "number of employees or range (e.g. '100', '100-500', '1000+') if found, else null",
  "revenue": "turnover / annual revenue (e.g. '$10M', '$1.2B') if found, else null",
  "offices": "comma-separated list of other cities/regions or office locations mentioned, else null"
}}
"""
                try:
                    response = await client.chat.completions.create(
                        model=model_name,
                        messages=[
                            {"role": "system", "content": system_prompt_search},
                            {"role": "user", "content": user_prompt_search},
                        ],
                        response_format={"type": "json_object"},
                        temperature=0.0,
                        max_tokens=400,
                    )
                    search_res = json.loads(response.choices[0].message.content or "{}")
                    for k in ("headquarters", "founded_year", "employee_count", "revenue", "offices"):
                        if search_res.get(k) and not res.get(k):
                            res[k] = search_res[k]
                except Exception:
                    pass
            else:
                # Step-5 path: Owler via search snippets + heuristic extract (no OpenAI SDK)
                from vendor_intel.enrichment.gap_fill.heuristic_extract import (
                    heuristic_extract_fields,
                )

                heur = heuristic_extract_fields(snippets_text, company=name)
                for k in ("headquarters", "founded_year", "employee_count", "revenue", "offices"):
                    if heur.get(k) and not res.get(k):
                        res[k] = heur[k]

    return res


async def enrich_company_owler(
    row: dict,
    *,
    client: Any,
    router: Any,
    model_name: str,
    idx: int = 0,
    total: int = 0,
) -> bool:
    """Enrich one company from Owler. Returns True if any field was updated."""
    if row.get("owler_enriched"):
        return False

    name = str(row.get("company") or row.get("brand") or "Unknown")
    domain = str(row.get("domain") or row.get("website") or "")

    extracted = await _fetch_owler_fields(
        name, domain, client=client, router=router, model_name=model_name
    )
    updates = apply_owler_to_row(row, extracted)
    row["owler_enriched"] = True
    row["owler_enrich_timestamp"] = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())

    prefix = f"  [{idx}/{total}] " if idx else "  "
    if updates:
        print(f"{prefix}Updated {name} from Owler: {', '.join(updates)}", flush=True)
        return True
    print(f"{prefix}Searched Owler for {name} but found no new details.", flush=True)
    return False


async def enrich_companies_owler(
    rows: list[dict],
    *,
    concurrent: int = 10,
    settings: Any | None = None,
) -> tuple[int, int]:
    """Run Owler enrichment for all rows. Returns (updated_count, total)."""
    import os

    from backend.config import get_settings, sync_opencode_to_openai_env
    from openai import AsyncOpenAI

    from vendor_intel.clients.search_router import FreeSearchRouter
    from vendor_intel.config import Settings

    sync_opencode_to_openai_env()
    cfg = get_settings()
    if not cfg.openai_api_key:
        print("  [owler] OPENAI_API_KEY not set — skipping Owler enrichment.", flush=True)
        return 0, len(rows)

    settings = settings or Settings.load()
    client = AsyncOpenAI(api_key=cfg.openai_api_key)
    model_name = (os.getenv("OPENAI_MODEL") or cfg.model_extract or "gpt-4o-mini").strip()
    router = FreeSearchRouter(settings)

    sem = asyncio.Semaphore(max(1, concurrent))
    updated = 0
    total = len(rows)

    async def _one(row: dict, idx: int) -> None:
        nonlocal updated
        async with sem:
            if await enrich_company_owler(
                row, client=client, router=router, model_name=model_name, idx=idx, total=total
            ):
                updated += 1

    await asyncio.gather(*[_one(r, i) for i, r in enumerate(rows, 1)])
    return updated, total
