"""Search-snippet backfill for missing company fields (HQ, founded, employees, revenue)."""
from __future__ import annotations

import json
import os
from typing import Any

from vendor_intel.clients.search_router import FreeSearchRouter
from vendor_intel.config import Settings
from vendor_intel.export.fmcg_landscape import crawl_snapshot, is_plausible_location
from vendor_intel.enrichment.gap_fill.search_queries import queries_for_fields


def _needs_backfill(row: dict[str, Any]) -> bool:
    crawl = row.get("_crawl") or {}
    hq = str(crawl.get("headquarters") or crawl.get("contact_address") or "").strip()
    has_hq = bool(hq and is_plausible_location(hq, allow_country_only=True))
    founded = str(crawl.get("founded_year") or "").strip()
    employees = str(crawl.get("employee_count") or "").strip()
    return not has_hq or not founded or not employees


def _merge_backfill(row: dict[str, Any], extracted: dict[str, Any]) -> bool:
    if not extracted:
        return False
    crawl = dict(row.get("_crawl") or {})
    changed = False
    for src_key, dst_key in (
        ("headquarters", "headquarters"),
        ("founded_year", "founded_year"),
        ("employee_count", "employee_count"),
        ("revenue", "revenue"),
    ):
        val = str(extracted.get(src_key) or "").strip()
        if not val or str(crawl.get(dst_key) or "").strip():
            continue
        if dst_key == "headquarters" and not is_plausible_location(val, allow_country_only=True):
            continue
        crawl[dst_key] = val
        changed = True
    if not changed:
        return False
    row["_crawl"] = crawl
    raw = dict(row.get("_crawl_raw") or {})
    data = dict(raw.get("data") or {})
    company = dict(data.get("company") or {})
    location = dict(data.get("location") or {})
    financials = dict(data.get("financials") or {})
    if extracted.get("headquarters"):
        location["headquarters"] = extracted["headquarters"]
    if extracted.get("founded_year"):
        company["founded_year"] = extracted["founded_year"]
    if extracted.get("employee_count"):
        financials["employee_count"] = extracted["employee_count"]
    if extracted.get("revenue"):
        financials["revenue"] = extracted["revenue"]
    data["company"] = company
    data["location"] = location
    data["financials"] = financials
    raw["data"] = data
    row["_crawl_raw"] = raw
    src = str(row.get("enrichment_source") or "")
    if "search_backfill" not in src:
        row["enrichment_source"] = f"{src}+search_backfill".strip("+")
    return True


def _backfill_use_llm() -> bool:
    raw = (os.getenv("GAP_FILL_USE_LLM") or "true").strip().lower()
    return raw not in ("0", "false", "no", "off")


async def _extract_from_snippets(company: str, domain: str, snippets: str) -> dict[str, Any]:
    if not snippets.strip():
        return {}
    if not _backfill_use_llm():
        from vendor_intel.enrichment.gap_fill.heuristic_extract import heuristic_extract_fields

        return heuristic_extract_fields(snippets, company=company)

    from backend.config import get_settings, sync_opencode_to_openai_env
    from openai import AsyncOpenAI

    sync_opencode_to_openai_env()
    cfg = get_settings()
    if not cfg.openai_api_key:
        from vendor_intel.enrichment.gap_fill.heuristic_extract import heuristic_extract_fields

        return heuristic_extract_fields(snippets, company=company)
    client = AsyncOpenAI(api_key=cfg.openai_api_key)
    model = (os.getenv("OPENAI_MODEL") or cfg.model_extract or "gpt-4o-mini").strip()
    prompt = (
        f"Company: {company}\nDomain: {domain}\n\n"
        "From the search snippets below, extract ONLY facts explicitly stated. "
        "Return JSON with keys: headquarters, founded_year, employee_count, revenue. "
        "Use empty string for unknown fields. Do not invent data.\n\n"
        f"{snippets[:6000]}"
    )
    try:
        resp = await client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"},
            temperature=0,
        )
        text = (resp.choices[0].message.content or "").strip()
        data = json.loads(text) if text else {}
        return data if isinstance(data, dict) else {}
    except Exception:
        from vendor_intel.enrichment.gap_fill.heuristic_extract import heuristic_extract_fields

        return heuristic_extract_fields(snippets, company=company)


async def backfill_row(row: dict[str, Any], router: FreeSearchRouter) -> bool:
    """Fill missing HQ/founded/employees from web search snippets."""
    if not _needs_backfill(row):
        return False
    company = str(row.get("company") or row.get("brand") or "").strip()
    domain = str(row.get("domain") or row.get("website") or "").strip().lower()
    domain = domain.removeprefix("https://").removeprefix("http://").removeprefix("www.").split("/")[0]
    if not company:
        return False

    crawl = row.get("_crawl") or {}
    fields: list[str] = []
    if not str(crawl.get("founded_year") or "").strip():
        fields.append("founded_year")
    hq = str(crawl.get("headquarters") or crawl.get("contact_address") or "").strip()
    if not (hq and is_plausible_location(hq, allow_country_only=True)):
        fields.append("headquarters")
    if not str(crawl.get("employee_count") or "").strip():
        fields.append("employee_count")
    if not fields:
        fields = ["founded_year", "headquarters", "employee_count"]

    queries = queries_for_fields(company, fields, domain=domain, max_queries=3)
    lines: list[str] = []
    for q in queries:
        hits = await router.search(q, discovery_mode=True)
        for h in (hits or [])[:5]:
            lines.append(f"- [{q}] {h.title}: {h.snippet} ({h.link})")
    if not lines:
        return False
    extracted = await _extract_from_snippets(company, domain, "\n".join(lines))
    return _merge_backfill(row, extracted)


async def backfill_companies(
    result: dict[str, Any],
    *,
    limit: int | None = None,
    settings: Settings | None = None,
) -> tuple[int, int]:
    """Backfill rows in result that lack plausible HQ or key firmographics."""
    from vendor_intel.utils.scrape_logging import alog, log

    settings = settings or Settings.load()
    router = FreeSearchRouter(settings)
    rows = [r for r in (result.get("relevant_companies") or []) if r.get("is_relevant", True)]
    todo = [r for r in rows if _needs_backfill(r)]
    if limit is not None:
        todo = todo[: max(0, limit)]
    log(f"Search backfill queue: {len(todo)} companies")
    ok = 0
    total = len(todo)
    for i, row in enumerate(todo, 1):
        name = str(row.get("company") or row.get("brand") or "")[:50]
        await alog(f"[backfill {i}/{total}] {name}")
        try:
            if await backfill_row(row, router):
                ok += 1
                crawl = crawl_snapshot(row.get("_crawl_raw"))
                row["_crawl"] = {**dict(row.get("_crawl") or {}), **crawl}
                hq = str(crawl.get("headquarters") or "")[:60]
                await alog(f"[backfill {i}/{total}] updated {name} | HQ={hq or '—'}")
        except Exception as exc:
            await alog(f"[backfill {i}/{total}] failed {name}: {exc}", level="ERROR")
    log(f"Search backfill done: {ok} updated, {total - ok} unchanged/failed")
    return ok, len(todo) - ok
