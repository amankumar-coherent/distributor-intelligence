"""Post-enrich: verify / fill Founded + HQ from LinkedIn MCP company profiles."""
from __future__ import annotations

import asyncio
import re
from typing import Any

from vendor_intel.enrichment.gap_fill.gaps import (
    has_founded,
    has_headquarters,
    has_linkedin,
    is_plausible_location,
    norm,
    sync_crawl_from_raw,
)
from vendor_intel.integrations.linkedin_mcp import (
    enrich_company_via_mcp,
    linkedin_mcp_available,
    linkedin_mcp_delay,
)


def _year_token(value: Any) -> str:
    m = re.search(r"(19|20)\d{2}", str(value or ""))
    return m.group(0) if m else norm(value)


def _hq_normalize(value: str) -> str:
    return re.sub(r"\s+", " ", norm(value).lower())


def _hq_matches(a: str, b: str) -> bool:
    aa, bb = _hq_normalize(a), _hq_normalize(b)
    if not aa or not bb:
        return False
    return aa == bb or aa in bb or bb in aa


def apply_linkedin_firmographics(
    row: dict[str, Any],
    profile: dict[str, Any],
    *,
    linkedin_url: str = "",
) -> list[str]:
    """Merge LinkedIn MCP firmographics; fill gaps and verify/correct existing values.

    Also fills Cities/Regions (offices), company About text, and Key Brands when
    LinkedIn about clearly names brands/products (never invents).
    """
    if not profile:
        return []

    crawl = dict(row.get("_crawl") or {})
    updates: list[str] = []
    verified: dict[str, bool] = dict(crawl.get("linkedin_verified") or {})
    company = norm(row.get("company") or row.get("brand"))

    li_url = norm(linkedin_url or profile.get("linkedin_url"))
    if li_url and "linkedin.com" in li_url and not has_linkedin(row):
        crawl["linkedin"] = li_url.split("?")[0].rstrip("/")
        row["linkedin"] = crawl["linkedin"]
        updates.append("linkedin")

    # Founded year — fill if missing, correct if MCP differs
    li_founded = _year_token(profile.get("founded_year"))
    cur_founded = _year_token(crawl.get("founded_year"))
    if li_founded:
        if not cur_founded:
            crawl["founded_year"] = li_founded
            updates.append("founded_year")
            verified["founded_year"] = True
        elif cur_founded != li_founded:
            crawl["founded_year_previous"] = cur_founded
            crawl["founded_year"] = li_founded
            updates.append("founded_year_corrected")
            verified["founded_year"] = True
        else:
            verified["founded_year"] = True

    # Headquarters — fill if missing, correct if MCP differs
    li_hq = norm(profile.get("headquarters"))
    cur_hq = norm(crawl.get("headquarters") or crawl.get("contact_address"))
    if li_hq and is_plausible_location(li_hq, allow_country_only=True):
        if not cur_hq:
            crawl["headquarters"] = li_hq
            updates.append("headquarters")
            verified["headquarters"] = True
        elif not _hq_matches(cur_hq, li_hq):
            crawl["headquarters_previous"] = cur_hq
            crawl["headquarters"] = li_hq
            updates.append("headquarters_corrected")
            verified["headquarters"] = True
        else:
            verified["headquarters"] = True

    emp = norm(profile.get("employee_count"))
    if emp and not norm(crawl.get("employee_count")):
        crawl["employee_count"] = emp
        updates.append("employee_count")

    # Cities / Regions from LinkedIn Locations — only if distinct from HQ
    # (do not copy headquarters into Cities / Regions).
    cities = norm(profile.get("cities_regions") or profile.get("offices"))
    hq_now = norm(crawl.get("headquarters") or crawl.get("contact_address") or li_hq)
    if cities and is_plausible_location(cities, allow_country_only=True):
        from vendor_intel.export.fmcg_landscape import _places_equivalent, _LOCATION_REJECT_KW

        if _LOCATION_REJECT_KW.search(cities):
            cities = ""
        elif hq_now and _places_equivalent(cities, hq_now):
            cities = ""
        if cities:
            if not norm(crawl.get("cities_regions")):
                crawl["cities_regions"] = cities
                updates.append("cities_regions")
            offices = norm(profile.get("offices") or cities)
            if offices and not (hq_now and _places_equivalent(offices, hq_now)):
                if not norm(crawl.get("offices")):
                    crawl["offices"] = offices
                    updates.append("offices")

    # Company-written About / Overview from LinkedIn
    about = norm(profile.get("about_description"))
    if about and len(about) >= 40:
        existing_about = norm(crawl.get("about") or row.get("summary") or row.get("company_summary"))
        if not existing_about or len(existing_about) < 40:
            crawl["about"] = about[:2000]
            crawl["linkedin_about"] = about[:2000]
            if not norm(row.get("summary")):
                row["summary"] = about[:500]
            if isinstance(row.get("verdict"), dict) and not norm(
                row["verdict"].get("company_summary") or row["verdict"].get("summary")
            ):
                row["verdict"]["company_summary"] = about[:800]
            updates.append("about")
        elif not norm(crawl.get("linkedin_about")):
            crawl["linkedin_about"] = about[:2000]
            updates.append("linkedin_about")

    specs = norm(profile.get("specialties"))
    if specs and not norm(crawl.get("certifications") or crawl.get("industries")):
        crawl["industries"] = specs[:500]
        updates.append("specialties")

    # Key Brands Represented — only when LinkedIn about clearly names brands/products
    # that are not just the company name (align with Excel Key Brands Represented).
    kb_raw = norm(profile.get("key_brands_from_about"))
    if kb_raw and not norm(crawl.get("key_brands_represented")):
        from vendor_intel.pipeline.csv_fields import sanitize_key_brands

        brand = norm(row.get("brand"))
        kb = sanitize_key_brands(kb_raw, company=company, brand=brand)
        if kb and kb.lower() != company.lower():
            crawl["key_brands_represented"] = kb
            crawl["key_brands_source"] = "linkedin_mcp_about"
            updates.append("key_brands_represented")

    if verified:
        crawl["linkedin_verified"] = verified
        crawl["firmographic_source_linkedin_mcp"] = True

    if updates or verified:
        row["_crawl"] = crawl
        if updates:
            sync_crawl_from_raw(row)
    return updates


def needs_linkedin_mcp_firmographics(row: dict[str, Any]) -> bool:
    """True when Founded/HQ missing, unverified, or other MCP-fillable columns missing."""
    if not row.get("is_relevant", True):
        return False
    crawl = row.get("_crawl") or {}
    # Already attempted MCP with a finished response (not timeout) — skip on resume.
    if crawl.get("linkedin_mcp_attempted") and not crawl.get("linkedin_mcp_timeout"):
        return False
    verified = crawl.get("linkedin_verified") or {}
    if not has_founded(row) or not has_headquarters(row):
        return True
    if not verified.get("founded_year") or not verified.get("headquarters"):
        return True
    if not has_linkedin(row):
        return True
    if not norm(crawl.get("employee_count")):
        return True
    # Also re-run when cities / about / key brands still empty (LinkedIn can fill these)
    if not norm(crawl.get("offices") or crawl.get("cities_regions")):
        return True
    if not norm(crawl.get("about") or crawl.get("linkedin_about") or row.get("summary")):
        return True
    if not norm(crawl.get("key_brands_represented")):
        return True
    return False


async def verify_row_firmographics_via_mcp(row: dict[str, Any]) -> list[str]:
    if not linkedin_mcp_available():
        return []
    name = norm(row.get("company") or row.get("brand"))
    if not name:
        return []
    crawl = row.get("_crawl") or {}
    existing_li = norm(crawl.get("linkedin") or row.get("linkedin"))
    profile, li_url = await enrich_company_via_mcp(name, linkedin_url=existing_li)
    return apply_linkedin_firmographics(row, profile, linkedin_url=li_url)


async def verify_firmographics_for_market(
    rows: list[dict[str, Any]],
    *,
    concurrent: int = 1,
    verify_all: bool = True,
    persist_fn: Any = None,
) -> dict[str, Any]:
    """Verify / fill Founded + HQ via LinkedIn MCP (skips rows already MCP-verified).

    persist_fn: optional callback() invoked after each company so Ctrl+C resume
    only re-queues remaining rows (JSON must be saved by the caller).
    """
    if not linkedin_mcp_available():
        return {"skipped": True, "reason": "linkedin_mcp_unavailable", "queued": 0}

    delay = linkedin_mcp_delay()
    try:
        # One company may need search + profile (2 browser scrapes). Default 45s
        # caused every row to TIMEOUT — skipped. Allow ~2× call timeout.
        call_timeout = float(__import__("os").getenv("LINKEDIN_MCP_VERIFY_TIMEOUT_SEC", "300"))
    except ValueError:
        call_timeout = 300.0
    call_timeout = max(90.0, call_timeout)

    if verify_all:
        todo = [r for r in rows if needs_linkedin_mcp_firmographics(r)]
    else:
        todo = [
            r for r in rows
            if r.get("is_relevant", True) and (not has_founded(r) or not has_headquarters(r))
        ]

    skipped_verified = sum(
        1 for r in rows
        if r.get("is_relevant", True) and not needs_linkedin_mcp_firmographics(r)
    )

    if not todo:
        return {
            "queued": 0,
            "updated": 0,
            "verified_both": 0,
            "skipped_already_verified": skipped_verified,
            "errors": 0,
            "timeouts": 0,
            "delay_sec": delay,
        }

    print(
        f"  [linkedin_verify] queue={len(todo)} "
        f"(skip already done={skipped_verified}) "
        f"timeout={call_timeout:.0f}s delay~{delay:.1f}s",
        flush=True,
    )

    sem = asyncio.Semaphore(max(1, min(concurrent, 1)))  # serial — MCP hangs under concurrency
    updated = 0
    verified_both = 0
    errors = 0
    timeouts = 0

    def _checkpoint() -> None:
        if persist_fn is None:
            return
        try:
            persist_fn()
        except Exception as exc:
            print(f"  [linkedin_verify] checkpoint save failed: {exc}", flush=True)

    async def _one(row: dict[str, Any], idx: int) -> None:
        nonlocal updated, verified_both, errors, timeouts
        async with sem:
            name = str(row.get("company") or "?")[:50]
            print(f"  [linkedin_verify] starting [{idx}/{len(todo)}] {name}", flush=True)
            try:
                applied = await asyncio.wait_for(
                    verify_row_firmographics_via_mcp(row),
                    timeout=call_timeout,
                )
                if applied:
                    updated += 1
                    crawl = dict(row.get("_crawl") or {})
                    crawl["linkedin_mcp_attempted"] = True
                    crawl.pop("linkedin_mcp_timeout", None)
                    row["_crawl"] = crawl
                    print(
                        f"  [linkedin_verify] [{idx}/{len(todo)}] {name} +{applied}",
                        flush=True,
                    )
                else:
                    # Mark attempt so resume does not re-queue forever with zero progress
                    crawl = dict(row.get("_crawl") or {})
                    crawl["linkedin_mcp_attempted"] = True
                    crawl.pop("linkedin_mcp_timeout", None)
                    row["_crawl"] = crawl
                    print(
                        f"  [linkedin_verify] [{idx}/{len(todo)}] {name} (no updates)",
                        flush=True,
                    )
                v = (row.get("_crawl") or {}).get("linkedin_verified") or {}
                if v.get("founded_year") and v.get("headquarters"):
                    verified_both += 1
            except asyncio.TimeoutError:
                timeouts += 1
                crawl = dict(row.get("_crawl") or {})
                crawl["linkedin_mcp_timeout"] = True
                row["_crawl"] = crawl
                print(
                    f"  [linkedin_verify] [{idx}/{len(todo)}] {name} TIMEOUT after {call_timeout:.0f}s — skipped",
                    flush=True,
                )
            except Exception as exc:
                errors += 1
                print(f"  [linkedin_verify] [{idx}/{len(todo)}] {name} ERROR: {exc}", flush=True)
            _checkpoint()
            await asyncio.sleep(delay)

    # Sequential gather via semaphore=1 so one hang cannot block the whole batch forever
    for i, r in enumerate(todo, 1):
        await _one(r, i)

    return {
        "queued": len(todo),
        "updated": updated,
        "verified_both": verified_both,
        "skipped_already_verified": skipped_verified,
        "errors": errors,
        "timeouts": timeouts,
        "delay_sec": delay,
    }
