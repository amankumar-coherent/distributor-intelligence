"""Correct wrong company / brand / domain using web search + LinkedIn + OpenAI.

Runs in post-market pipeline before Excel export. Writes fixes back to the row JSON.
"""
from __future__ import annotations

import json
import os
import re
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

from vendor_intel.enrichment.apollo_contacts import DOMAIN_OVERRIDES
from vendor_intel.enrichment.gap_fill.gaps import domain_from_row, norm
from vendor_intel.enrichment.gap_fill.sources import _search_snippets
from vendor_intel.utils.domain_corrections import crawl_host_for_domain, resolve_best_domain

# Directory / aggregator domains — not official company sites.
_BAD_DOMAIN_FRAGMENTS = (
    "cbinsights.com",
    "zoominfo.com",
    "leadiq.com",
    "crunchbase.com",
    "dnb.com",
    "linkedin.com",
    "facebook.com",
    "twitter.com",
    "wikipedia.org",
    "bloomberg.com",
    "reuters.com",
    "marketwatch.com",
    "glassdoor.com",
    "indeed.com",
    "yelp.com",
    "trustpilot.com",
    "amazon.com",
    "alibaba.com",
    "indiamart.com",
    "tradeindia.com",
    "zaubacorp.com",
    "oilmonster.com",
    "industry-plaza.com",
    "simplybook.me",
    "slideserve.com",
    "office.com",
    "azureedge.net",
)

_LINKEDIN_COMPANY = re.compile(r"linkedin\.com/company/([a-zA-Z0-9_-]+)", re.I)


def _normalize_domain(raw: str) -> str:
    s = norm(raw).lower()
    if not s:
        return ""
    if "@" in s:
        s = s.split("@", 1)[1]
    if "://" in s:
        s = urlparse(s if "://" in s else f"https://{s}").netloc or s
    return s.removeprefix("www.").split("/")[0].strip(".")


def is_bad_domain(domain: str) -> bool:
    d = _normalize_domain(domain)
    if not d:
        return True
    if d in DOMAIN_OVERRIDES:
        return True
    return any(b in d for b in _BAD_DOMAIN_FRAGMENTS)


def _linkedin_url_from_snippets(snippets: str) -> str:
    m = _LINKEDIN_COMPANY.search(snippets or "")
    if not m:
        return ""
    return f"https://www.linkedin.com/company/{m.group(1)}"


async def gather_brand_evidence(
    row: dict[str, Any],
    router: Any,
    *,
    key_brands_only: bool = False,
    market_query: str = "",
) -> str:
    company = norm(row.get("company") or row.get("brand"))
    domain = domain_from_row(row)
    parts: list[str] = []
    crawl = row.get("_crawl") if isinstance(row.get("_crawl"), dict) else {}

    # Prefer already-fetched LinkedIn / crawl text first (no network hang)
    for label, key in (
        ("LinkedIn about", "linkedin_about"),
        ("LinkedIn about", "about"),
        ("Company summary", "summary"),
    ):
        text = norm(crawl.get(key) if key != "summary" else (crawl.get(key) or row.get("summary")))
        if text and len(text) >= 40:
            parts.append(f"{label}:\n{text[:2500]}")
            break

    products = norm(row.get("key_products") or crawl.get("products"))
    if products:
        parts.append(f"Stored products:\n{products[:1200]}")

    crawl_text = norm(crawl.get("crawl_text"))
    if crawl_text:
        parts.append(f"Local website crawl:\n{crawl_text[:2500]}")

    if key_brands_only:
        market = norm(market_query) or "this market"
        queries = [
            f'"{company}" products platforms solutions {market}',
        ]
        if domain and not is_bad_domain(domain):
            queries.append(f'"{company}" site:{domain} products solutions')
        else:
            queries.append(f'"{company}" product portfolio solutions brands')
    else:
        queries = [
            f'"{company}" official website company',
            f'site:linkedin.com/company "{company}"',
        ]
        if domain and not is_bad_domain(domain):
            queries.append(f'"{company}" site:{domain} about')
        elif domain:
            queries.append(f'"{company}" official site -site:linkedin.com')

    # Cap search time so a dead DDGS backend cannot hang the whole batch
    try:
        search_timeout = float(os.getenv("KEY_BRANDS_SEARCH_TIMEOUT_SEC", "45"))
    except ValueError:
        search_timeout = 45.0
    search_timeout = max(15.0, search_timeout)

    import asyncio

    for q in queries:
        try:
            snippets = await asyncio.wait_for(
                _search_snippets(router, q, max_hits=6),
                timeout=search_timeout,
            )
        except asyncio.TimeoutError:
            print(f"  [search] TIMEOUT after {search_timeout:.0f}s for {q[:70]!r}", flush=True)
            continue
        except Exception as exc:
            print(f"  [search] error for {q[:50]!r}: {exc}", flush=True)
            continue
        if snippets:
            parts.append(f"Query: {q}\n{snippets}")

    return "\n\n".join(parts)[:8000]


async def _llm_key_brands_only(
    row: dict[str, Any],
    evidence: str,
    *,
    market_query: str = "",
) -> dict[str, str]:
    if not evidence.strip():
        return {}
    from backend.config import get_settings, sync_opencode_to_openai_env
    from openai import AsyncOpenAI

    sync_opencode_to_openai_env()
    cfg = get_settings()
    if not cfg.openai_api_key:
        return {}

    company = norm(row.get("company") or row.get("brand"))
    role = norm(row.get("role"))
    products = norm(row.get("key_products"))

    client = AsyncOpenAI(api_key=cfg.openai_api_key)
    model = (os.getenv("OPENAI_MODEL") or cfg.model_extract or "gpt-4o-mini").strip()

    prompt = f"""You extract Key Brands Represented for a B2B channel-partner market intelligence export.

MARKET: {market_query or "general B2B market"}
COMPANY: {company}
CHANNEL TYPE / ROLE: {role or "unknown"}
KEY PRODUCTS (stored): {products[:300] or "none"}

EVIDENCE (website crawl + search — use ONLY this, do not guess):
{evidence}

Task: List the product lines, platform names, or solution brands this company offers or represents in this market.
- Use semicolon-separated names (e.g. "SecureX Platform; OT Defender; ICS Shield")
- Do NOT repeat the company name ({company}) as a brand
- For distributors / dealers / resellers / VARs / channel partners: vendor brands they represent or distribute
- For system integrators / security integrators / installers / MSSPs / MSPs: vendor platforms or product brands they deploy/resell
- For converters / fabricators: material or product brands they convert/fabricate under or for
- For manufacturers that also have a channel role: their named platforms/products in this market
- Leave empty if evidence does not name specific brands/products

Return JSON:
{{
  "key_brands_represented": "semicolon-separated brand/product names or empty",
  "confidence": "high|medium|low",
  "reason": "brief note citing evidence"
}}
"""

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
        return {}


async def _llm_brand_correction(
    row: dict[str, Any],
    evidence: str,
    *,
    market_query: str = "",
) -> dict[str, str]:
    if not evidence.strip():
        return {}
    from backend.config import get_settings, sync_opencode_to_openai_env
    from openai import AsyncOpenAI

    sync_opencode_to_openai_env()
    cfg = get_settings()
    if not cfg.openai_api_key:
        return {}

    company = norm(row.get("company") or row.get("brand"))
    brand = norm(row.get("brand"))
    domain = domain_from_row(row)
    role = norm(row.get("role"))
    products = norm(row.get("key_products"))

    client = AsyncOpenAI(api_key=cfg.openai_api_key)
    model = (os.getenv("OPENAI_MODEL") or cfg.model_extract or "gpt-4o-mini").strip()

    prompt = f"""You are a data-quality editor for a B2B market intelligence database.

MARKET: {market_query or "general B2B market"}

STORED RECORD:
- company: {company}
- brand: {brand}
- domain: {domain}
- role: {role}
- key_products: {products[:300]}

EVIDENCE (web search + LinkedIn snippets + crawl — use ONLY this, do not guess):
{evidence}

Tasks:
1. Is the stored domain the company's OFFICIAL corporate website? (not LinkedIn, ZoomInfo, directories)
2. Is the stored company name the correct legal/trade name for that domain?
3. If this is a channel partner (distributor/reseller/VAR/integrator/MSSP/dealer/converter), what brands do they represent (Key Brands Represented)?

Return JSON:
{{
  "identity_status": "ok|wrong_domain|wrong_name|wrong_both|directory|unclear",
  "correct_company": "official company name or empty if stored is fine",
  "correct_brand": "trade/consumer brand if different from company, else empty",
  "correct_domain": "official domain only e.g. acme.com — no https/www, or empty if stored is fine",
  "key_brands_represented": "semicolon-separated brand names for channel partners, or empty",
  "linkedin_company_url": "full linkedin.com/company/... URL if found in evidence, else empty",
  "apply_fix": true/false,
  "confidence": "high|medium|low",
  "reason": "brief note citing evidence"
}}

Rules:
- Set apply_fix=true only when confidence is high or medium AND evidence clearly supports a correction.
- For wrong_domain: correct_domain must appear in evidence (LinkedIn company page or official site snippet).
- Do NOT invent domains. Empty correct_* fields mean keep stored value.
- If stored record is a directory listing (ZoomInfo, Crunchbase, etc.), set identity_status=directory and provide correct_domain if found.
"""

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
        return {}


def _stored_key_brands(row: dict[str, Any]) -> str:
    crawl = row.get("_crawl") if isinstance(row.get("_crawl"), dict) else {}
    return norm(crawl.get("key_brands_represented"))


def apply_key_brands_only(row: dict[str, Any], parsed: dict[str, str]) -> list[str]:
    """Write only _crawl.key_brands_represented. No other fields are touched."""
    conf = norm(parsed.get("confidence")).lower()
    if conf not in ("high", "medium"):
        return []

    kb = norm(parsed.get("key_brands_represented"))
    if not kb:
        return []

    company = norm(row.get("company") or row.get("brand"))
    if company and kb.lower() == company.lower():
        return []

    crawl = row.setdefault("_crawl", {})
    if norm(crawl.get("key_brands_represented")) == kb:
        return []

    crawl["key_brands_represented"] = kb
    row["_key_brands_updated"] = {
        "source": "brand_domain_resolve",
        "confidence": conf,
        "reason": norm(parsed.get("reason"))[:500],
    }
    return ["key_brands_represented"]


def apply_brand_correction(row: dict[str, Any], parsed: dict[str, str]) -> list[str]:
    """Apply LLM correction to row. Returns list of field names updated."""
    if not parsed.get("apply_fix"):
        return []

    conf = norm(parsed.get("confidence")).lower()
    if conf not in ("high", "medium"):
        return []

    updates: list[str] = []
    crawl = row.setdefault("_crawl", {})

    for src_key, dst_key in (
        ("correct_company", "company"),
        ("correct_brand", "brand"),
    ):
        val = norm(parsed.get(src_key))
        if val and val.lower() != norm(row.get(dst_key)).lower():
            row[f"_{dst_key}_before"] = row.get(dst_key)
            row[dst_key] = val
            updates.append(dst_key)

    raw_dom = norm(parsed.get("correct_domain"))
    if raw_dom:
        fixed = resolve_best_domain(
            DOMAIN_OVERRIDES.get(_normalize_domain(raw_dom), raw_dom),
            name=norm(row.get("company")),
        )
        fixed = crawl_host_for_domain(fixed)
        cur = domain_from_row(row)
        if fixed and _normalize_domain(fixed) != _normalize_domain(cur):
            row["_domain_before"] = row.get("domain")
            row["domain"] = fixed
            row["website"] = fixed if not fixed.startswith("http") else fixed
            if not str(row["website"]).startswith("http"):
                row["website"] = f"https://{fixed}"
            updates.append("domain")

    kb = norm(parsed.get("key_brands_represented"))
    if kb:
        crawl["key_brands_represented"] = kb
        updates.append("key_brands_represented")

    li = norm(parsed.get("linkedin_company_url"))
    if li and "linkedin.com/company" in li.lower():
        crawl["linkedin"] = crawl.get("linkedin") or li
        updates.append("linkedin")

    if updates:
        row["_identity_corrected"] = {
            "source": "brand_domain_resolve",
            "identity_status": norm(parsed.get("identity_status")),
            "reason": norm(parsed.get("reason"))[:500],
            "fields": updates,
        }

    return updates


def needs_key_brands_resolve(row: dict[str, Any], *, cache: dict[str, dict] | None = None) -> bool:
    key = norm(row.get("company") or row.get("brand")).lower()
    if cache and key in cache:
        cached = cache[key]
        if cached.get("key_brands_only") and norm(cached.get("key_brands_represented")):
            return False
        # Don't retry forever after a hard skip/timeout in this run's cache
        if cached.get("skipped") and cached.get("error"):
            return False

    kb = _stored_key_brands(row)
    if not kb:
        return True

    company = norm(row.get("company") or row.get("brand"))
    if company and kb.lower() == company.lower():
        return True

    brand = norm(row.get("brand"))
    if brand and company and brand.lower() == company.lower() and kb.lower() == company.lower():
        return True

    return False


def needs_brand_resolve(row: dict[str, Any], *, cache: dict[str, dict] | None = None) -> bool:
    if row.get("_identity_corrected"):
        return False
    key = norm(row.get("company") or row.get("brand")).lower()
    if cache and key in cache and cache[key].get("skipped") is not True:
        cached = cache[key]
        if cached.get("apply_fix") is False and norm(cached.get("identity_status")) == "ok":
            return False

    domain = domain_from_row(row)
    if is_bad_domain(domain):
        return True
    if not domain:
        return True

    company = norm(row.get("company") or row.get("brand"))
    brand = norm(row.get("brand"))
    if brand and company and brand.lower() != company.lower():
        # Brand differs from company — worth verifying for distributors.
        return True

    return False


async def resolve_brand_domain_for_row(
    row: dict[str, Any],
    router: Any,
    *,
    market_query: str = "",
    force: bool = False,
    key_brands_only: bool = False,
) -> list[str]:
    if key_brands_only:
        if not force and not needs_key_brands_resolve(row):
            return []
    elif not force and not needs_brand_resolve(row):
        return []

    evidence = await gather_brand_evidence(
        row,
        router,
        key_brands_only=key_brands_only,
        market_query=market_query,
    )
    if not evidence.strip():
        return []

    if key_brands_only:
        parsed = await _llm_key_brands_only(row, evidence, market_query=market_query)
        return apply_key_brands_only(row, parsed) if parsed else []

    parsed = await _llm_brand_correction(row, evidence, market_query=market_query)
    if not parsed:
        return []

    # Heuristic LinkedIn URL from snippets if LLM missed it
    if not norm(parsed.get("linkedin_company_url")):
        li = _linkedin_url_from_snippets(evidence)
        if li:
            parsed["linkedin_company_url"] = li

    return apply_brand_correction(row, parsed)


def _load_cache(cache_path: Path) -> dict[str, dict]:
    if not cache_path.is_file():
        return {}
    try:
        data = json.loads(cache_path.read_text(encoding="utf-8"))
        if isinstance(data, dict):
            return data
    except Exception:
        pass
    return {}


def _save_cache(cache_path: Path, cache: dict[str, dict]) -> None:
    cache_path.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")


async def resolve_brand_domains_for_market(
    rows: list[dict[str, Any]],
    *,
    market_query: str = "",
    market_dir: Path | None = None,
    concurrent: int = 3,
    force: bool = False,
    founded_hq_only: bool = True,
    key_brands_only: bool = False,
) -> dict[str, Any]:
    import asyncio

    from vendor_intel.clients.search_router import FreeSearchRouter
    from vendor_intel.config import Settings

    cache_path = (market_dir or Path(".")) / "brand_resolve_cache.json"
    cache = _load_cache(cache_path)

    candidates = [r for r in rows if r.get("is_relevant", True)]
    if founded_hq_only:
        from vendor_intel.enrichment.gap_fill.gaps import has_founded, has_headquarters

        candidates = [r for r in candidates if has_founded(r) and has_headquarters(r)]

    if force:
        todo = candidates
    elif key_brands_only:
        todo = [r for r in candidates if needs_key_brands_resolve(r, cache=cache)]
    else:
        todo = [r for r in candidates if needs_brand_resolve(r, cache=cache)]

    if not todo:
        return {
            "queued": 0,
            "corrected": 0,
            "skipped_ok": len(candidates),
            "errors": 0,
            "key_brands_only": key_brands_only,
        }

    router = FreeSearchRouter(Settings.load())
    # Serial for key-brands: concurrent DDGS searches hang under rate-limits
    workers = 1 if key_brands_only else max(1, concurrent)
    sem = asyncio.Semaphore(workers)
    corrected = 0
    errors = 0
    timeouts = 0
    mode = "key_brands" if key_brands_only else "brand_domain"
    try:
        row_timeout = float(os.getenv("KEY_BRANDS_ROW_TIMEOUT_SEC", "120"))
    except ValueError:
        row_timeout = 120.0
    row_timeout = max(45.0, row_timeout)

    print(
        f"  [{mode}] queue={len(todo)} workers={workers} row_timeout={row_timeout:.0f}s "
        f"cache={cache_path.name} ({len(cache)} cached)",
        flush=True,
    )

    async def _one(row: dict[str, Any], idx: int) -> None:
        nonlocal corrected, errors, timeouts
        async with sem:
            name = str(row.get("company") or row.get("brand") or "?")[:50]
            key = norm(row.get("company") or row.get("brand")).lower()
            print(f"  [{mode}] starting [{idx}/{len(todo)}] {name}", flush=True)
            try:
                applied = await asyncio.wait_for(
                    resolve_brand_domain_for_row(
                        row,
                        router,
                        market_query=market_query,
                        force=force,
                        key_brands_only=key_brands_only,
                    ),
                    timeout=row_timeout,
                )
                if key_brands_only:
                    meta = row.get("_key_brands_updated") or {}
                    cache[key] = {
                        "key_brands_only": True,
                        "key_brands_represented": _stored_key_brands(row),
                        "fields": applied,
                        "reason": meta.get("reason", ""),
                    }
                else:
                    meta = row.get("_identity_corrected") or {}
                    cache[key] = {
                        "identity_status": meta.get(
                            "identity_status", "ok" if not applied else "fixed"
                        ),
                        "apply_fix": bool(applied),
                        "fields": applied,
                        "reason": meta.get("reason", ""),
                    }
                if applied:
                    corrected += 1
                    if key_brands_only:
                        kb = _stored_key_brands(row)
                        print(
                            f"  [{mode}] [{idx}/{len(todo)}] {name} -> key brands: {kb[:80]!r}",
                            flush=True,
                        )
                    else:
                        dom = domain_from_row(row)
                        print(
                            f"  [{mode}] [{idx}/{len(todo)}] {name} -> "
                            f"{norm(row.get('company'))!r} @ {dom!r} (+{applied})",
                            flush=True,
                        )
                else:
                    print(f"  [{mode}] [{idx}/{len(todo)}] {name} (no updates)", flush=True)
            except asyncio.TimeoutError:
                timeouts += 1
                cache[key] = {
                    "key_brands_only": bool(key_brands_only),
                    "error": f"timeout after {row_timeout:.0f}s",
                    "skipped": True,
                }
                print(
                    f"  [{mode}] [{idx}/{len(todo)}] {name} TIMEOUT after {row_timeout:.0f}s — skipped",
                    flush=True,
                )
            except Exception as exc:
                errors += 1
                cache[key] = {"error": str(exc), "skipped": True}
                print(f"  [{mode}] [{idx}/{len(todo)}] {name} ERROR: {exc}", flush=True)
            # Checkpoint after every company so stop/resume works
            try:
                _save_cache(cache_path, cache)
            except Exception:
                pass

    for i, r in enumerate(todo, 1):
        await _one(r, i)
    _save_cache(cache_path, cache)

    return {
        "queued": len(todo),
        "corrected": corrected,
        "errors": errors,
        "timeouts": timeouts,
        "cache_path": str(cache_path),
        "key_brands_only": key_brands_only,
    }
