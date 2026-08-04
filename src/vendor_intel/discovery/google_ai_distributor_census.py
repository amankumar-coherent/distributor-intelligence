"""Google AI scraper distributor census.

Asks Google AI (via :15551 + Chrome extension) for distributor / channel-partner
company lists for the current market. The extension expands Show more / View all
before scrape; we then extract company names into DiscoveryHit rows.

Prompts adapt to the market name, e.g.:
  "List 100 avocado oil distributors and wholesalers worldwide with company names"
"""
from __future__ import annotations

import json
import os
import re
from typing import Any

from vendor_intel.clients.google_ai_scraper import (
    google_ai_ask,
    google_ai_discovery_enabled,
    google_ai_enabled,
    google_ai_ping,
    heuristic_clean_scrape_markdown,
)
from vendor_intel.discovery.country_sweep import PRIORITY_COUNTRIES
from vendor_intel.models import DiscoveryHit

_NAME_LINE_RE = re.compile(
    r"^\s*(?:[-*•]|\d+[.)])\s+(?:\*\*)?([A-Z][^:*\n]{1,80}?)(?:\*\*)?(?:\s*[-–—:|]|\s*$)",
    re.M,
)
_BOLD_RE = re.compile(r"\*\*([A-Z][^*]{1,70})\*\*")
_JUNK_NAME = re.compile(
    r"^(?:ai overview|google|wikipedia|here are|top \d+|list of|official|"
    r"visit|read more|show more|web results|distributor|wholesaler|importer|"
    r"united states|india|china|germany|france|global|worldwide)\b",
    re.I,
)


def distributor_census_enabled() -> bool:
    if not google_ai_enabled() or not google_ai_discovery_enabled():
        return False
    raw = (os.getenv("GOOGLE_AI_DISTRIBUTOR_CENSUS") or "true").strip().lower()
    return raw in ("1", "true", "yes", "on")


def distributor_census_target() -> int:
    try:
        return max(20, min(500, int(os.getenv("GOOGLE_AI_DISTRIBUTOR_CENSUS_TARGET", "100") or "100")))
    except ValueError:
        return 100


def distributor_census_max_queries() -> int:
    try:
        return max(3, min(80, int(os.getenv("GOOGLE_AI_DISTRIBUTOR_CENSUS_MAX_QUERIES", "35") or "35")))
    except ValueError:
        return 35


def _market_label(market: str) -> str:
    m = (market or "").strip()
    if m.lower().endswith(" market"):
        m = m[: -len(" market")].strip()
    return m or "market"


def build_distributor_census_queries(
    market: str,
    *,
    target: int | None = None,
    max_queries: int | None = None,
) -> list[str]:
    """Market-aware Google AI prompts that ask for distributor company names."""
    m = _market_label(market)
    t = target if target is not None else distributor_census_target()
    limit = max_queries if max_queries is not None else distributor_census_max_queries()

    queries: list[str] = [
        (
            f"List {t} {_market_label(m)} distributors wholesalers and importers "
            f"worldwide. Reply with a numbered list of company names only "
            f"(distributors and channel partners, no pure manufacturers)."
        ),
        (
            f"Top {min(t, 50)} {m} distributor companies and bulk wholesalers "
            f"globally — company names only, Show more full list."
        ),
        (
            f"{m} authorized distributors and channel partners worldwide "
            f"company names list (at least {min(t, 40)} names)."
        ),
        (
            f"B2B {m} foodservice and ingredient distributors company names "
            f"list worldwide."
        ),
    ]

    # Regional batches — forces Google AI Overview lists per country.
    for country in PRIORITY_COUNTRIES:
        queries.append(
            f"List {m} distributors wholesalers importers in {country}. "
            f"Numbered company names only (channel partners, not manufacturers)."
        )
        if len(queries) >= limit:
            break

    # Dedupe while preserving order
    out: list[str] = []
    seen: set[str] = set()
    for q in queries:
        key = re.sub(r"\s+", " ", q.lower()).strip()
        if key in seen:
            continue
        seen.add(key)
        out.append(q)
        if len(out) >= limit:
            break
    return out


def extract_company_names_heuristic(markdown: str) -> list[str]:
    """Pull company-like names from AI Overview markdown (bullets / bold)."""
    text = heuristic_clean_scrape_markdown(markdown or "", max_chars=12000)
    found: list[str] = []
    seen: set[str] = set()

    def _add(raw: str) -> None:
        name = re.sub(r"\s+", " ", (raw or "").strip(" .,-;:|"))
        name = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", name)
        name = name.strip(" *")
        if len(name) < 3 or len(name) > 80:
            return
        if _JUNK_NAME.match(name):
            return
        if name.lower().startswith("http"):
            return
        # Drop sentences
        if name.count(" ") > 8 or name.endswith("."):
            return
        key = re.sub(r"[^a-z0-9]", "", name.lower())
        if len(key) < 3 or key in seen:
            return
        seen.add(key)
        found.append(name)

    for m in _NAME_LINE_RE.finditer(text):
        _add(m.group(1))
    for m in _BOLD_RE.finditer(text):
        _add(m.group(1))

    # Fallback: lines that look like "Company Name - description"
    for line in text.splitlines():
        s = line.strip()
        if not s or len(s) > 100:
            continue
        if re.match(r"^[-*•\d]", s):
            s = re.sub(r"^[-*•\d.)\s]+", "", s)
        if " - " in s:
            _add(s.split(" - ", 1)[0])
        elif ":" in s and len(s.split(":")[0]) < 60:
            left = s.split(":", 1)[0]
            if left[0:1].isupper():
                _add(left)

    return found


async def extract_company_names_llm(market: str, markdown: str) -> list[str]:
    """Optional OpenAI pass: return distributor company names as JSON list."""
    api_key = (os.getenv("OPENAI_API_KEY") or "").strip()
    if not api_key or len(markdown or "") < 40:
        return []
    base_url = (os.getenv("OPENAI_BASE_URL") or "https://api.openai.com/v1").strip().rstrip("/")
    model = (
        os.getenv("GOOGLE_AI_SCRAPER_LLM_MODEL")
        or os.getenv("OPENAI_MODEL")
        or "gpt-4o-mini"
    )
    trimmed = heuristic_clean_scrape_markdown(markdown, max_chars=7000)
    prompt = (
        f"Market: {_market_label(market)}\n"
        "From the Google AI Overview text below, extract ONLY company names that are "
        "distributors, wholesalers, importers, resellers, or channel partners. "
        "EXCLUDE pure manufacturers, brands that only produce, directories, and "
        "non-companies. Return JSON: {\"companies\": [\"Name1\", \"Name2\", ...]}\n\n"
        f"TEXT:\n{trimmed}"
    )
    try:
        import httpx

        async with httpx.AsyncClient(timeout=60.0) as client:
            r = await client.post(
                f"{base_url}/chat/completions",
                headers={"Authorization": f"Bearer {api_key}"},
                json={
                    "model": model,
                    "temperature": 0,
                    "response_format": {"type": "json_object"},
                    "messages": [
                        {
                            "role": "system",
                            "content": "Extract distributor company names as JSON only.",
                        },
                        {"role": "user", "content": prompt},
                    ],
                },
            )
            if r.status_code != 200:
                return []
            data = r.json()
            content = (
                ((data.get("choices") or [{}])[0].get("message") or {}).get("content") or ""
            )
            parsed = json.loads(content)
            names = parsed.get("companies") if isinstance(parsed, dict) else parsed
            if not isinstance(names, list):
                return []
            out: list[str] = []
            seen: set[str] = set()
            for n in names:
                s = str(n or "").strip()
                if len(s) < 3:
                    continue
                key = re.sub(r"[^a-z0-9]", "", s.lower())
                if key in seen:
                    continue
                seen.add(key)
                out.append(s)
            return out
    except Exception:
        return []


def names_to_discovery_hits(
    names: list[str],
    *,
    market: str,
    prompt_id: str,
) -> list[DiscoveryHit]:
    hits: list[DiscoveryHit] = []
    m = _market_label(market)
    for i, name in enumerate(names):
        # Unique synthetic domain so later domain-dedupe does not collapse the census.
        # Brand/domain resolve replaces these with real websites in post-market.
        slug = re.sub(r"[^a-z0-9]", "", name.lower())[:48] or f"co{i + 1}"
        domain = f"{slug}.gai.census"
        hits.append(
            DiscoveryHit(
                name_raw=name,
                source_url=f"https://{domain}/",
                source_domain=domain,
                prompt_id=f"{prompt_id}_{i + 1}",
                backend="google_ai_census",
                snippet=(
                    f"Google AI distributor census for {m}: {name} "
                    f"(channel partner / distributor; resolve official website later)."
                ),
                funnel_level="gai_census",
                search_theme="distributor_census",
            )
        )
    return hits


async def run_google_ai_distributor_census(
    market: str,
    *,
    target: int | None = None,
    max_queries: int | None = None,
    existing_names: set[str] | None = None,
) -> list[DiscoveryHit]:
    """Run multi-query Google AI census; stop once target unique names reached."""
    if not distributor_census_enabled():
        return []
    if not await google_ai_ping():
        print("  [gai-census] Google AI scraper not reachable — skip", flush=True)
        return []

    tgt = target if target is not None else distributor_census_target()
    queries = build_distributor_census_queries(
        market, target=tgt, max_queries=max_queries
    )
    print(
        f"  [gai-census] Asking Google AI for ~{tgt} {_market_label(market)} "
        f"distributors ({len(queries)} queries; Show more expanded in extension)",
        flush=True,
    )

    have = {
        re.sub(r"[^a-z0-9]", "", n.lower())
        for n in (existing_names or set())
        if n
    }
    all_names: list[str] = []
    hits: list[DiscoveryHit] = []

    for qi, query in enumerate(queries, start=1):
        if len(all_names) >= tgt:
            break
        data = await google_ai_ask(query, llm_clean=True)
        md = str(data.get("markdown") or "")
        if data.get("error") and not md:
            print(
                f"  [gai-census] query {qi}/{len(queries)} error: {data.get('error')}",
                flush=True,
            )
            continue

        names = await extract_company_names_llm(market, md)
        if len(names) < 3:
            heur = extract_company_names_heuristic(md)
            # merge
            seen = {re.sub(r"[^a-z0-9]", "", n.lower()) for n in names}
            for h in heur:
                k = re.sub(r"[^a-z0-9]", "", h.lower())
                if k not in seen:
                    names.append(h)
                    seen.add(k)

        new: list[str] = []
        for n in names:
            k = re.sub(r"[^a-z0-9]", "", n.lower())
            if not k or k in have:
                continue
            have.add(k)
            new.append(n)
            all_names.append(n)

        batch_hits = names_to_discovery_hits(
            new, market=market, prompt_id=f"GAI_CENSUS{qi}"
        )
        hits.extend(batch_hits)
        print(
            f"  [gai-census] [{qi}/{len(queries)}] +{len(new)} names "
            f"(total {len(all_names)}/{tgt}) — {query[:70]}…",
            flush=True,
        )

    print(
        f"  [gai-census] Done: {len(all_names)} distributor names from Google AI",
        flush=True,
    )
    return hits


def census_from_scope(scope: dict[str, Any] | None, query: str = "") -> str:
    if isinstance(scope, dict):
        m = str(scope.get("market") or "").strip()
        if m:
            return m
    return (query or "").strip() or "market"
