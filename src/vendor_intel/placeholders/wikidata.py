"""Wikidata SPARQL + Search API — free, no API key required.

Wikidata is open data. Use a descriptive User-Agent (Wikimedia etiquette).
Docs: https://www.wikidata.org/wiki/Wikidata:Data_access
SPARQL: https://query.wikidata.org/sparql
Search: https://www.wikidata.org/w/api.php?action=wbsearchentities
"""
from __future__ import annotations

import asyncio
import os
import re
from typing import Any
from urllib.parse import urlparse

from vendor_intel.placeholders._http import get_json

WIKIDATA_ENABLED = True
WIKIDATA_SPARQL_URL = "https://query.wikidata.org/sparql"
WIKIDATA_API_URL = "https://www.wikidata.org/w/api.php"
WIKIDATA_USER_AGENT = (
    "VendorIntelPipeline/0.1 (research; vendor-intel agent; "
    "+https://github.com/local; mailto:research@local)"
)

# Rate-limit SPARQL (Wikimedia asks clients not to hammer the endpoint).
_SPARQL_MIN_INTERVAL_SEC = float(os.getenv("WIKIDATA_SPARQL_DELAY_SEC", "1.0"))
_last_sparql_at = 0.0
_sparql_lock = asyncio.Lock()

_YEAR_RE = re.compile(r"(18|19|20)\d{2}")

_SPARQL_PARENT = """
SELECT ?parentLabel WHERE {{
  ?item rdfs:label "{name}"@en .
  ?item wdt:P31/wdt:P279* wd:Q4830453 .
  ?item wdt:P749 ?parent .
  ?parent rdfs:label ?parentLabel .
  FILTER(LANG(?parentLabel) = "en")
}}
LIMIT 3
"""

_SPARQL_BY_QID = """
SELECT ?item ?itemLabel ?founded ?hqLabel ?website ?employees ?industryLabel ?parentLabel WHERE {{
  BIND(wd:{qid} AS ?item)
  OPTIONAL {{ ?item wdt:P571 ?founded . }}
  OPTIONAL {{ ?item wdt:P159 ?hq . }}
  OPTIONAL {{ ?item wdt:P856 ?website . }}
  OPTIONAL {{ ?item wdt:P1128 ?employees . }}
  OPTIONAL {{ ?item wdt:P452 ?industry . }}
  OPTIONAL {{ ?item wdt:P749 ?parent . }}
  SERVICE wikibase:label {{ bd:serviceParam wikibase:language "en". }}
}}
LIMIT 5
"""

_SPARQL_BY_LABEL = """
SELECT ?item ?itemLabel ?founded ?hqLabel ?website ?employees ?industryLabel ?parentLabel WHERE {{
  ?item rdfs:label "{name}"@en .
  ?item wdt:P31/wdt:P279* wd:Q4830453 .
  OPTIONAL {{ ?item wdt:P571 ?founded . }}
  OPTIONAL {{ ?item wdt:P159 ?hq . }}
  OPTIONAL {{ ?item wdt:P856 ?website . }}
  OPTIONAL {{ ?item wdt:P1128 ?employees . }}
  OPTIONAL {{ ?item wdt:P452 ?industry . }}
  OPTIONAL {{ ?item wdt:P749 ?parent . }}
  SERVICE wikibase:label {{ bd:serviceParam wikibase:language "en". }}
}}
LIMIT 5
"""

# Require host-boundary match so microsoft.com does not match microsoft.com/ja-jp
# subsidiaries preferentially when ranking — still returned, then ranked in Python.
_SPARQL_BY_DOMAIN = """
SELECT ?item ?itemLabel ?founded ?hqLabel ?website ?employees ?industryLabel ?parentLabel WHERE {{
  ?item wdt:P856 ?website .
  FILTER(REGEX(LCASE(STR(?website)), "^https?://(www\\\\.)?{domain_escaped}([/?#]|$)", "i"))
  ?item wdt:P31/wdt:P279* wd:Q4830453 .
  OPTIONAL {{ ?item wdt:P571 ?founded . }}
  OPTIONAL {{ ?item wdt:P159 ?hq . }}
  OPTIONAL {{ ?item wdt:P1128 ?employees . }}
  OPTIONAL {{ ?item wdt:P452 ?industry . }}
  OPTIONAL {{ ?item wdt:P749 ?parent . }}
  SERVICE wikibase:label {{ bd:serviceParam wikibase:language "en". }}
}}
LIMIT 10
"""


def is_enabled() -> bool:
    return bool(WIKIDATA_ENABLED)


def _headers() -> dict[str, str]:
    return {
        "Accept": "application/sparql-results+json",
        "User-Agent": WIKIDATA_USER_AGENT,
    }


def _safe_label(name: str) -> str:
    return (name or "").replace("\\", " ").replace('"', "").strip()[:120]


def _normalize_domain(domain: str) -> str:
    d = (domain or "").strip().lower()
    if not d:
        return ""
    if "://" in d:
        d = urlparse(d).netloc or d
    d = d.removeprefix("www.").split("/")[0].strip()
    return d


def _year_from_value(raw: Any) -> str:
    s = str(raw or "").strip()
    m = _YEAR_RE.search(s)
    return m.group(0) if m else ""


def _qid_from_uri(uri: str) -> str:
    # https://www.wikidata.org/entity/Q312 → Q312
    s = (uri or "").strip()
    if not s:
        return ""
    if s.startswith("Q") and s[1:].isdigit():
        return s
    m = re.search(r"/(Q\d+)$", s)
    return m.group(1) if m else ""


def _binding_value(binding: dict[str, Any], key: str) -> str:
    node = binding.get(key) or {}
    return str(node.get("value") or "").strip()


def _website_host(url: str) -> str:
    return _normalize_domain(url)


def _domain_match_score(website: str, preferred_domain: str) -> int:
    """Higher = better host match. Exact host beats subdomains / unrelated."""
    if not preferred_domain:
        return 0
    host = _website_host(website)
    if not host:
        return 0
    if host == preferred_domain:
        return 100
    if host.endswith("." + preferred_domain):
        return 40  # store.apple.com
    if preferred_domain.endswith("." + host):
        return 20
    return 0


def _parse_company_bindings(
    bindings: list[dict[str, Any]],
    *,
    preferred_domain: str = "",
    preferred_name: str = "",
) -> dict[str, str] | None:
    if not bindings:
        return None
    # Prefer: exact website host match → top-level company (no parent) → richer fields.
    best: dict[str, str] | None = None
    best_score = -1
    name_l = (preferred_name or "").lower().strip()
    for b in bindings:
        founded = _year_from_value(_binding_value(b, "founded"))
        hq = _binding_value(b, "hqLabel")
        website = _binding_value(b, "website")
        employees = _binding_value(b, "employees")
        industry = _binding_value(b, "industryLabel")
        parent = _binding_value(b, "parentLabel")
        item_uri = _binding_value(b, "item")
        label = _binding_value(b, "itemLabel")
        qid = _qid_from_uri(item_uri)
        row = {
            "wikidata_qid": qid,
            "wikidata_label": label,
            "founded_year": founded,
            "headquarters": hq,
            "website": website,
            "employee_count": employees,
            "industry": industry,
            "parent_org": parent,
            "source_url": f"https://www.wikidata.org/wiki/{qid}" if qid else "https://www.wikidata.org",
        }
        score = 0
        score += _domain_match_score(website, preferred_domain)
        score += 10 * sum(
            1
            for k in ("founded_year", "headquarters", "website", "employee_count", "industry")
            if row.get(k)
        )
        if not parent:
            score += 25  # prefer parent/top-level entity
        if name_l and label:
            label_l = label.lower()
            if label_l == name_l:
                score += 30
            elif name_l in label_l or label_l in name_l:
                score += 10
        # Prefer corporate HQ labels over regional sites.
        if website and preferred_domain:
            host = _website_host(website)
            path = urlparse(website).path or ""
            if host == preferred_domain and (path in ("", "/")):
                score += 20
            if any(x in website.lower() for x in ("/ja-", "/zh-", "/fr-", "/de-", "/cn/", ".cn/")):
                score -= 15
        if score > best_score:
            best_score = score
            best = row
    return best


async def _throttle_sparql() -> None:
    global _last_sparql_at
    async with _sparql_lock:
        import time

        now = time.monotonic()
        wait = _SPARQL_MIN_INTERVAL_SEC - (now - _last_sparql_at)
        if wait > 0:
            await asyncio.sleep(wait)
        _last_sparql_at = time.monotonic()


async def _sparql(query: str) -> list[dict[str, Any]]:
    if not is_enabled():
        return []
    await _throttle_sparql()
    try:
        data = await get_json(
            WIKIDATA_SPARQL_URL,
            params={"query": query, "format": "json"},
            headers=_headers(),
            timeout=45.0,
        )
        return list((data.get("results") or {}).get("bindings") or [])
    except Exception:
        return []


async def search_entity_qid(company_name: str, *, limit: int = 5) -> list[dict[str, str]]:
    """Free MediaWiki search → candidate QIDs (no API key)."""
    if not is_enabled():
        return []
    q = _safe_label(company_name)
    if not q:
        return []
    try:
        data = await get_json(
            WIKIDATA_API_URL,
            params={
                "action": "wbsearchentities",
                "search": q,
                "language": "en",
                "type": "item",
                "limit": str(max(1, min(limit, 10))),
                "format": "json",
            },
            headers={"User-Agent": WIKIDATA_USER_AGENT, "Accept": "application/json"},
            timeout=20.0,
        )
        out: list[dict[str, str]] = []
        for hit in data.get("search") or []:
            qid = str(hit.get("id") or "").strip()
            if not qid.startswith("Q"):
                continue
            out.append(
                {
                    "qid": qid,
                    "label": str(hit.get("label") or "").strip(),
                    "description": str(hit.get("description") or "").strip(),
                }
            )
        return out
    except Exception:
        return []


def _looks_like_org(desc: str, label: str, company: str) -> bool:
    blob = f"{label} {desc}".lower()
    company_l = (company or "").lower()
    org_hints = (
        "company",
        "corporation",
        "manufacturer",
        "enterprise",
        "business",
        "brand",
        "electronics",
        "technology",
        "firm",
        "group",
        "inc",
        "ltd",
        "gmbh",
        "ag",
    )
    if any(h in blob for h in org_hints):
        return True
    # Label contains the company name (or vice versa)
    if company_l and (company_l in label.lower() or label.lower() in company_l):
        return True
    return False


async def lookup_parent_org(company_name: str) -> tuple[str | None, str | None]:
    """Backward-compatible parent-org lookup used by Phase 2/3 enrichment."""
    if not is_enabled():
        return None, None
    safe = _safe_label(company_name)
    if not safe:
        return None, None
    bindings = await _sparql(_SPARQL_PARENT.format(name=safe))
    if bindings:
        label = _binding_value(bindings[0], "parentLabel")
        if label:
            return label, "https://www.wikidata.org"
    # Fallback: full company lookup
    firm = await lookup_company(company_name)
    parent = (firm or {}).get("parent_org") or ""
    if parent:
        return parent, (firm or {}).get("source_url") or "https://www.wikidata.org"
    return None, None


async def lookup_company(
    company_name: str,
    *,
    domain: str = "",
) -> dict[str, str] | None:
    """Look up firmographics from free Wikidata SPARQL (no API key).

    Returns keys usable by gap-fill merge_extracted:
      founded_year, headquarters, website, employee_count, industry,
      parent_org, wikidata_qid, wikidata_label, source_url
    """
    if not is_enabled():
        return None
    name = _safe_label(company_name)
    dom = _normalize_domain(domain)
    if not name and not dom:
        return None

    # 1) Domain match is most precise when available.
    if dom:
        domain_escaped = re.escape(dom)
        parsed = _parse_company_bindings(
            await _sparql(_SPARQL_BY_DOMAIN.format(domain_escaped=domain_escaped)),
            preferred_domain=dom,
            preferred_name=name,
        )
        if parsed and (parsed.get("founded_year") or parsed.get("headquarters") or parsed.get("website")):
            # Accept only if host is the same registrable domain.
            if not dom or _domain_match_score(parsed.get("website") or "", dom) >= 40 or not parsed.get("website"):
                return parsed

    # 2) Exact English label match for business enterprises.
    if name:
        parsed = _parse_company_bindings(
            await _sparql(_SPARQL_BY_LABEL.format(name=name)),
            preferred_domain=dom,
            preferred_name=name,
        )
        if parsed and (parsed.get("founded_year") or parsed.get("headquarters") or parsed.get("website")):
            if not dom or _domain_match_score(parsed.get("website") or "", dom) >= 40 or not parsed.get("website"):
                return parsed

    # 3) Fuzzy search → QID → SPARQL by entity id.
    if name:
        candidates = await search_entity_qid(name, limit=5)
        best: dict[str, str] | None = None
        best_score = -1
        for cand in candidates:
            if not _looks_like_org(cand.get("description", ""), cand.get("label", ""), name):
                continue
            qid = cand["qid"]
            parsed = _parse_company_bindings(
                await _sparql(_SPARQL_BY_QID.format(qid=qid)),
                preferred_domain=dom,
                preferred_name=name,
            )
            if not parsed:
                continue
            score = 0
            score += _domain_match_score(parsed.get("website") or "", dom)
            score += 10 * sum(
                1 for k in ("founded_year", "headquarters", "employee_count", "industry") if parsed.get(k)
            )
            if not parsed.get("parent_org"):
                score += 25
            if score > best_score and (
                parsed.get("founded_year")
                or parsed.get("headquarters")
                or parsed.get("website")
                or parsed.get("parent_org")
            ):
                best_score = score
                best = parsed
        if best:
            return best

    return None
