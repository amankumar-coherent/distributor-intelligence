"""Focused Google/DDGS/SearXNG query builders for Step-5 fill (no LLM).

Prefer short, high-precision search-bar phrasing over keyword stuffing.
"""
from __future__ import annotations

import re


def _clean_company(name: str) -> str:
    return re.sub(r"\s+", " ", (name or "").strip())


def _clean_domain(domain: str) -> str:
    d = (domain or "").strip().lower()
    d = d.removeprefix("https://").removeprefix("http://").removeprefix("www.")
    return d.split("/")[0].strip()


def _quoted(name: str) -> str:
    n = _clean_company(name)
    if not n:
        return ""
    if " " in n or any(c in n for c in ".-&"):
        return f'"{n}"'
    return n


def firmographic_queries(company: str, *, domain: str = "") -> list[str]:
    """HQ / founded / employees / revenue — 1–2 tight queries."""
    qn = _quoted(company)
    if not qn:
        return []
    dom = _clean_domain(domain)
    out = [
        f"{qn} headquarters OR founded OR employees OR revenue",
        f"{qn} \"founded in\" OR \"headquartered in\" OR \"number of employees\"",
    ]
    if dom:
        out.append(f"{qn} ({dom}) headquarters founded")
    return out


def leadership_queries(company: str, *, domain: str = "") -> list[str]:
    """CEO / founder — natural Google phrasing (not title keyword dumps)."""
    qn = _quoted(company)
    if not qn:
        return []
    dom = _clean_domain(domain)
    out = [
        f"{qn} CEO OR founder",
        f"{qn} \"Chief Executive Officer\" OR \"Managing Director\"",
        f"{qn} leadership OR \"management team\" OR executive",
    ]
    if dom:
        # Domain as soft hint, not glued keyword soup
        out.insert(1, f"{qn} CEO site:{dom}")
        out.append(f'"{dom}" CEO OR founder')
    return out


def category_queries(company: str, *, market: str = "", domain: str = "") -> list[str]:
    qn = _quoted(company)
    if not qn:
        return []
    market = re.sub(r"\s+", " ", (market or "").strip())
    out = [f"{qn} distributor OR wholesale OR brands OR products"]
    if market:
        # Keep market short — drop leading "Global "
        mkt = re.sub(r"(?i)^global\s+", "", market).strip()
        if mkt:
            out.append(f"{qn} {mkt} brands OR partners OR portfolio")
    dom = _clean_domain(domain)
    if dom:
        out.append(f"{qn} brands OR products {dom}")
    return out


def build_gap_fill_queries(
    company: str,
    *,
    domain: str = "",
    market: str = "",
    need_firmographics: bool = False,
    need_categories: bool = False,
    need_contact: bool = False,
    max_queries: int = 3,
) -> list[str]:
    """Pick the best 1–3 search-bar queries for missing landscape gaps."""
    ranked: list[str] = []
    # Leadership first when missing — high business value and scrapes well from ZoomInfo/etc.
    if need_contact:
        ranked.extend(leadership_queries(company, domain=domain)[:2])
    if need_firmographics:
        ranked.extend(firmographic_queries(company, domain=domain)[:2])
    if need_categories:
        ranked.extend(category_queries(company, market=market, domain=domain)[:1])

    # De-dupe while preserving order
    seen: set[str] = set()
    out: list[str] = []
    for q in ranked:
        key = q.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(q)
        if len(out) >= max_queries:
            break
    return out


def backfill_query(company: str, *, domain: str = "") -> str:
    """Single DDGS/SearXNG query for HQ/founded/employees backfill."""
    qn = _quoted(company)
    dom = _clean_domain(domain)
    if dom:
        return f"{qn} headquarters OR founded OR employees OR revenue {dom}"
    return f"{qn} headquarters OR founded OR employees OR revenue"
