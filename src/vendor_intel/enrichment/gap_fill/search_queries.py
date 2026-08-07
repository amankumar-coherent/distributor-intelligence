"""Focused Google / Google-AI / DDGS / SearXNG query builders for Step-5 fill.

Use short, field-specific search-bar phrasing (one fact per query), e.g.:
  {company} founded year
  {company} headquarters
  {company} number of employees
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


# Gap key → Google search phrasing (append after company name)
_FIELD_PHRASE: dict[str, str] = {
    "founded_year": "founded year",
    "headquarters": "headquarters",
    "employee_count": "number of employees",
    "revenue": "revenue turnover",
    "offices": "offices locations",
    "ownership": "ownership parent company",
    "core_categories": "products categories",
    "specialty_focus": "specialty focus",
    "key_brands_represented": "brands products portfolio",
    "contact_person": "CEO founder",
    "linkedin": "LinkedIn company",
    "website": "official website",
}

# Preferred ask order when several gaps are open
_FIELD_ORDER: tuple[str, ...] = (
    "founded_year",
    "headquarters",
    "employee_count",
    "ownership",
    "offices",
    "website",
    "core_categories",
    "specialty_focus",
    "key_brands_represented",
    "revenue",
    "linkedin",
    "contact_person",
)


def field_query(company: str, field: str, *, domain: str = "") -> str:
    """One Google query for one missing field: ``{company} founded year``."""
    qn = _quoted(company)
    if not qn:
        return ""
    phrase = _FIELD_PHRASE.get(field) or field.replace("_", " ")
    q = f"{qn} {phrase}"
    dom = _clean_domain(domain)
    # Soft site hint only for brand/category lookups (keeps HQ/founded global)
    if dom and field in {"core_categories", "specialty_focus", "key_brands_represented", "website"}:
        q = f"{q} {dom}"
    return q


def queries_for_fields(
    company: str,
    fields: list[str] | tuple[str, ...] | set[str],
    *,
    domain: str = "",
    market: str = "",
    max_queries: int = 6,
) -> list[str]:
    """Build ``{company} <field phrase>`` queries for the given gaps."""
    wanted = {str(f).strip() for f in fields if str(f).strip()}
    if not wanted:
        return []

    ranked: list[str] = []
    for key in _FIELD_ORDER:
        if key not in wanted:
            continue
        q = field_query(company, key, domain=domain)
        if q:
            ranked.append(q)

    # Any unknown gap keys still get a plain phrase
    for key in sorted(wanted):
        if key in _FIELD_PHRASE or key in _FIELD_ORDER:
            continue
        q = field_query(company, key, domain=domain)
        if q:
            ranked.append(q)

    # Optional market + brands query when categories/brands missing
    mkt = re.sub(r"\s+", " ", (market or "").strip())
    mkt = re.sub(r"(?i)^global\s+", "", mkt).strip()
    if mkt and wanted & {"core_categories", "key_brands_represented", "specialty_focus"}:
        qn = _quoted(company)
        if qn:
            ranked.append(f"{qn} {mkt} brands distributors")

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


def firmographic_queries(company: str, *, domain: str = "") -> list[str]:
    """HQ / founded / employees — one query per fact (legacy helper)."""
    return queries_for_fields(
        company,
        ("founded_year", "headquarters", "employee_count"),
        domain=domain,
        max_queries=3,
    )


def leadership_queries(company: str, *, domain: str = "") -> list[str]:
    """CEO / founder — only when contact fill is explicitly requested."""
    qn = _quoted(company)
    if not qn:
        return []
    out = [
        f"{qn} CEO founder",
        f"{qn} Chief Executive Officer",
    ]
    dom = _clean_domain(domain)
    if dom:
        out.append(f"{qn} CEO site:{dom}")
    return out


def category_queries(company: str, *, market: str = "", domain: str = "") -> list[str]:
    return queries_for_fields(
        company,
        ("core_categories", "key_brands_represented", "specialty_focus"),
        domain=domain,
        market=market,
        max_queries=3,
    )


def build_gap_fill_queries(
    company: str,
    *,
    domain: str = "",
    market: str = "",
    need_firmographics: bool = False,
    need_categories: bool = False,
    need_contact: bool = False,
    needed_fields: list[str] | tuple[str, ...] | set[str] | None = None,
    max_queries: int = 6,
) -> list[str]:
    """Pick field-specific Google queries for missing landscape gaps.

    Prefer ``needed_fields`` (exact gap keys). Legacy booleans still work.
    """
    fields: list[str] = []
    if needed_fields:
        fields.extend(str(f) for f in needed_fields)
    else:
        if need_firmographics:
            fields.extend(
                ("founded_year", "headquarters", "offices", "employee_count", "revenue")
            )
        if need_categories:
            fields.extend(
                ("core_categories", "specialty_focus", "key_brands_represented")
            )
        if need_contact:
            fields.append("contact_person")

    # De-dupe preserving order
    seen: set[str] = set()
    uniq: list[str] = []
    for f in fields:
        if f in seen:
            continue
        seen.add(f)
        uniq.append(f)

    return queries_for_fields(
        company,
        uniq,
        domain=domain,
        market=market,
        max_queries=max_queries,
    )


def backfill_queries(company: str, *, domain: str = "", max_queries: int = 3) -> list[str]:
    """DDGS/SearXNG firmographic queries (one fact each)."""
    return queries_for_fields(
        company,
        ("founded_year", "headquarters", "employee_count"),
        domain=domain,
        max_queries=max_queries,
    )


def backfill_query(company: str, *, domain: str = "") -> str:
    """Single DDGS/SearXNG query (first firmographic ask). Prefer ``backfill_queries``."""
    qs = backfill_queries(company, domain=domain, max_queries=1)
    return qs[0] if qs else ""
