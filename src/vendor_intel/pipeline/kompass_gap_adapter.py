"""Bridge accepted Kompass distributors into the existing gap-fill engine.

Why an adapter and not a new crawler: `enrichment/gap_fill/` already knows
how to fill every column the Kompass export leaves empty — it crawls the
company site (/about, /products, /contact), reads LinkedIn, Wikidata,
Wikipedia and Crunchbase, and caches results. The Kompass stage simply
never called it, because the two halves speak different vocabularies:

  * Kompass produces `VerifiedDistributor` dataclasses whose values are
    later renamed into 21 display columns ("Key Brands Represented").
  * gap-fill operates on a plain dict with snake_case keys, and writes
    what it learns into a nested `_crawl` sub-dict.

This module translates in both directions. It deliberately does NOT
re-implement any extraction: everything here is field naming and merge
policy, so improvements to the gap-fill sources flow through for free.
"""
from __future__ import annotations

from typing import Any

# gap-fill key  ->  VerifiedDistributor attribute.
# Only these have a real home on the dataclass; everything else round-trips
# through `extra_columns` under its display name (see _EXTRA_BY_GAP).
_DIRECT_BY_GAP = {
    "website": "website",
    "founded_year": "founded_year",
    "headquarters": "headquarters",
    "employee_count": "employees",
    "contact_person": "contact_person",
    "email": "email",
    "phone": "phone",
    "key_brands_represented": "brands",
    "core_categories": "products",
    "summary": "description",
}

# gap-fill key -> the export column it feeds.
_EXTRA_BY_GAP = {
    "contact_role": "Role",
    "office_email": "Office Email",
    "linkedin": "LinkedIn",
    "specialty_focus": "Specialty Focus",
    "ownership": "Ownership",
    "retail_ecommerce": "Online / Offline / Both",
    "regions_served": "Continent / Geography",
    "offices": "Operational Presence",
}

_PLACEHOLDER = {"", "not publicly disclosed", "n/a", "none", "unknown", "-"}


def _real(value: Any) -> str:
    """A value we should treat as already known (so gap-fill leaves it alone)."""
    text = str(value or "").strip()
    return "" if text.lower() in _PLACEHOLDER else text


def to_gap_row(d: Any) -> dict[str, Any]:
    """Build the dict gap-fill expects from a VerifiedDistributor.

    Pre-seeding `_crawl` with what Kompass already established is what stops
    the engine re-researching (and possibly overwriting) good data: every
    `has_*` check in gaps.py reads `_crawl` first, so a seeded field reads as
    "no gap" and its source is skipped.
    """
    crawl: dict[str, Any] = {}
    for gap_key, attr in _DIRECT_BY_GAP.items():
        value = _real(getattr(d, attr, ""))
        if value:
            crawl[gap_key] = value

    extra = d.extra_columns or {}
    for gap_key, column in _EXTRA_BY_GAP.items():
        value = _real(extra.get(column))
        if value:
            crawl[gap_key] = value

    # Core Categories seeded with the bare market name carries no information
    # — treat it as a gap so the products page actually gets read.
    market = str(getattr(d, "market", "") or "").strip().lower()
    if crawl.get("core_categories", "").strip().lower() == market:
        crawl.pop("core_categories", None)

    return {
        "company": d.company_name,
        "brand": d.company_name,
        "domain": _real(getattr(d, "website", "")),
        "website": _real(getattr(d, "website", "")),
        "country": _real(getattr(d, "country", "")),
        "market": getattr(d, "market", ""),
        "_crawl": crawl,
    }


def apply_gap_row(d: Any, row: dict[str, Any]) -> list[str]:
    """Copy anything gap-fill learned back onto the dataclass.

    Only fills blanks. gap-fill's own merge already refuses to overwrite a
    seeded field, but this second guard means a caller can pass a partially
    seeded row without risking a good Kompass value being replaced by a
    weaker web guess. Returns the export column names actually changed.
    """
    crawl = row.get("_crawl")
    if not isinstance(crawl, dict):
        return []

    changed: list[str] = []

    for gap_key, attr in _DIRECT_BY_GAP.items():
        value = _real(crawl.get(gap_key))
        if value and not _real(getattr(d, attr, "")):
            setattr(d, attr, value)
            changed.append(attr)

    extra = dict(d.extra_columns or {})
    for gap_key, column in _EXTRA_BY_GAP.items():
        value = _real(crawl.get(gap_key))
        if value and not _real(extra.get(column)):
            extra[column] = value
            changed.append(column)
    if extra != (d.extra_columns or {}):
        d.extra_columns = extra

    return changed
