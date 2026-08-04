"""Post-enrich: fill Founded + HQ from free Wikidata for rows LinkedIn can't cover.

Many well-known companies do not list a founding year on their LinkedIn page
(e.g. Fitbit, Nokia, Sony). Wikidata (open data, no API key) has structured
inception (P571) and headquarters (P159) for most established brands, so it is
an effective merge-only fallback after the LinkedIn MCP firmographic step.

Merge-only: existing Founded / Headquarters values are never overwritten.
Resume-safe: each row is flagged ``wikidata_firmo_attempted`` and skipped on
re-run; the caller's ``persist_fn`` is invoked after every company.
"""
from __future__ import annotations

import re
from typing import Any

from vendor_intel.enrichment.gap_fill.gaps import has_founded, has_headquarters
from vendor_intel.placeholders import wikidata

_LEGAL_SUFFIX_RE = re.compile(
    r"[,\s]+(inc\.?|incorporated|corp\.?|corporation|co\.?|company|ltd\.?|limited|"
    r"llc|plc|gmbh|ag|sa|s\.a\.|group|holdings?)\b\.?$",
    re.I,
)


def _clean_company_name(name: str) -> str:
    """Normalize a company name for Wikidata matching.

    Strips trailing parentheticals and common legal suffixes, e.g.
      'Under Armour (UA Record)' -> 'Under Armour'
      'Amazfit (Zepp Health Corp.)' -> 'Amazfit'
      'Fitbit, Inc.' -> 'Fitbit'
      'Casio Computer Co., Ltd' -> 'Casio Computer'
    """
    n = (name or "").strip()
    # Drop a trailing "(...)" note (product line / parent brand in parentheses).
    n = re.sub(r"\s*\([^()]*\)\s*$", "", n).strip(" .,")
    n = _LEGAL_SUFFIX_RE.sub("", n).strip(" .,")
    n = _LEGAL_SUFFIX_RE.sub("", n).strip(" .,")  # handle two suffixes ("Casio Computer Co., Ltd")
    return n or (name or "").strip()


def _domain_of(row: dict[str, Any]) -> str:
    return str(row.get("website") or row.get("domain") or "").strip()


def needs_wikidata_firmographics(row: dict[str, Any]) -> bool:
    if not row.get("is_relevant", True):
        return False
    if (row.get("_crawl") or {}).get("wikidata_firmo_attempted"):
        return False
    return not has_founded(row) or not has_headquarters(row)


async def _lookup(name: str, domain: str) -> dict[str, Any] | None:
    """Wikidata lookup with a cleaned-name retry."""
    info = await wikidata.lookup_company(name, domain=domain)
    if info and (info.get("founded_year") or info.get("headquarters")):
        return info
    cleaned = _clean_company_name(name)
    if cleaned and cleaned.lower() != name.lower():
        alt = await wikidata.lookup_company(cleaned, domain=domain)
        if alt and (alt.get("founded_year") or alt.get("headquarters")):
            return alt
    return info


async def fill_firmographics_via_wikidata_for_market(
    rows: list[dict[str, Any]],
    *,
    persist_fn: Any = None,
) -> dict[str, Any]:
    """Fill missing Founded / HQ from Wikidata (merge-only). Returns stats dict."""
    if not wikidata.is_enabled():
        return {"skipped": True, "reason": "wikidata_disabled", "queued": 0}

    todo = [r for r in rows if needs_wikidata_firmographics(r)]
    if not todo:
        return {"queued": 0, "founded_filled": 0, "hq_filled": 0}

    print(
        f"  [wikidata_firmo] queue={len(todo)} (missing Founded/HQ after LinkedIn)",
        flush=True,
    )

    def _checkpoint() -> None:
        if persist_fn is None:
            return
        try:
            persist_fn()
        except Exception as exc:
            print(f"  [wikidata_firmo] checkpoint save failed: {exc}", flush=True)

    founded_filled = hq_filled = errors = 0
    for i, row in enumerate(todo, 1):
        name = str(row.get("company") or row.get("brand") or "").strip()
        crawl = dict(row.get("_crawl") or {})
        if not name:
            crawl["wikidata_firmo_attempted"] = True
            row["_crawl"] = crawl
            _checkpoint()
            continue
        try:
            info = await _lookup(name, _domain_of(row))
        except Exception as exc:
            errors += 1
            print(f"  [wikidata_firmo] [{i}/{len(todo)}] {name[:40]} ERROR: {exc}", flush=True)
            info = None

        added: list[str] = []
        if info:
            if info.get("founded_year") and not has_founded(row):
                crawl["founded_year"] = str(info["founded_year"])
                added.append(f"founded={info['founded_year']}")
                founded_filled += 1
            if info.get("headquarters") and not has_headquarters(row):
                crawl["headquarters"] = str(info["headquarters"])
                added.append(f"hq={info['headquarters']}")
                hq_filled += 1
        crawl["wikidata_firmo_attempted"] = True
        row["_crawl"] = crawl
        if added:
            print(f"  [wikidata_firmo] [{i}/{len(todo)}] {name[:40]} +{', '.join(added)}", flush=True)
        _checkpoint()

    return {
        "queued": len(todo),
        "founded_filled": founded_filled,
        "hq_filled": hq_filled,
        "errors": errors,
    }
