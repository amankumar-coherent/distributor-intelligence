"""Post-enrich: fill Founded + HQ from the company's OWN website (last resort).

Small/regional companies (e.g. Microwear, VikingCloud, Riddhi Chemicals) are
often absent from Wikidata and Wikipedia, yet state their founding year and
location on their own site ("Since 2008", "2008 – Company Founded",
"Established in Shenzhen"). This pass crawls the homepage + About/Company/Our-Story
pages and extracts Founded + HQ.

Merge-only (never overwrites existing values). Resume-safe: each row is flagged
``website_firmo_attempted`` and skipped on re-run; ``persist_fn`` is called after
every company.
"""
from __future__ import annotations

import re
from datetime import datetime
from typing import Any
from urllib.parse import urlparse

from vendor_intel.enrichment.gap_fill.gaps import has_founded, has_headquarters
from vendor_intel.integrations.page_fetch import fetch_page_text

_NOW_YEAR = datetime.now().year


async def _fetch(url: str) -> str:
    """Capable fetch (ddgs → httpx/trafilatura → Jina Reader) for JS-heavy sites."""
    try:
        text, _ = await fetch_page_text(url, max_chars=16000)
        return text or ""
    except Exception:
        return ""

# Founding-year cues, most reliable first.
_FOUNDED_STRONG = re.compile(
    r"(?:founded|established|incorporated|est\.)\s*(?:in|on)?\s*[:\-]?\s*((?:18|19|20)\d{2})",
    re.I,
)
_FOUNDED_COMPANY = re.compile(
    r"company\s+founded\D{0,12}((?:18|19|20)\d{2})", re.I
)
_FOUNDED_SINCE = re.compile(r"\bsince\s+((?:18|19|20)\d{2})\b", re.I)

_HQ_RE = re.compile(
    r"(?:headquartered|head\s*office(?:d)?|based|located|established)\s+in\s+"
    r"([A-Z][A-Za-z.\-]+(?:,?\s+[A-Z][A-Za-z.\-]+){0,3})",
)

_ABOUT_PATHS = ("", "/about", "/about-us", "/pages/about-us", "/company")


def _valid_year(y: str) -> bool:
    try:
        n = int(y)
    except ValueError:
        return False
    return 1800 <= n <= _NOW_YEAR


def _extract_founded(text: str) -> str:
    for rx in (_FOUNDED_STRONG, _FOUNDED_COMPANY):
        for m in rx.finditer(text):
            if _valid_year(m.group(1)):
                return m.group(1)
    # "since YYYY" — accept the earliest plausible (brand tagline year)
    years = [m.group(1) for m in _FOUNDED_SINCE.finditer(text) if _valid_year(m.group(1))]
    if years:
        return min(years)
    return ""


def _extract_hq(text: str) -> str:
    for m in _HQ_RE.finditer(text):
        cand = re.sub(r"\s+", " ", m.group(1)).strip(" ,.;-")
        # Drop verbs / filler mistakenly captured after "based/located in".
        if cand and cand.lower() not in {"the", "our", "a", "order", "stock", "compliance"}:
            return cand
    return ""


def _base_urls(row: dict[str, Any]) -> list[str]:
    raw = str(row.get("website") or row.get("domain") or "").strip()
    if not raw:
        return []
    if "://" not in raw:
        raw = "https://" + raw
    p = urlparse(raw)
    host = p.netloc or p.path
    if not host:
        return []
    base = f"https://{host.rstrip('/')}"
    return [base + path for path in _ABOUT_PATHS]


async def website_firmographics(row: dict[str, Any]) -> dict[str, str]:
    """Crawl the company's site for a founding year + HQ. Stops once both found."""
    need_f = not has_founded(row)
    need_h = not has_headquarters(row)
    out: dict[str, str] = {}
    for url in _base_urls(row):
        text = await _fetch(url)
        if not text:
            continue
        if need_f and not out.get("founded_year"):
            yr = _extract_founded(text)
            if yr:
                out["founded_year"] = yr
        if need_h and not out.get("headquarters"):
            hq = _extract_hq(text)
            if hq:
                out["headquarters"] = hq
        if (not need_f or out.get("founded_year")) and (not need_h or out.get("headquarters")):
            break
    return out


def needs_website_firmographics(row: dict[str, Any]) -> bool:
    if not row.get("is_relevant", True):
        return False
    crawl = row.get("_crawl") or {}
    if crawl.get("website_firmo_attempted"):
        return False
    if not str(row.get("website") or row.get("domain") or "").strip():
        return False
    return not has_founded(row) or not has_headquarters(row)


async def fill_firmographics_via_website_for_market(
    rows: list[dict[str, Any]],
    *,
    persist_fn: Any = None,
) -> dict[str, Any]:
    """Fill missing Founded / HQ by crawling each company's own website (merge-only)."""
    todo = [r for r in rows if needs_website_firmographics(r)]
    if not todo:
        return {"queued": 0, "founded_filled": 0, "hq_filled": 0}

    print(f"  [website_firmo] queue={len(todo)} (crawl own site for Founded/HQ)", flush=True)

    def _checkpoint() -> None:
        if persist_fn is None:
            return
        try:
            persist_fn()
        except Exception as exc:
            print(f"  [website_firmo] checkpoint save failed: {exc}", flush=True)

    founded_filled = hq_filled = errors = 0
    for i, row in enumerate(todo, 1):
        name = str(row.get("company") or row.get("brand") or "").strip()
        crawl = dict(row.get("_crawl") or {})
        try:
            info = await website_firmographics(row)
        except Exception as exc:
            errors += 1
            print(f"  [website_firmo] [{i}/{len(todo)}] {name[:40]} ERROR: {exc}", flush=True)
            info = {}

        added: list[str] = []
        if info.get("founded_year") and not has_founded(row):
            crawl["founded_year"] = str(info["founded_year"])
            added.append(f"founded={info['founded_year']}")
            founded_filled += 1
        if info.get("headquarters") and not has_headquarters(row):
            crawl["headquarters"] = str(info["headquarters"])
            added.append(f"hq={info['headquarters']}")
            hq_filled += 1
        crawl["website_firmo_attempted"] = True
        row["_crawl"] = crawl
        if added:
            print(f"  [website_firmo] [{i}/{len(todo)}] {name[:40]} +{', '.join(added)}", flush=True)
        _checkpoint()

    return {"queued": len(todo), "founded_filled": founded_filled, "hq_filled": hq_filled, "errors": errors}
