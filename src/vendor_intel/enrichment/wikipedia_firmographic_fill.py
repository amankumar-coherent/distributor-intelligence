"""Post-enrich: fill Founded + HQ from Wikipedia infoboxes (free, no API key).

Runs after the Wikidata fallback. Wikipedia often has a company infobox with
``founded`` / ``hq_location`` even when the Wikidata SPARQL query misses the
entity (e.g. Amazfit, VikingCloud). We fetch the article wikitext, parse the
infobox, and fall back to the REST summary prose.

Merge-only (never overwrites existing values). Resume-safe: each row is flagged
``wikipedia_firmo_attempted`` and skipped on re-run; ``persist_fn`` is called
after every company.
"""
from __future__ import annotations

import re
from typing import Any

from vendor_intel.clients.http_proxy import httpx_async_client
from vendor_intel.clients.wikipedia_search import wikipedia_search
from vendor_intel.enrichment.gap_fill.gaps import has_founded, has_headquarters

_UA = {"User-Agent": "VendorIntelPipeline/1.0 (research; vendor-intel)"}
_WIKI_API = "https://en.wikipedia.org/w/api.php"
_REST_SUMMARY = "https://en.wikipedia.org/api/rest_v1/page/summary/"
_YEAR_RE = re.compile(r"(18|19|20)\d{2}")

_LEGAL_SUFFIX_RE = re.compile(
    r"[,\s]+(inc\.?|incorporated|corp\.?|corporation|co\.?|company|ltd\.?|limited|"
    r"llc|plc|gmbh|ag|sa|s\.a\.|group|holdings?)\b\.?$",
    re.I,
)


def _clean_company_name(name: str) -> str:
    n = (name or "").strip()
    n = re.sub(r"\s*\([^()]*\)\s*$", "", n).strip(" .,")
    n = _LEGAL_SUFFIX_RE.sub("", n).strip(" .,")
    n = _LEGAL_SUFFIX_RE.sub("", n).strip(" .,")
    return n or (name or "").strip()


def _strip_wikitext(val: str) -> str:
    """Reduce an infobox field value to plain text."""
    s = val or ""
    s = re.sub(r"<ref[^>]*/>", "", s)
    s = re.sub(r"<ref[^>]*>.*?</ref>", "", s, flags=re.S | re.I)
    s = re.sub(r"<!--.*?-->", "", s, flags=re.S)
    s = re.sub(r"<br\s*/?>", ", ", s, flags=re.I)
    # [[link|text]] -> text ; [[text]] -> text
    s = re.sub(r"\[\[(?:[^\]|]*\|)?([^\]|]+)\]\]", r"\1", s)
    # {{nowrap|X}} / {{flag|X}} -> X (drop template name, keep last arg)
    s = re.sub(r"\{\{[^{}]*?\|([^{}|]+)\}\}", r"\1", s)
    s = re.sub(r"\{\{[^{}]*\}\}", "", s)
    s = re.sub(r"</?[^>]+>", "", s)
    s = s.replace("'''", "").replace("''", "")
    s = re.sub(r"\s+", " ", s).strip(" ,.;-")
    return s


def _founded_from_value(val: str) -> str:
    """Extract a founding year from an infobox founded/foundation field."""
    if not val:
        return ""
    # {{Start date and age|2015|9|1}} / {{start date|1969}}
    m = re.search(r"\{\{\s*start date[^}]*?\|\s*((?:18|19|20)\d{2})", val, re.I)
    if m:
        return m.group(1)
    m = _YEAR_RE.search(_strip_wikitext(val))
    return m.group(0) if m else ""


def _parse_infobox(wikitext: str) -> dict[str, str]:
    """Parse founded + headquarters from a company infobox in article wikitext."""
    out: dict[str, str] = {}
    if not wikitext:
        return out

    # Capture infobox fields including multi-line values (value runs until the
    # next line that starts a new "| field =" or closes the template).
    fields: dict[str, str] = {}
    # [^\S\n] = horizontal whitespace only, so an empty field ("| hq_location =")
    # does not swallow the following "| hq_location_city = ..." line.
    for m in re.finditer(
        r"\n\|[^\S\n]*([a-zA-Z_0-9 ]+?)[^\S\n]*=[^\S\n]*((?:[^\n]|\n(?!\s*[|}]))*)",
        wikitext,
    ):
        key = m.group(1).strip().lower().replace(" ", "_")
        if key not in fields:
            fields[key] = m.group(2).strip()

    for k in ("founded", "foundation", "formation", "established", "foundeddate"):
        if fields.get(k):
            yr = _founded_from_value(fields[k])
            if yr:
                out["founded_year"] = yr
                break

    # Headquarters: prefer explicit hq fields, then location.
    city = _strip_wikitext(fields.get("hq_location_city", ""))
    country = _strip_wikitext(fields.get("hq_location_country", ""))
    hq = ""
    if city:
        hq = f"{city}, {country}" if country and country.lower() not in city.lower() else city
    if not hq:
        for k in ("hq_location", "headquarters", "location", "location_city"):
            v = _strip_wikitext(fields.get(k, ""))
            if v and len(v) >= 3:
                lc = _strip_wikitext(fields.get("location_country", "")) if k == "location_city" else ""
                hq = f"{v}, {lc}" if lc and lc.lower() not in v.lower() else v
                break
    if hq and len(hq) >= 3:
        out["headquarters"] = hq
    return out


_FOUNDED_PROSE = re.compile(
    r"(?:\bfounded|\bestablished|\bincorporated)\s+(?:in|on)?\s*"
    r"(?:[A-Z][a-z]+\s+)?(?:\d{1,2},?\s+)?((?:18|19|20)\d{2})",
    re.I,
)
_HQ_PROSE = re.compile(
    r"(?:headquartered|based)\s+in\s+([A-Z][A-Za-z.\-]+(?:,?\s+[A-Z][A-Za-z.\-]+){0,3})",
)


def _parse_summary(extract: str) -> dict[str, str]:
    out: dict[str, str] = {}
    if not extract:
        return out
    m = _FOUNDED_PROSE.search(extract)
    if m:
        out["founded_year"] = m.group(1)
    m = _HQ_PROSE.search(extract)
    if m:
        out["headquarters"] = m.group(1).strip(" ,.")
    return out


def _title_matches(title: str, name: str) -> bool:
    t = (title or "").lower().strip()
    n = (name or "").lower().strip()
    if not t or not n:
        return False
    if t == n or n in t or t in n:
        return True
    # token overlap (e.g. "Bisley International" vs "Bisley (company)")
    nt = {w for w in re.split(r"\W+", n) if len(w) > 2}
    tt = {w for w in re.split(r"\W+", t) if len(w) > 2}
    return bool(nt) and len(nt & tt) >= max(1, len(nt) // 2)


async def _fetch_wikitext(title: str) -> str:
    try:
        async with httpx_async_client(timeout=20.0, headers=_UA) as c:
            r = await c.get(
                _WIKI_API,
                params={
                    "action": "query", "prop": "revisions", "rvprop": "content",
                    "rvslots": "main", "format": "json", "titles": title, "redirects": 1,
                },
            )
            r.raise_for_status()
            pages = (r.json().get("query") or {}).get("pages") or {}
            for p in pages.values():
                revs = p.get("revisions") or []
                if revs:
                    return revs[0].get("slots", {}).get("main", {}).get("*", "") or ""
    except Exception:
        return ""
    return ""


async def _fetch_summary(title: str) -> str:
    try:
        from urllib.parse import quote

        async with httpx_async_client(timeout=20.0, headers=_UA) as c:
            r = await c.get(_REST_SUMMARY + quote(title.replace(" ", "_"), safe=""))
            if r.status_code == 200:
                return str(r.json().get("extract") or "")
    except Exception:
        return ""
    return ""


async def wikipedia_firmographics(name: str) -> dict[str, str]:
    """Return {founded_year, headquarters} from the best-matching Wikipedia article."""
    cleaned = _clean_company_name(name)
    for query in (name, cleaned) if cleaned.lower() != name.lower() else (name,):
        try:
            hits = await wikipedia_search(query, max_results=5)
        except Exception:
            hits = []
        title = next((h.title for h in hits if _title_matches(h.title, cleaned)), "")
        if not title:
            continue
        result: dict[str, str] = {}
        result.update(_parse_infobox(await _fetch_wikitext(title)))
        if not (result.get("founded_year") and result.get("headquarters")):
            summ = _parse_summary(await _fetch_summary(title))
            for k, v in summ.items():
                result.setdefault(k, v)
        if result.get("founded_year") or result.get("headquarters"):
            result["wikipedia_title"] = title
            return result
    return {}


def needs_wikipedia_firmographics(row: dict[str, Any]) -> bool:
    if not row.get("is_relevant", True):
        return False
    if (row.get("_crawl") or {}).get("wikipedia_firmo_attempted"):
        return False
    return not has_founded(row) or not has_headquarters(row)


async def fill_firmographics_via_wikipedia_for_market(
    rows: list[dict[str, Any]],
    *,
    persist_fn: Any = None,
) -> dict[str, Any]:
    """Fill missing Founded / HQ from Wikipedia infoboxes (merge-only)."""
    todo = [r for r in rows if needs_wikipedia_firmographics(r)]
    if not todo:
        return {"queued": 0, "founded_filled": 0, "hq_filled": 0}

    print(f"  [wikipedia_firmo] queue={len(todo)} (missing Founded/HQ after Wikidata)", flush=True)

    def _checkpoint() -> None:
        if persist_fn is None:
            return
        try:
            persist_fn()
        except Exception as exc:
            print(f"  [wikipedia_firmo] checkpoint save failed: {exc}", flush=True)

    founded_filled = hq_filled = errors = 0
    for i, row in enumerate(todo, 1):
        name = str(row.get("company") or row.get("brand") or "").strip()
        crawl = dict(row.get("_crawl") or {})
        if not name:
            crawl["wikipedia_firmo_attempted"] = True
            row["_crawl"] = crawl
            _checkpoint()
            continue
        try:
            info = await wikipedia_firmographics(name)
        except Exception as exc:
            errors += 1
            print(f"  [wikipedia_firmo] [{i}/{len(todo)}] {name[:40]} ERROR: {exc}", flush=True)
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
        crawl["wikipedia_firmo_attempted"] = True
        row["_crawl"] = crawl
        if added:
            print(f"  [wikipedia_firmo] [{i}/{len(todo)}] {name[:40]} +{', '.join(added)}", flush=True)
        _checkpoint()

    return {"queued": len(todo), "founded_filled": founded_filled, "hq_filled": hq_filled, "errors": errors}
