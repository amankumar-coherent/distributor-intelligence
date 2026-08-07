"""Shared "run one market, write its exports" glue.

Extracted from run_query.py so the same logic backs both the interactive/
manual CLI (run_query.py, unchanged behavior) and the batch driver
(run_batch.py) — no more copy-pasting this per market script.
"""
from __future__ import annotations

import asyncio
import json
import re
from pathlib import Path
from typing import Any

from vendor_intel.config import Settings, _project_root


def slugify(query: str, country: str) -> str:
    text = f"{query}_{country}".lower()
    return "".join(c if c.isalnum() else "_" for c in text).strip("_")[:80]


# Generic slug tokens that must NOT alone match curated query/seed files.
# Without this, "Global Avocado Oil Market" matches global_firewall_market.txt
# via shared {global, market} and appends firewall/MSP phrases onto oil searches.
_GENERIC_MARKET_TOKENS = frozenset(
    {
        "global",
        "world",
        "worldwide",
        "international",
        "market",
        "markets",
        "industry",
        "industries",
        "sector",
        "sectors",
        "report",
        "analysis",
        "outlook",
        "forecast",
        "size",
        "share",
        "growth",
        "trends",
        "the",
        "and",
        "for",
        "with",
    }
)


def _slug_tokens(text: str) -> set[str]:
    return {t for t in text.lower().split("_") if len(t) > 2}


def _distinctive_tokens(tokens: set[str]) -> set[str]:
    return {t for t in tokens if t not in _GENERIC_MARKET_TOKENS}


def _best_curated_file(run_slug: str, files: list[Path]) -> Path | None:
    """Pick curated .txt whose distinctive market tokens match the run slug.

    Every non-generic token in the filename must appear in the market slug
    (e.g. ``global_firewall_market`` needs ``firewall``). Shared tokens like
    ``global`` / ``market`` alone never match — that previously attached
    firewall/MSP queries onto avocado oil and similar markets.
    """
    run_distinct = _distinctive_tokens(_slug_tokens(run_slug))
    if not run_distinct:
        return None

    best_file: Path | None = None
    best_score = 0
    for f in files:
        f_distinct = _distinctive_tokens(_slug_tokens(f.stem))
        if not f_distinct:
            continue
        overlap = f_distinct & run_distinct
        if overlap != f_distinct:
            continue
        score = len(overlap)
        if score > best_score:
            best_file, best_score = f, score
    return best_file


def parse_sections(raw: str | None) -> list[str]:
    """Parse a CEO section list. Separator is ';' if present, else ','."""
    if not raw or not raw.strip():
        return []
    sep = ";" if ";" in raw else ","
    return [s.strip() for s in raw.split(sep) if s.strip()]


_SEED_DOMAIN_RE = re.compile(
    r"^(?:https?://)?(?:www\.)?[a-z0-9][a-z0-9-]*(?:\.[a-z0-9-]+)+/?$", re.I
)


def _looks_like_domain(s: str) -> bool:
    s = (s or "").strip()
    return bool(s) and " " not in s and bool(_SEED_DOMAIN_RE.match(s))


def _clean_domain(s: str) -> str:
    return (
        (s or "")
        .strip()
        .lower()
        .replace("https://", "")
        .replace("http://", "")
        .removeprefix("www.")
        .rstrip("/")
    )


def load_seeds_file(path: str | None) -> list[tuple[str, str | None, str | None]]:
    """Read must-have companies from a file. Each line is one company, optionally with a
    pinned official domain and/or an intended section, in any order, after '|':

        Ingram Micro | ingrammicro.com
        Phillips 66 | Recovered Oil / Base Oil Downstream Offtake
        Redington | redingtongroup.com | Regional Distribution

    A field that looks like a bare domain (no spaces, has a dot) is treated as the
    pinned domain (skips the web lookup); any other field is the section.
    Returns [(name, domain_or_None, section_or_None)]. '#' comments ignored."""
    if not path or not path.strip():
        return []
    p = Path(path)
    if not p.exists():
        print(f"  WARNING: seeds file not found: {path}")
        return []
    rows: list[tuple[str, str | None, str | None]] = []
    for line in p.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if "|" in line:
            parts = [f.strip() for f in line.split("|")]
            name = parts[0]
            domain: str | None = None
            section: str | None = None
            for f in parts[1:]:
                if not f:
                    continue
                if domain is None and _looks_like_domain(f):
                    domain = _clean_domain(f)
                elif section is None:
                    section = f
            if name:
                rows.append((name, domain, section))
        else:
            for n in parse_sections(line):  # legacy: ';'/','-separated names, no section
                rows.append((n, None, None))
    return rows


def auto_seed_names(query: str, country: str) -> list[tuple[str, str | None, str | None]]:
    """Find the curated seed file for this market by distinctive keyword match.

    Ignores generic tokens (global/market/…) so unrelated seed files are not
    attached via weak overlap.
    """
    seeds_dir = _project_root() / "queries" / "seeds"
    if not seeds_dir.exists():
        return []
    best_file = _best_curated_file(
        slugify(query, country),
        list(seeds_dir.glob("*.txt")),
    )
    return load_seeds_file(str(best_file)) if best_file else []


def auto_discovery_queries(query: str, country: str) -> list[str]:
    """Find curated extra discovery queries for this market.

    File convention: queries/discovery/<market_slug>.txt — one search phrase per
    line ('#' comments and blanks ignored). These run in addition to the built-in
    discovery templates (see multi_source_discovery.discover_additional_sources).

    Matching requires distinctive market tokens (not just global/market), so
    avocado oil cannot inherit firewall/MSP query lists.
    """
    qdir = _project_root() / "queries" / "discovery"
    if not qdir.exists():
        return []
    best_file = _best_curated_file(
        slugify(query, country),
        list(qdir.glob("*.txt")),
    )
    if not best_file:
        return []
    out: list[str] = []
    for line in best_file.read_text(encoding="utf-8").splitlines():
        s = line.split("#", 1)[0].strip()
        if s:
            out.append(s)
    return list(dict.fromkeys(out))


# Marker that introduces the desired functional categories inside a query string.
_FUNC_MARKER = re.compile(
    r"\b(?:functionality|sections?|include|company\s+types?|roles?|profiles?)\s*[:=\-]\s*(.+)$",
    re.I,
)


def parse_market_line(raw: str) -> tuple[str, str, list[str]]:
    """Parse one detailed query string into (market, country, sections).

    Accepts, in one input box:
      - "Market | Country | Section A; Section B; Section C"
      - "Market | Country"
      - "Market ; functionality: Section A, Section B, Section C"
      - "Market | Country ; sections: A; B; C"
    The functionality/sections clause may use the keywords functionality / sections /
    include / company types / roles / profiles followed by ':' '=' or '-'.
    """
    raw = (raw or "").strip()
    sections: list[str] = []
    country = ""

    m = _FUNC_MARKER.search(raw)
    if m:
        sections = parse_sections(m.group(1))
        raw = raw[: m.start()].strip().rstrip("|,;:- ").strip()

    if "|" in raw:
        parts = [p.strip() for p in raw.split("|")]
        market = parts[0]
        if len(parts) > 1 and parts[1]:
            country = parts[1]
        if len(parts) > 2 and not sections:
            sections = parse_sections(parts[2])
    else:
        market = raw

    return market, (country or "global"), sections


async def execute_and_export_async(
    query: str,
    country: str,
    settings: Settings,
    output_dir: Path,
    *,
    enrich_limit: int | None = None,
    classify_limit: int | None = None,
    sections: list[str] | None = None,
    seeds: list[str] | None = None,
    out_name: str | None = None,
    fmcg_xlsx: bool = False,
    market_type: str = "fmcg",
    resume_checkpoint: bool = False,
    market_slug: str | None = None,
) -> dict[str, Any]:
    """Run the pipeline for one market and write CSV/DOCX/(optional XLSX)/JSON
    into `output_dir`. Returns the pipeline result dict with the written paths
    stashed under `_csv_path` / `_docx_path` / `_xlsx_path` / `_json_path`.

    Awaits the pipeline directly (no nested asyncio.run) so it can run as one
    of several concurrent markets on a single shared event loop — see
    batch/orchestrator.py. `execute_and_export` below is a sync wrapper around
    this for callers (run_query.py) that aren't already inside a running loop.
    """
    from vendor_intel.pipeline.orchestrator import (
        run_pipeline,
        save_pipeline_csv,
        save_pipeline_docx,
    )

    query = query.strip()
    country = country.strip() or "global"
    # Auto-load curated must-have companies for this market by convention, so the run
    # command stays clean (no seed list shown). File: queries/seeds/<market_slug>.txt
    # Each seed may carry an intended section ("Name | Section") which is honored at export.
    merged_seeds: list[str] = list(seeds or [])
    seed_sections: dict[str, str] = {}
    seed_domains: dict[str, str] = {}
    auto = auto_seed_names(query, country)
    for name, dom, sec in auto:
        if name not in merged_seeds:
            merged_seeds.append(name)
        if sec:
            seed_sections[name] = sec
        if dom:
            seed_domains[name] = dom
    if auto:
        print(
            f"  [seeds] auto-loaded {len(auto)} curated companies "
            f"({len(seed_domains)} domain-pinned, {len(seed_sections)} with a fixed section)",
            flush=True,
        )
    elif not merged_seeds:
        print("  [seeds] no curated seed file matched this market (pure discovery)", flush=True)
    merged_seeds = list(dict.fromkeys(merged_seeds))
    extra_queries = auto_discovery_queries(query, country)
    if extra_queries:
        print(
            f"  [discovery] auto-loaded {len(extra_queries)} curated discovery queries",
            flush=True,
        )
    query_context = {
        "industry": query,
        "country": country,
        "functions": [],
        "sections": list(sections or []),
        "seed_companies": merged_seeds,
        "seed_sections": seed_sections,
        "seed_domains": seed_domains,
        "extra_discovery_queries": extra_queries,
        "resume_checkpoint": bool(resume_checkpoint),
        "market_slug": (market_slug or "").strip() or None,
    }
    result = await run_pipeline(
        query_context,
        settings,
        enrich_limit=enrich_limit,
        classify_limit=classify_limit,
    )

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    if out_name:
        base = Path(str(out_name)).stem  # drop any extension
        slug = "".join(c if c.isalnum() else "_" for c in base.lower()).strip("_")[:90]
    else:
        slug = slugify(query, country)
    # Batch worker already passes output/<market_slug> — do not nest again as
    # output/<slug>/<slug>/ (that hid JSON from post_market LinkedIn/Excel refresh).
    requested_slug = (market_slug or "").strip()
    if requested_slug and output_dir.name == requested_slug:
        market_dir = output_dir
        slug = requested_slug
    else:
        market_dir = output_dir / slug
        market_dir.mkdir(parents=True, exist_ok=True)
    csv_path = market_dir / f"{slug}.csv"
    json_path = market_dir / f"{slug}.json"
    docx_path = market_dir / f"{slug}.docx"
    save_pipeline_csv(result, str(csv_path))
    try:
        save_pipeline_docx(result, str(docx_path))
        result["_docx_path"] = str(docx_path)
    except Exception as exc:
        print(f"  [pipeline] DOCX export failed: {exc}", flush=True)

    # Channel-partner purity before Excel (wide discovery earlier; clean export here)
    try:
        from vendor_intel.intelligence.channel_purity import (
            channel_purity_enabled,
            filter_channel_pure_rows,
        )
        from vendor_intel.pipeline.web_expand import market_family as _market_family

        if channel_purity_enabled():
            family = _market_family(query)
            before = list(result.get("relevant_companies") or [])
            kept, dropped = filter_channel_pure_rows(
                before, family=family, query=query
            )
            result["relevant_companies"] = kept
            result["channel_purity"] = {
                "before": len(before),
                "kept": len(kept),
                "dropped": len(dropped),
                "dropped_sample": dropped[:50],
            }
            print(
                f"  [pipeline] Channel purity: kept {len(kept)}/{len(before)} "
                f"(dropped {len(dropped)} OEM/utility/non-channel)",
                flush=True,
            )
            if dropped:
                purity_path = market_dir / f"{slug}_channel_purity_dropped.json"
                purity_path.write_text(
                    json.dumps(dropped, indent=2, ensure_ascii=False), encoding="utf-8"
                )
    except Exception as exc:
        print(f"  [pipeline] Channel purity skipped: {exc}", flush=True)

    if fmcg_xlsx:
        try:
            from vendor_intel.export.fmcg_landscape import export_fmcg_xlsx

            xlsx_path = market_dir / f"{slug}.xlsx"
            result["_xlsx_path"] = export_fmcg_xlsx(result, xlsx_path, market_type=market_type)
        except Exception as exc:
            print(f"  [pipeline] FMCG Excel export failed: {exc}", flush=True)
    try:
        json_path.write_text(json.dumps(result, indent=2, default=str), encoding="utf-8")
    except Exception:
        pass
    result["_csv_path"] = str(csv_path)
    result["_json_path"] = str(json_path)
    result["_market_dir"] = str(market_dir)
    return result


def execute_and_export(
    query: str,
    country: str,
    settings: Settings,
    output_dir: Path,
    *,
    enrich_limit: int | None = None,
    classify_limit: int | None = None,
    sections: list[str] | None = None,
    seeds: list[str] | None = None,
    out_name: str | None = None,
    fmcg_xlsx: bool = False,
    market_type: str = "fmcg",
    resume_checkpoint: bool = False,
    market_slug: str | None = None,
) -> dict[str, Any]:
    """Sync wrapper around execute_and_export_async — for callers that are NOT
    already inside a running event loop (run_query.py's interactive/CLI modes).
    Do not call this from within async code; await execute_and_export_async
    directly instead (asyncio.run() cannot nest inside a running loop).
    """
    return asyncio.run(
        execute_and_export_async(
            query,
            country,
            settings,
            output_dir,
            enrich_limit=enrich_limit,
            classify_limit=classify_limit,
            sections=sections,
            seeds=seeds,
            out_name=out_name,
            fmcg_xlsx=fmcg_xlsx,
            market_type=market_type,
            resume_checkpoint=resume_checkpoint,
            market_slug=market_slug,
        )
    )
