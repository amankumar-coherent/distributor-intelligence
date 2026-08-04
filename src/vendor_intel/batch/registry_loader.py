"""Populate the SQLite market registry (batch/db.py) from a market-list input file.

Two input formats, auto-detected by file extension:

  .txt — one market per line, the same "Query | Country | Section1; Section2"
         convention already used by queries/markets.txt (parsed with
         pipeline.run_and_export.parse_market_line — identical behavior to
         run_query.py's single-query/--file modes). market_type defaults to
         "general", priority to 100. '#' lines and blanks are skipped.

  .csv — columns: query, country, market_type, sections, seeds_file, priority
         (all but `query` optional) — the realistic format for a
         100-10,000-row, spreadsheet-authored market list.
"""
from __future__ import annotations

import csv
from pathlib import Path

from vendor_intel.batch import db
from vendor_intel.pipeline.run_and_export import parse_market_line, parse_sections, slugify


def _row_from_txt_line(raw: str) -> dict | None:
    raw = raw.strip()
    if not raw or raw.startswith("#"):
        return None
    query, country, sections = parse_market_line(raw)
    if not query:
        return None
    return {
        "query": query,
        "country": country,
        "market_type": "general",
        "sections": sections,
        "seeds_file": None,
        "priority": 100,
    }


def _row_from_csv_row(row: dict) -> dict | None:
    # Accept query OR name (neon export uses `name`)
    query = (row.get("query") or row.get("name") or "").strip()
    if not query:
        return None
    country = (row.get("country") or "global").strip() or "global"
    market_type = (row.get("market_type") or row.get("sector_code") or "general").strip() or "general"
    sections = parse_sections(row.get("sections"))
    seeds_file = (row.get("seeds_file") or "").strip() or None
    priority_raw = (row.get("priority") or "").strip()
    try:
        priority = int(priority_raw) if priority_raw else 100
    except ValueError:
        priority = 100
    return {
        "query": query,
        "country": country,
        "market_type": market_type,
        "sections": sections,
        "seeds_file": seeds_file,
        "priority": priority,
    }


def parse_input_file(path: Path) -> list[dict]:
    """Parse `path` into a flat list of market row dicts (query/country/
    market_type/sections/seeds_file/priority). Does not touch the database."""
    rows: list[dict] = []
    if path.suffix.lower() == ".csv":
        with path.open(encoding="utf-8-sig", newline="") as f:
            for row in csv.DictReader(f):
                parsed = _row_from_csv_row(row)
                if parsed:
                    rows.append(parsed)
    else:
        for line in path.read_text(encoding="utf-8").splitlines():
            parsed = _row_from_txt_line(line)
            if parsed:
                rows.append(parsed)
    return rows


def populate_registry(
    input_path: str | Path,
    batch_id: str,
    *,
    profile: str = "quality",
    cap: str = "standard",
) -> dict[str, int]:
    """Load `input_path` and upsert every row into the market registry.

    Idempotent on market_slug (query+country) — re-running against the same
    file resumes a batch rather than duplicating rows: a market already
    in-flight or done keeps its status (see db.upsert_market).
    """
    path = Path(input_path)
    if not path.exists():
        raise FileNotFoundError(f"market list not found: {path}")

    rows = parse_input_file(path)
    if not rows:
        raise ValueError(f"no markets parsed from {path}")

    db.create_batch(batch_id, input_file=str(path), total_markets=len(rows))

    inserted = 0
    updated = 0
    skipped_duplicate_in_file = 0
    seen_slugs: set[str] = set()
    for row in rows:
        slug = slugify(row["query"], row["country"])
        if slug in seen_slugs:
            # Same (query, country) appears twice in this file — keep the
            # first occurrence's priority/sections rather than upserting twice.
            skipped_duplicate_in_file += 1
            continue
        seen_slugs.add(slug)
        _market_id, was_new = db.upsert_market(
            batch_id,
            query=row["query"],
            country=row["country"],
            market_slug=slug,
            market_type=row["market_type"],
            sections=row["sections"],
            seeds_file=row["seeds_file"],
            profile=profile,
            cap=cap,
            priority=row["priority"],
        )
        if was_new:
            inserted += 1
        else:
            updated += 1

    return {
        "total": len(rows),
        "inserted": inserted,
        "updated": updated,
        "skipped_duplicate_in_file": skipped_duplicate_in_file,
    }
