#!/usr/bin/env python3
"""Push a vendor-intel pipeline JSON export into Neon Postgres.

Mapping (agent → Neon):
  company / brand     → companies.name  (UNIQUE upsert)
  geography / scope   → companies.country
  market query        → markets.name    (UNIQUE upsert)
  company ↔ market    → market_company_link
  website, role, …    → company_metadata (source=vendor_intel, data=JSONB)

Usage:
  python scripts/push_export_to_neon.py output/demo/medical_devices_distributors_india.json
  python scripts/push_export_to_neon.py output/demo/foo.json --dry-run
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(ROOT / "src"))

from dotenv import load_dotenv

load_dotenv(ROOT / ".env")

SOURCE = "vendor_intel"


def _database_url() -> str:
    url = (os.environ.get("DATABASE_URL") or "").strip()
    if not url:
        raise SystemExit(
            "DATABASE_URL is not set. Add it to agents/vendor-intel/.env"
        )
    return url


def _connect():
    import psycopg
    from psycopg.types.json import Jsonb

    return psycopg.connect(_database_url()), Jsonb


def _load_export(path: Path) -> dict[str, Any]:
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise SystemExit(f"Expected JSON object in {path}")
    return data


def _rows(data: dict[str, Any]) -> list[dict[str, Any]]:
    rows = data.get("relevant_companies") or []
    return [r for r in rows if isinstance(r, dict) and (r.get("company") or r.get("brand"))]


def _market_name(data: dict[str, Any]) -> str:
    """Canonical Neon export key — must align with CMI report.keyword for market matching."""
    query = (data.get("query") or data.get("query_context") or "").strip()
    if query:
        return query

    scope = data.get("scope") or {}
    market = (scope.get("market") or "").strip()
    geo = ""
    geos = scope.get("geographies") or []
    if isinstance(geos, list) and geos:
        geo = str(geos[0]).strip()
    elif isinstance(geos, str):
        geo = geos.strip()
    if market and geo:
        return f"{market} ({geo})"
    if market:
        return market
    return "Untitled market"


def _country_for_row(row: dict[str, Any], data: dict[str, Any]) -> str | None:
    signals = row.get("signals") or {}
    hq = (signals.get("hq_country") or "").strip() if isinstance(signals, dict) else ""
    if hq:
        return hq
    geos = (data.get("scope") or {}).get("geographies") or []
    if isinstance(geos, list) and geos:
        return str(geos[0]).strip() or None
    return None


def _metadata_payload(row: dict[str, Any], data: dict[str, Any]) -> dict[str, Any]:
    """Persist the full pipeline company row (+ run context) into company_metadata.data."""
    payload: dict[str, Any] = {}
    for key, value in row.items():
        if key == "_crawl" and isinstance(value, dict):
            crawl = dict(value)
            text = crawl.get("crawl_text")
            if isinstance(text, str) and len(text) > 12_000:
                crawl["crawl_text"] = text[:12_000] + "…"
            payload["_crawl"] = crawl
        else:
            payload[key] = value

    payload.setdefault("website", row.get("website") or row.get("domain"))
    payload["market_query"] = data.get("query")
    payload["pipeline_profile"] = data.get("pipeline_profile")
    payload["market_name"] = _market_name(data)
    payload["export_scope"] = data.get("scope")
    payload["query_context"] = data.get("query_context")
    payload["completed_at"] = data.get("completed_at")
    return payload


def _geography(data: dict[str, Any]) -> str | None:
    scope = data.get("scope") or {}
    geos = scope.get("geographies") or []
    if isinstance(geos, list) and geos:
        return str(geos[0]).strip() or None
    if isinstance(geos, str):
        return geos.strip() or None
    return None


def _push_vendor_intel_tables(
    cur,
    Jsonb,
    *,
    path: Path,
    data: dict[str, Any],
    rows: list[dict[str, Any]],
    market_name: str,
    merge_companies: bool = False,
) -> dict[str, int]:
    from vendor_intel_tables import (
        ensure_tables,
        flatten_company_row,
        merge_company_flat,
        merge_export_payload,
    )

    ensure_tables(cur)
    geo = _geography(data)
    query_ctx = data.get("query_context")
    scope = data.get("scope")
    completed = data.get("completed_at")

    cur.execute(
        """
        INSERT INTO vendor_intel_exports (
          market_name, query, geography, export_file, pipeline_profile,
          scope, query_context, company_count, completed_at, pushed_at, updated_at
        )
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, NOW(), NOW())
        ON CONFLICT (market_name) DO UPDATE SET
          query = EXCLUDED.query,
          geography = EXCLUDED.geography,
          export_file = EXCLUDED.export_file,
          pipeline_profile = EXCLUDED.pipeline_profile,
          scope = EXCLUDED.scope,
          query_context = EXCLUDED.query_context,
          company_count = EXCLUDED.company_count,
          completed_at = EXCLUDED.completed_at,
          pushed_at = NOW(),
          updated_at = NOW()
        RETURNING id
        """,
        (
            market_name,
            data.get("query"),
            geo,
            str(path),
            data.get("pipeline_profile"),
            Jsonb(scope if isinstance(scope, dict) else {}),
            Jsonb(query_ctx if isinstance(query_ctx, dict) else {}),
            len(rows),
            completed,
        ),
    )
    export_id = cur.fetchone()[0]
    if not merge_companies:
        cur.execute("DELETE FROM vendor_intel_companies WHERE export_id = %s", (export_id,))

    inserted = 0
    merged = 0
    for idx, row in enumerate(rows):
        flat = flatten_company_row(row)
        name = flat["company_name"]
        if not name:
            continue
        payload = _metadata_payload(row, data)

        if merge_companies:
            cur.execute(
                """
                SELECT brand, country, role, website, key_products, summary,
                       company_function, parent, confidence, quality_score, export_data
                FROM vendor_intel_companies
                WHERE export_id = %s AND company_name = %s
                LIMIT 1
                """,
                (export_id, name),
            )
            existing_row = cur.fetchone()
            if existing_row:
                existing = {
                    "brand": existing_row[0],
                    "country": existing_row[1],
                    "role": existing_row[2],
                    "website": existing_row[3],
                    "key_products": existing_row[4],
                    "summary": existing_row[5],
                    "company_function": existing_row[6],
                    "parent": existing_row[7],
                    "confidence": existing_row[8],
                    "quality_score": existing_row[9],
                    "export_data": existing_row[10],
                }
                payload = merge_export_payload(existing.get("export_data"), payload)
                flat = merge_company_flat(existing, flat)
                merged += 1

        flat["export_data"] = payload
        cur.execute(
            """
            INSERT INTO vendor_intel_companies (
              export_id, sort_order, company_name, brand, country, role, website,
              key_products, summary, company_function, parent, confidence, quality_score, export_data
            )
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            ON CONFLICT (export_id, company_name) DO UPDATE SET
              sort_order = EXCLUDED.sort_order,
              brand = EXCLUDED.brand,
              country = EXCLUDED.country,
              role = EXCLUDED.role,
              website = EXCLUDED.website,
              key_products = EXCLUDED.key_products,
              summary = EXCLUDED.summary,
              company_function = EXCLUDED.company_function,
              parent = EXCLUDED.parent,
              confidence = EXCLUDED.confidence,
              quality_score = EXCLUDED.quality_score,
              export_data = EXCLUDED.export_data,
              pushed_at = NOW()
            """,
            (
                export_id,
                idx,
                name,
                flat["brand"],
                flat["country"],
                flat["role"],
                flat["website"],
                flat["key_products"],
                flat["summary"],
                flat["company_function"],
                flat["parent"],
                flat["confidence"],
                flat["quality_score"],
                Jsonb(payload),
            ),
        )
        inserted += 1

    return {
        "export_id": str(export_id),
        "vendor_intel_companies": inserted,
        "vendor_intel_companies_merged": merged,
    }


def should_push_neon() -> bool:
    return bool((os.environ.get("DATABASE_URL") or "").strip())


def push_export(path: Path, *, dry_run: bool = False, merge_companies: bool = False) -> dict[str, int]:
    data = _load_export(path)
    rows = _rows(data)
    market_name = _market_name(data)

    stats = {
        "companies_upserted": 0,
        "markets_upserted": 0,
        "links_upserted": 0,
        "metadata_upserted": 0,
        "vendor_intel_exports": 0,
        "vendor_intel_companies": 0,
        "vendor_intel_companies_merged": 0,
        "skipped": 0,
        "total_rows": len(rows),
    }

    if not rows:
        print(f"No relevant_companies in {path}")
        return stats

    print(f"Market : {market_name}")
    print(f"Rows   : {len(rows)}")
    if merge_companies:
        print("Mode   : merge (fill empty fields only)")
    if dry_run:
        print("[dry-run] No database writes.")
        for row in rows[:5]:
            print(f"  - {row.get('company')} | {row.get('role')} | {row.get('website')}")
        if len(rows) > 5:
            print(f"  … +{len(rows) - 5} more")
        return stats

    conn, Jsonb = _connect()
    try:
        from vendor_intel_tables import merge_export_payload

        with conn.cursor() as cur:
            cur.execute(
                """
                INSERT INTO markets (id, name, status, created_at, updated_at)
                VALUES (gen_random_uuid(), %s, 'active', NOW(), NOW())
                ON CONFLICT (name) DO UPDATE
                  SET status = EXCLUDED.status,
                      updated_at = NOW()
                RETURNING id
                """,
                (market_name,),
            )
            market_id = cur.fetchone()[0]
            stats["markets_upserted"] = 1

            for row in rows:
                name = (row.get("brand") or row.get("company") or "").strip()
                if not name:
                    stats["skipped"] += 1
                    continue
                country = _country_for_row(row, data)

                cur.execute(
                    """
                    INSERT INTO companies (id, name, country, has_public_esg, esg_scoring, created_at)
                    VALUES (gen_random_uuid(), %s, %s, false, 'pending', NOW())
                    ON CONFLICT (name) DO UPDATE
                      SET country = COALESCE(EXCLUDED.country, companies.country)
                    RETURNING id
                    """,
                    (name, country),
                )
                company_id = cur.fetchone()[0]
                stats["companies_upserted"] += 1

                cur.execute(
                    """
                    INSERT INTO market_company_link (market_id, company_id)
                    VALUES (%s, %s)
                    ON CONFLICT (market_id, company_id) DO NOTHING
                    """,
                    (market_id, company_id),
                )
                if cur.rowcount:
                    stats["links_upserted"] += 1

                payload = _metadata_payload(row, data)
                if merge_companies:
                    cur.execute(
                        """
                        SELECT data FROM company_metadata
                        WHERE company_id = %s AND source = %s
                        LIMIT 1
                        """,
                        (company_id, SOURCE),
                    )
                    meta_row = cur.fetchone()
                    if meta_row and meta_row[0]:
                        payload = merge_export_payload(meta_row[0], payload)

                cur.execute(
                    """
                    INSERT INTO company_metadata (company_id, source, data, fetched_at)
                    VALUES (%s, %s, %s, NOW())
                    ON CONFLICT (company_id) DO UPDATE
                      SET source = EXCLUDED.source,
                          data = EXCLUDED.data,
                          fetched_at = NOW()
                    """,
                    (company_id, SOURCE, Jsonb(payload)),
                )
                stats["metadata_upserted"] += 1

            vic = _push_vendor_intel_tables(
                cur,
                Jsonb,
                path=path,
                data=data,
                rows=rows,
                market_name=market_name,
                merge_companies=merge_companies,
            )
            stats["vendor_intel_exports"] = 1
            stats["vendor_intel_companies"] = vic["vendor_intel_companies"]
            stats["vendor_intel_companies_merged"] = vic.get("vendor_intel_companies_merged", 0)

        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

    return stats


def main() -> None:
    parser = argparse.ArgumentParser(description="Push vendor-intel JSON export to Neon")
    parser.add_argument("json_path", help="Path to pipeline JSON export")
    parser.add_argument("--dry-run", action="store_true", help="Print plan only")
    parser.add_argument(
        "--merge",
        action="store_true",
        help="Merge into Neon (fill empty fields only). Default: replace all companies for the market.",
    )
    args = parser.parse_args()

    path = Path(args.json_path)
    if not path.is_absolute():
        path = (Path.cwd() / path).resolve()
    if not path.exists():
        raise SystemExit(f"File not found: {path}")

    stats = push_export(path, dry_run=args.dry_run, merge_companies=args.merge)
    print("----")
    for k, v in stats.items():
        print(f"  {k}: {v}")
    print("Done.")


if __name__ == "__main__":
    main()
