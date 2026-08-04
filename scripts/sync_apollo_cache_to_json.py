#!/usr/bin/env python3
"""Merge apollo_contact_cache + apollo_enriched_snapshot into export JSON + Excel.

Upgrades partial names (Pe***o) to full names when cache/snapshot has better data.
Fills missing email / linkedin from cache or snapshot without overwriting existing values.
"""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT))

from export_founded_hq_v2 import export_one, find_main_json, has_founded, has_headquarters
from vendor_intel.enrichment.gap_fill.gaps import merge_apollo_contact, norm


def _is_partial(name: str) -> bool:
    return "*" in (name or "")


def _best_name(*names: str) -> str:
    cleaned = [norm(n) for n in names if norm(n)]
    if not cleaned:
        return ""
    full = [n for n in cleaned if not _is_partial(n)]
    if full:
        return max(full, key=len)
    return cleaned[0]


def _sync_neon_filtered(market_dir: Path, data: dict, filtered: list[dict]) -> Path:
    ctx = data.get("query_context") or {}
    market = str(ctx.get("industry") or data.get("query") or market_dir.name)
    out = {
        "query": market,
        "query_context": ctx,
        "relevant_companies": filtered,
        "completed_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "filter_note": "Founded + HQ — Apollo cache/snapshot sync",
    }
    path = market_dir / "_neon_filtered_export.json"
    path.write_text(json.dumps(out, indent=2, ensure_ascii=False), encoding="utf-8")
    return path


def _payload_from_sources(cache_entry: dict, snapshot_entry: dict) -> dict[str, str]:
    out: dict[str, str] = {}
    crawl = cache_entry if isinstance(cache_entry, dict) else {}
    snap = snapshot_entry if isinstance(snapshot_entry, dict) else {}

    name = _best_name(
        str(snap.get("contact_person") or ""),
        str(crawl.get("contact_person") or ""),
    )
    if name:
        out["contact_person"] = name

    for field in ("contact_role", "email", "linkedin", "apollo_person_id"):
        for src in (snap, crawl):
            val = norm(src.get(field))
            if val and field not in out:
                out[field] = val
    return out


def main() -> None:
    slug = sys.argv[1] if len(sys.argv) > 1 else "cyber_physical_systems_security"
    market_dir = ROOT / "output" / slug
    json_path = find_main_json(market_dir)
    if not json_path:
        raise SystemExit(f"No export JSON in {market_dir}")

    cache_path = market_dir / "apollo_contact_cache.json"
    snapshot_path = market_dir / "apollo_enriched_snapshot.json"
    cache: dict[str, dict] = {}
    snapshot: dict[str, dict] = {}
    if cache_path.is_file():
        cache = json.loads(cache_path.read_text(encoding="utf-8"))
    if snapshot_path.is_file():
        snapshot = json.loads(snapshot_path.read_text(encoding="utf-8"))

    data = json.loads(json_path.read_text(encoding="utf-8"))
    rows = [
        r
        for r in data.get("relevant_companies") or []
        if r.get("is_relevant", True) and has_founded(r) and has_headquarters(r)
    ]

    updated = 0
    stats = {"name": 0, "role": 0, "email": 0, "linkedin": 0, "partial_upgraded": 0}

    for row in rows:
        key = norm(row.get("company") or row.get("brand")).lower()
        before = dict(row.get("_crawl") or {})
        payload = _payload_from_sources(cache.get(key) or {}, snapshot.get(key) or {})
        if not payload:
            continue

        was_partial = _is_partial(norm(before.get("contact_person")))
        merge_apollo_contact(
            row,
            payload,
            source="apollo_cache_sync",
            fill_empty_only=True,
            fields=frozenset({"contact_person", "contact_role", "email", "linkedin"}),
        )
        # Force full-name upgrade from snapshot when gap-fill blocked partial swap.
        after = dict(row.get("_crawl") or {})
        best = _best_name(
            str(snapshot.get(key, {}).get("contact_person") or ""),
            str(cache.get(key, {}).get("contact_person") or ""),
            str(after.get("contact_person") or ""),
        )
        if best and _is_partial(norm(before.get("contact_person"))) and not _is_partial(best):
            after["contact_person"] = best
            row["_crawl"] = after
            stats["partial_upgraded"] += 1

        changed = any(norm(before.get(f)) != norm(after.get(f)) for f in ("contact_person", "contact_role", "email", "linkedin"))
        if changed:
            updated += 1
            for f in ("contact_person", "contact_role", "email", "linkedin"):
                if norm(before.get(f)) != norm(after.get(f)):
                    stats[f.replace("contact_", "") if f.startswith("contact_") else f] = stats.get(f.replace("contact_", "") if f.startswith("contact_") else f, 0) + 1
            parts = []
            if after.get("contact_person"):
                parts.append(f"name={after.get('contact_person')!r}")
            if after.get("email"):
                parts.append(f"email={after.get('email')}")
            if after.get("linkedin"):
                parts.append("linkedin=yes")
            print(f"  synced {row.get('company')}: {', '.join(parts)}")

    data["apollo_cache_sync"] = {
        "updated": updated,
        "stats": stats,
        "completed_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
    }
    json_path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    neon_path = _sync_neon_filtered(market_dir, data, rows)
    export_stats = export_one(market_dir, suffix="founded_and_hq_v2")

    print(f"\nSynced {updated} companies from cache/snapshot.")
    print(f"Partial names upgraded: {stats.get('partial_upgraded', 0)}")
    print(f"Excel: {export_stats.get('file')}")
    print(f"Neon JSON: {neon_path}")


if __name__ == "__main__":
    main()
