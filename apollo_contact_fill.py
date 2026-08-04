#!/usr/bin/env python3
"""Fill contact person name from Apollo.io for a market export.

Default: gap-fill only (keep existing names), free People Search API, contact_person only.

Usage:
  .venv\\Scripts\\python.exe apollo_contact_fill.py cyber_physical_systems_security
  .venv\\Scripts\\python.exe apollo_contact_fill.py cyber_physical_systems_security --enrich
  .venv\\Scripts\\python.exe apollo_contact_fill.py cyber_physical_systems_security --replace-all
  .venv\\Scripts\\python.exe apollo_contact_fill.py cyber_physical_systems_security --limit 5
  .venv\\Scripts\\python.exe apollo_contact_fill.py cyber_physical_systems_security --dry-run
"""
from __future__ import annotations

import argparse
import asyncio
import json
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT))

from export_founded_hq_v2 import export_one, find_main_json, has_founded, has_headquarters  # noqa: E402

CONTACT_BUNDLE_FIELDS = frozenset({"contact_person", "email", "linkedin"})
ROLE_ONLY_FIELDS = frozenset({"contact_role"})


def _load_protected_apollo_keys(market_dir: Path) -> frozenset[str]:
    """Companies from first Apollo enrich run — never overwrite."""
    path = market_dir / "apollo_enriched_snapshot.json"
    if not path.is_file():
        return frozenset()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return frozenset()
    return frozenset(str(k).strip().lower() for k in data.keys())


def _fields_label(fields: frozenset[str]) -> str:
    if fields == CONTACT_BUNDLE_FIELDS:
        return "contact_person + email + linkedin"
    if fields == frozenset({"contact_person", "contact_role"}):
        return "contact_person + contact_role (Role column)"
    if fields == ROLE_ONLY_FIELDS:
        return "contact_role only (job title)"
    if fields == frozenset({"contact_person"}):
        return "contact_person only"
    return ", ".join(sorted(fields))


def _load_rows(
    slug: str,
    *,
    founded_hq_only: bool,
    market_dir: Path | None = None,
) -> tuple[Path, dict, list[dict]]:
    if market_dir is not None:
        md = Path(market_dir)
        json_path = find_main_json(md)
        if not json_path:
            raise SystemExit(f"No export JSON in {market_dir}")
    else:
        md = ROOT / "output" / slug
        json_path = find_main_json(md)
        if not json_path:
            raise SystemExit(f"No export JSON in {md}")
    data = json.loads(json_path.read_text(encoding="utf-8"))
    rows = [r for r in (data.get("relevant_companies") or []) if r.get("is_relevant", True)]
    if founded_hq_only:
        rows = [r for r in rows if has_founded(r) and has_headquarters(r)]
    return json_path, data, rows


def _sync_neon_filtered(market_dir: Path, data: dict, filtered: list[dict]) -> Path:
    ctx = data.get("query_context") or {}
    market = str(ctx.get("industry") or data.get("query") or market_dir.name)
    out = {
        "query": market,
        "query_context": ctx,
        "relevant_companies": filtered,
        "completed_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "filter_note": "Founded + HQ — Apollo contact name fill",
    }
    path = market_dir / "_neon_filtered_export.json"
    path.write_text(json.dumps(out, indent=2, ensure_ascii=False), encoding="utf-8")
    return path


async def run(
    slug: str,
    *,
    founded_hq_only: bool,
    replace_all: bool,
    limit: int | None,
    concurrent: int,
    enrich_match: bool,
    merge_fields: frozenset[str],
    dry_run: bool,
    push_neon: bool,
    request_delay: float,
    resume: bool,
    skip_apollo_filled: bool,
    companies: list[str] | None = None,
    market_dir: Path | None = None,
) -> dict:
    from vendor_intel.enrichment.apollo_contacts import ApolloContactClient, company_domain, contact_names_match
    from vendor_intel.enrichment.gap_fill.gaps import (
        has_contact_person,
        is_apollo_found_contact,
        is_apollo_protected_row,
        is_empty_value,
        merge_apollo_contact,
        needs_apollo_fields,
        norm,
    )

    json_path, data, rows = _load_rows(slug, founded_hq_only=founded_hq_only, market_dir=market_dir)
    market_dir = json_path.parent
    fill_empty_only = not replace_all
    protected_keys = _load_protected_apollo_keys(market_dir) if skip_apollo_filled else frozenset()

    cache_path = market_dir / "apollo_contact_cache.json"
    progress_path = market_dir / "apollo_contact_progress.json"
    cache: dict[str, dict] = {}
    if cache_path.is_file():
        try:
            cache = json.loads(cache_path.read_text(encoding="utf-8"))
        except Exception:
            cache = {}

    todo = list(rows)
    if companies:
        want = {norm(c).lower() for c in companies if norm(c)}
        todo = [
            r
            for r in todo
            if norm(r.get("company") or r.get("brand") or "").lower() in want
        ]
        if not todo:
            print(f"  WARNING: no rows matched --companies {companies}")
    if skip_apollo_filled:
        todo = [r for r in todo if not is_apollo_protected_row(r, protected_keys)]
    if merge_fields == ROLE_ONLY_FIELDS:
        todo = [
            r
            for r in todo
            if is_apollo_found_contact(
                r,
                cache.get(str(r.get("company") or r.get("brand") or "?").lower()),
            )
        ]
    if fill_empty_only:
        todo = [r for r in todo if needs_apollo_fields(r, merge_fields)]
    if limit is not None:
        todo = todo[: max(0, limit)]

    skipped_protected = (
        sum(1 for r in rows if is_apollo_protected_row(r, protected_keys)) if skip_apollo_filled else 0
    )

    print(f"\n=== APOLLO CONTACT FILL: {slug} ===")
    print(f"  Companies (filtered): {len(rows)}")
    print(f"  Skip Apollo-already-filled: {skipped_protected}")
    print(f"  To process: {len(todo)}")
    print(f"  Mode: {'GAP FILL (empty cells only)' if fill_empty_only else 'REPLACE ALL'}")
    print(f"  API: {'people/match enrich (uses credits)' if enrich_match else 'mixed_people/api_search (free)'}")
    print(f"  Columns: {_fields_label(merge_fields)}")
    print(f"  Request delay: {request_delay}s between Apollo calls")
    print(
        "  Progress: each company calls Apollo search (~3s+). "
        "If rate-limited (429), waits 5–90s between retries — terminal may look idle.",
        flush=True,
    )
    if merge_fields == ROLE_ONLY_FIELDS:
        print("  Role source: same Apollo contact already on row (cache first, then name-matched search)")
    if resume:
        print("  Resume: skip companies already in apollo_contact_cache.json")
    if dry_run:
        print("  [dry-run] No writes.")

    if not todo:
        print("  Nothing to do.")
        return {"processed": 0, "filled": 0}

    filled = 0
    skipped = 0
    errors = 0
    resumed = 0
    recent: list[dict] = []
    t0 = time.perf_counter()
    started_at = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    def _write_track(*, current: dict | None = None, status: str = "running") -> None:
        contact_filled_total = sum(1 for r in rows if has_contact_person(r))
        now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        payload = {
            "slug": slug,
            "status": status,
            "started_at": started_at,
            "updated_at": now,
            "mode": "gap_fill" if fill_empty_only else "replace_all",
            "total_companies": len(rows),
            "contact_filled_total": contact_filled_total,
            "gaps_remaining": len(rows) - contact_filled_total,
            "queue_size": len(todo),
            "processed_this_run": filled + skipped + errors + resumed,
            "filled_this_run": filled,
            "skipped_this_run": skipped,
            "errors_this_run": errors,
            "resumed_from_cache": resumed,
            "elapsed_sec": round(time.perf_counter() - t0, 1),
            "current": current,
            "recent": recent[-10:],
        }
        if not dry_run:
            progress_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")

    def _save_progress() -> None:
        if dry_run:
            return
        json_path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
        cache_path.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")

    def _is_partial_name(name: str) -> bool:
        return "*" in (name or "")

    def _merge_cache_entry(existing: dict | None, new: dict | None) -> dict:
        """Preserve prior enrich fields when a later call writes slim/skipped entries."""
        prev = dict(existing or {})
        incoming = dict(new or {})
        if not incoming:
            return prev
        if incoming.get("skipped") and prev.get("contact_person"):
            return {**prev, "skipped": True}
        out = {**prev}
        for key, val in incoming.items():
            if val in (None, "", False):
                continue
            if key == "contact_person":
                old = norm(str(out.get("contact_person") or ""))
                new_name = norm(str(val))
                if not old or (_is_partial_name(old) and not _is_partial_name(new_name)):
                    out[key] = val
            elif key in ("email", "linkedin", "contact_role", "apollo_person_id", "phone"):
                if not norm(str(out.get(key) or "")):
                    out[key] = val
            else:
                out[key] = val
        return out

    enrich_snapshot_path = market_dir / "apollo_enrich_live_snapshot.json"
    enrich_snapshot: dict[str, dict] = {}
    if enrich_snapshot_path.is_file():
        try:
            enrich_snapshot = json.loads(enrich_snapshot_path.read_text(encoding="utf-8"))
        except Exception:
            enrich_snapshot = {}

    def _save_enrich_snapshot(company_key: str, payload: dict) -> None:
        if dry_run or not enrich_match or merge_fields != CONTACT_BUNDLE_FIELDS:
            return
        if not payload.get("contact_person"):
            return
        entry = {
            k: payload[k]
            for k in ("contact_person", "contact_role", "email", "linkedin", "apollo_person_id", "phone")
            if payload.get(k)
        }
        if entry:
            enrich_snapshot[company_key] = entry
            enrich_snapshot_path.write_text(
                json.dumps(enrich_snapshot, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )

    _write_track(current={"phase": "starting"})

    async with ApolloContactClient(request_delay=request_delay) as apollo:
        for idx, row in enumerate(todo, 1):
            name = str(row.get("company") or row.get("brand") or "?")
            cache_key = name.lower()
            if resume:
                cached = cache.get(cache_key) or {}
                if cached.get("error"):
                    pass
                elif merge_fields == ROLE_ONLY_FIELDS:
                    row_name = norm((row.get("_crawl") or {}).get("contact_person"))
                    cached_name = str(cached.get("contact_person") or "").strip()
                    cached_role = str(cached.get("contact_role") or "").strip()
                    if (
                        cached_role
                        and cached_name
                        and contact_names_match(row_name, cached_name)
                        and not needs_apollo_fields(row, merge_fields)
                    ):
                        print(
                            f"  [{idx}/{len(todo)}] {name} -> {cached_role} (cached role, skip)",
                            flush=True,
                        )
                        resumed += 1
                        continue
                elif cached.get("contact_person") and not cached.get("error"):
                    # Name may already be on the row — still merge if Role/other fields empty
                    if fill_empty_only and not needs_apollo_fields(row, merge_fields):
                        print(
                            f"  [{idx}/{len(todo)}] {name} -> "
                            f"{(row.get('_crawl') or {}).get('contact_person')} (keep existing)",
                            flush=True,
                        )
                        continue
                    if not dry_run:
                        merge_apollo_contact(
                            row,
                            cached,
                            source="apollo",
                            fill_empty_only=fill_empty_only,
                            fields=merge_fields,
                        )
                    # If Role still missing after cache merge, fall through to live Apollo search
                    if fill_empty_only and needs_apollo_fields(row, merge_fields):
                        pass  # continue to API below
                    else:
                        filled += 1
                        resumed += 1
                        role_note = ""
                        if cached.get("contact_role"):
                            role_note = f" / {cached.get('contact_role')}"
                        print(
                            f"  [{idx}/{len(todo)}] {name} -> "
                            f"{cached.get('contact_person')}{role_note} (cached)",
                            flush=True,
                        )
                        continue

            print(f"  [{idx}/{len(todo)}] {name} ...", flush=True)
            t_company = time.perf_counter()

            if merge_fields == ROLE_ONLY_FIELDS:
                cached = cache.get(cache_key) or {}
                cached_name = str(cached.get("contact_person") or "").strip()
                cached_role = str(cached.get("contact_role") or "").strip()
                row_name = norm((row.get("_crawl") or {}).get("contact_person"))
                if (
                    cached_role
                    and cached_name
                    and contact_names_match(row_name, cached_name)
                ):
                    result = {
                        "contact_person": row_name,
                        "contact_role": cached_role,
                        "apollo_from_cache": "true",
                    }
                    print(f"    using cached role for {row_name!r}", flush=True)
                else:
                    dom = company_domain(row)
                    print(f"    domain={dom or '(none)'} — match role to {row_name!r}", flush=True)
                    _write_track(
                        current={
                            "index": idx,
                            "total": len(todo),
                            "company": name,
                            "domain": dom,
                            "phase": "apollo_role_match",
                        }
                    )
                    try:
                        result = await apollo.find_role_for_existing_contact(row)
                    except Exception as exc:
                        errors += 1
                        cache[cache_key] = _merge_cache_entry(cache.get(cache_key), {"error": str(exc)})
                        print(f"  [{idx}/{len(todo)}] {name} ERROR: {exc}", flush=True)
                        _write_track(current=None)
                        _save_progress()
                        continue
            else:
                dom = company_domain(row)
                print(f"    domain={dom or '(none)'} — starting Apollo search", flush=True)
                _write_track(
                    current={
                        "index": idx,
                        "total": len(todo),
                        "company": name,
                        "domain": dom,
                        "phase": "apollo_search",
                    }
                )
                try:
                    result = await apollo.find_contact(row, enrich_match=enrich_match)
                except Exception as exc:
                    errors += 1
                    cache[cache_key] = _merge_cache_entry(cache.get(cache_key), {"error": str(exc)})
                    recent.append({"company": name, "status": "error", "detail": str(exc)[:120]})
                    print(f"  [{idx}/{len(todo)}] {name} ERROR: {exc}", flush=True)
                    _write_track(current=None)
                    _save_progress()
                    continue

            if not result or result.get("error"):
                skipped += 1
                cache[cache_key] = _merge_cache_entry(cache.get(cache_key), result or {"skipped": True})
                if result and result.get("error"):
                    errors += 1
                    err = str(result["error"])
                    recent.append({"company": name, "status": "error", "detail": err[:120]})
                    if "429" in err and merge_fields != ROLE_ONLY_FIELDS:
                        print(
                            f"  [{idx}/{len(todo)}] {name} -> rate limited; "
                            f"waiting 60s before next company...",
                            flush=True,
                        )
                        await asyncio.sleep(60)
                    print(f"  [{idx}/{len(todo)}] {name} -> Apollo API error: {err[:120]}", flush=True)
                elif merge_fields == ROLE_ONLY_FIELDS:
                    recent.append({"company": name, "status": "no_role_match"})
                    print(f"  [{idx}/{len(todo)}] {name} -> no role for existing contact", flush=True)
                else:
                    recent.append({"company": name, "status": "no_match"})
                    print(f"  [{idx}/{len(todo)}] {name} -> no Apollo match", flush=True)
                _write_track(current=None)
                _save_progress()
                continue

            if not dry_run:
                merge_apollo_contact(
                    row,
                    result,
                    source="apollo",
                    fill_empty_only=fill_empty_only,
                    fields=merge_fields,
                )
            cache[cache_key] = _merge_cache_entry(cache.get(cache_key), result)
            _save_enrich_snapshot(cache_key, result)
            filled += 1
            obfuscated = result.get("apollo_name_obfuscated") == "true"
            ob_note = " [partial name]" if obfuscated else ""
            elapsed = time.perf_counter() - t_company
            extras = []
            if result.get("email"):
                extras.append("email")
            if result.get("linkedin"):
                extras.append("linkedin")
            extra_note = f" +{','.join(extras)}" if extras else ""
            role = str(result.get("contact_role") or "").strip()
            if merge_fields == ROLE_ONLY_FIELDS and role:
                print(
                    f"  [{idx}/{len(todo)}] {name} -> {role} ({elapsed:.0f}s)",
                    flush=True,
                )
            elif "contact_role" in merge_fields and role:
                print(
                    f"  [{idx}/{len(todo)}] {name} -> "
                    f"{result.get('contact_person')}{ob_note} — {role}{extra_note} ({elapsed:.0f}s)",
                    flush=True,
                )
            else:
                print(
                    f"  [{idx}/{len(todo)}] {name} -> {result.get('contact_person')}{ob_note}{extra_note} ({elapsed:.0f}s)",
                    flush=True,
                )
            recent.append(
                {
                    "company": name,
                    "status": "filled",
                    "contact_person": result.get("contact_person"),
                    "contact_role": result.get("contact_role"),
                    "email": result.get("email"),
                    "linkedin": bool(result.get("linkedin")),
                    "partial": obfuscated,
                    "elapsed_sec": round(elapsed, 1),
                }
            )
            _write_track(current=None)
            _save_progress()

    if apollo.credits_exhausted:
        print(
            "  NOTE: Apollo enrich credits exhausted — kept search names "
            "(some may be partial, e.g. Faisal Ma***k). Upgrade plan for full names/email.",
            flush=True,
        )

    stats = {
        "slug": slug,
        "total_filtered": len(rows),
        "processed": len(todo),
        "filled": filled,
        "skipped": skipped,
        "errors": errors,
        "resumed_from_cache": resumed,
        "replace_all": replace_all,
        "fill_empty_only": fill_empty_only,
        "merge_fields": sorted(merge_fields),
        "skip_apollo_filled": skip_apollo_filled,
        "skipped_protected": skipped_protected,
        "enrich_match": enrich_match,
        "request_delay_sec": request_delay,
        "apollo_credits_exhausted": apollo.credits_exhausted,
        "elapsed_sec": round(time.perf_counter() - t0, 1),
        "progress_file": str(progress_path),
    }

    _write_track(current=None, status="completed")

    if dry_run:
        print(json.dumps(stats, indent=2))
        return stats

    data["apollo_contact_fill"] = {
        **stats,
        "completed_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
    }
    json_path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    cache_path.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")
    filtered = [r for r in rows if has_founded(r) and has_headquarters(r)]
    neon_path = _sync_neon_filtered(market_dir, data, filtered)
    export_stats = export_one(market_dir, suffix="founded_and_hq_v2")
    stats["excel"] = export_stats.get("file")
    stats["neon_json"] = str(neon_path)

    if push_neon:
        try:
            sys.path.insert(0, str(ROOT / "scripts"))
            from push_export_to_neon import push_export

            # Full replace of 80 companies in Neon (not merge-only).
            neon_stats = push_export(neon_path, merge_companies=False)
            stats["neon_push"] = neon_stats
            print(f"  Neon push: {neon_stats.get('vendor_intel_companies')} companies (replaced)", flush=True)
        except Exception as exc:
            stats["neon_push_error"] = str(exc)
            print(f"  Neon push failed: {exc}", flush=True)

    print(json.dumps(stats, indent=2))
    return stats


def show_status(slug: str = "cyber_physical_systems_security") -> None:
    """Print live Apollo contact fill progress (run in a second terminal while fill runs)."""
    market_dir = ROOT / "output" / slug
    progress_path = market_dir / "apollo_contact_progress.json"
    cache_path = market_dir / "apollo_contact_cache.json"

    _, data, rows = _load_rows(slug, founded_hq_only=True)
    from vendor_intel.enrichment.gap_fill.gaps import has_contact_person

    filled_total = sum(1 for r in rows if has_contact_person(r))
    gaps = [r.get("company") for r in rows if not has_contact_person(r)]

    print(f"\n=== APOLLO STATUS: {slug} ===")
    print(f"  Contact filled: {filled_total}/{len(rows)}  |  Gaps: {len(gaps)}")

    if progress_path.is_file():
        prog = json.loads(progress_path.read_text(encoding="utf-8"))
        print(f"  Run status: {prog.get('status', '?')}")
        print(f"  Last update: {prog.get('updated_at', '?')}")
        print(
            f"  This run: processed={prog.get('processed_this_run', 0)} "
            f"filled={prog.get('filled_this_run', 0)} "
            f"skipped={prog.get('skipped_this_run', 0)} "
            f"errors={prog.get('errors_this_run', 0)} "
            f"elapsed={prog.get('elapsed_sec', 0)}s"
        )
        cur = prog.get("current")
        if cur:
            print(
                f"  NOW: [{cur.get('index')}/{cur.get('total')}] {cur.get('company')} "
                f"({cur.get('domain')}) — {cur.get('phase')}"
            )
        recent = prog.get("recent") or []
        if recent:
            print("  Recent:")
            for item in recent[-5:]:
                st = item.get("status")
                if st == "filled":
                    print(f"    + {item.get('company')} -> {item.get('contact_person')}")
                elif st == "error":
                    print(f"    ! {item.get('company')} ERROR: {item.get('detail', '')[:80]}")
                else:
                    print(f"    - {item.get('company')} -> no match")
    else:
        print("  No progress file yet — start apollo_contact_fill.py first.")

    if cache_path.is_file():
        cache = json.loads(cache_path.read_text(encoding="utf-8"))
        ok = sum(1 for v in cache.values() if isinstance(v, dict) and v.get("contact_person"))
        print(f"  Cache: {len(cache)} entries ({ok} with names)")

    if gaps:
        print(f"  Next gaps: {', '.join(str(g) for g in gaps[:8])}" + (" ..." if len(gaps) > 8 else ""))


def main() -> None:
    parser = argparse.ArgumentParser(description="Apollo contact replace for market export")
    parser.add_argument("slug", nargs="?", default="cyber_physical_systems_security")
    parser.add_argument(
        "--replace-all",
        action="store_true",
        help="Overwrite existing contact_person values (default: gap-fill empty rows only)",
    )
    parser.add_argument("--limit", type=int, default=None, help="Process only N companies")
    parser.add_argument("--concurrent", type=int, default=1, help="Parallel companies (default 1)")
    parser.add_argument(
        "--enrich",
        action="store_true",
        help="Use people/match for full name + email + LinkedIn (uses Apollo credits)",
    )
    parser.add_argument(
        "--contact-bundle",
        action="store_true",
        help="Fill contact_person + email + linkedin only (recommended with --enrich)",
    )
    parser.add_argument(
        "--role-only",
        action="store_true",
        help="Fill contact_role (job title) only — does not change name/email/linkedin",
    )
    parser.add_argument(
        "--with-role",
        action="store_true",
        help="Fill contact_person + contact_role (name and Role column) — no email/linkedin",
    )
    parser.add_argument(
        "--all-fields",
        action="store_true",
        help="Fill contact_person + contact_role + email + linkedin + phone",
    )
    parser.add_argument(
        "--skip-apollo-filled",
        action="store_true",
        help="Skip rows from first Apollo enrich (default: on when using --enrich)",
    )
    parser.add_argument(
        "--include-apollo-filled",
        action="store_true",
        help="Process all rows including first Apollo enrich batch (overrides --skip-apollo-filled)",
    )
    parser.add_argument(
        "--search-only",
        action="store_true",
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=2.0,
        help="Seconds between Apollo API calls (default 2; increase if you see 429 errors)",
    )
    parser.add_argument(
        "--resume",
        action="store_true",
        help="Skip companies already present in apollo_contact_cache.json",
    )
    parser.add_argument(
        "--status",
        action="store_true",
        help="Show progress only (read apollo_contact_progress.json)",
    )
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--no-neon", action="store_true")
    parser.add_argument(
        "--market-dir",
        type=str,
        default="",
        help="Market export folder (overrides output/<slug> lookup)",
    )
    parser.add_argument(
        "--companies",
        type=str,
        default="",
        help="Comma-separated company names to process only (e.g. 'BeyondTrust,Secomea')",
    )
    parser.add_argument(
        "--all-companies",
        action="store_true",
        help="Include all relevant companies (default: Founded+HQ only)",
    )
    args = parser.parse_args()

    from vendor_intel.placeholders.load_keys import apply_env_overrides

    apply_env_overrides()

    if args.status:
        show_status(args.slug)
        return

    if args.role_only:
        merge_fields = ROLE_ONLY_FIELDS
    elif args.with_role:
        merge_fields = frozenset({"contact_person", "contact_role"})
    elif args.all_fields:
        merge_fields = frozenset({"contact_person", "contact_role", "email", "linkedin", "phone"})
    elif args.contact_bundle or args.enrich:
        merge_fields = CONTACT_BUNDLE_FIELDS
    else:
        merge_fields = frozenset({"contact_person"})

    if args.include_apollo_filled:
        skip_apollo_filled = False
    elif args.skip_apollo_filled:
        skip_apollo_filled = True
    else:
        skip_apollo_filled = bool(args.enrich or args.role_only)

    company_filter = [c.strip() for c in args.companies.split(",") if c.strip()] or None
    market_dir_arg = Path(args.market_dir).resolve() if args.market_dir else None

    asyncio.run(
        run(
            args.slug,
            founded_hq_only=not args.all_companies,
            replace_all=args.replace_all,
            limit=args.limit,
            concurrent=args.concurrent,
            enrich_match=args.enrich and not args.search_only,
            merge_fields=merge_fields,
            dry_run=args.dry_run,
            push_neon=not args.no_neon,
            request_delay=max(0.5, args.delay),
            resume=args.resume,
            skip_apollo_filled=skip_apollo_filled,
            companies=company_filter,
            market_dir=market_dir_arg,
        )
    )


if __name__ == "__main__":
    main()
