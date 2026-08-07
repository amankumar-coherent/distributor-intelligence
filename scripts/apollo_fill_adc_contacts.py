#!/usr/bin/env python3
"""Apollo-fill Contact Person / Role / LinkedIn on ADC FINAL Excel.

Keeps Apollo partial/obfuscated names (e.g. \"John S***h\").
Gap-fills only empty / Not publicly disclosed rows — does not overwrite
existing verified full names.

Usage:
  .venv\\Scripts\\python.exe scripts\\apollo_fill_adc_contacts.py
  .venv\\Scripts\\python.exe scripts\\apollo_fill_adc_contacts.py --limit 20
  .venv\\Scripts\\python.exe scripts\\apollo_fill_adc_contacts.py --replace-all
"""
from __future__ import annotations

import argparse
import asyncio
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT / "scripts"))

from dotenv import load_dotenv

load_dotenv(ROOT / ".env")

import fill_adc_contact_role_linkedin as xls  # noqa: E402

OUT = ROOT / "output/web_expand/global_antibody_drug_conjugates_market_global"
XLSX = OUT / "global_antibody_drug_conjugates_market_FINAL.xlsx"
PROGRESS = OUT / "adc_apollo_contact_progress.jsonl"
CACHE = OUT / "adc_apollo_contact_cache.json"
SUMMARY = OUT / "adc_apollo_contact_summary.json"

NPD = {"", "n/a", "na", "none", "-", "—", "not publicly disclosed", "unknown", "tbd", "null"}


def _empty(v: object) -> bool:
    return str(v or "").strip().lower() in NPD


def _needs_apollo(row: dict) -> bool:
    cp = str(row.get("Contact Person") or "").strip()
    if _empty(cp):
        return True
    low = cp.lower()
    # Placeholders left by earlier scrapes
    if any(x in low for x in ("leadership", "joint md", "executive leadership")):
        return True
    return False


def _is_partial(name: str) -> bool:
    return "*" in (name or "")


def _load_cache() -> dict:
    if CACHE.exists():
        try:
            return json.loads(CACHE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def _save_cache(cache: dict) -> None:
    CACHE.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")


def _append_progress(rec: dict) -> None:
    with PROGRESS.open("a", encoding="utf-8") as f:
        f.write(json.dumps(rec, ensure_ascii=False) + "\n")


async def run(*, replace_all: bool, limit: int | None, resume: bool) -> dict:
    from vendor_intel.enrichment.apollo_contacts import ApolloContactClient

    if not (os.getenv("APOLLO_API_KEY") or "").strip():
        raise SystemExit("APOLLO_API_KEY missing in .env")

    # Backup once
    bak = XLSX.with_name(XLSX.stem + "_pre_apollo.xlsx")
    if not bak.exists():
        bak.write_bytes(XLSX.read_bytes())
        print(f"Backup: {bak.name}", flush=True)

    rows = xls.load_rows(XLSX)
    cache = _load_cache()
    done_cos = set()
    if resume and PROGRESS.exists():
        for line in PROGRESS.read_text(encoding="utf-8").splitlines():
            if not line.strip():
                continue
            try:
                rec = json.loads(line)
            except Exception:
                continue
            # Do NOT skip errors/timeouts/429 — those should retry after cooldown
            if rec.get("status") in {"filled", "no_match", "skipped_existing"}:
                done_cos.add(str(rec.get("company") or "").strip().lower())

    todo: list[dict] = []
    for r in rows:
        co = r["Company"]
        key = co.strip().lower()
        if resume and key in done_cos:
            continue
        if not replace_all and not _needs_apollo(r):
            continue
        todo.append(r)

    if limit is not None:
        todo = todo[:limit]

    print(
        f"Apollo ADC contact fill: {len(todo)} todo / {len(rows)} rows "
        f"(replace_all={replace_all}, keep partial names, enrich=OFF)",
        flush=True,
    )

    filled = 0
    partial_n = 0
    no_match = 0
    errors = 0
    skipped = 0

    delay = float(os.getenv("APOLLO_REQUEST_DELAY_SEC") or "3.0")
    # Exceeds Apollo client 429 backoff cap (default 120s × retries)
    per_company_timeout = float(os.getenv("APOLLO_COMPANY_TIMEOUT_SEC") or "600")
    rate_limit_sleep = float(os.getenv("APOLLO_RATE_LIMIT_SLEEP_SEC") or "1800")
    async with ApolloContactClient(verbose=True, request_delay=delay, timeout=45.0) as apollo:
        for i, r in enumerate(todo, 1):
            co = r["Company"]
            website = str(r.get("Website") or "").strip()
            cache_key = co.strip().lower()

            # Use cache if present
            if cache_key in cache and isinstance(cache[cache_key], dict) and cache[cache_key].get("contact_person"):
                hit = cache[cache_key]
                status = "cached"
            else:
                apollo_row = {"company": co, "website": website, "domain": ""}
                print(f"  [{i}/{len(todo)}] {co[:42]} … querying Apollo…", flush=True)
                try:
                    # enrich_match=False → keep free-search partial names
                    hit = await asyncio.wait_for(
                        apollo.find_contact(apollo_row, enrich_match=False),
                        timeout=per_company_timeout,
                    )
                    status = "live"
                except asyncio.TimeoutError:
                    hit = {"error": f"Timeout after {per_company_timeout:.0f}s"}
                    status = "error"
                except Exception as exc:
                    hit = {"error": f"{type(exc).__name__}: {exc}"}
                    status = "error"

                err_s = str((hit or {}).get("error") or "")
                if hit and ("429" in err_s or "rate limit" in err_s.lower() or "Timeout" in err_s):
                    # Likely rate-limit: cool down then retry this company once
                    print(
                        f"  !! Apollo cool-down {rate_limit_sleep:.0f}s then retry {co[:36]}…",
                        flush=True,
                    )
                    await asyncio.sleep(rate_limit_sleep)
                    try:
                        hit = await asyncio.wait_for(
                            apollo.find_contact(apollo_row, enrich_match=False),
                            timeout=per_company_timeout,
                        )
                        status = "live"
                    except Exception as exc:
                        hit = {"error": f"{type(exc).__name__}: {exc}"}
                        status = "error"

            if not hit or hit.get("error"):
                errors += 1 if hit and hit.get("error") else 0
                no_match += 0 if hit and hit.get("error") else 1
                if hit and hit.get("error"):
                    print(f"  [{i}/{len(todo)}] {co[:42]} -> ERROR {str(hit.get('error'))[:100]}", flush=True)
                    # Do not mark done — resume can retry later
                else:
                    print(f"  [{i}/{len(todo)}] {co[:42]} -> no Apollo match", flush=True)
                    _append_progress({"company": co, "status": "no_match"})
                    cache[cache_key] = {"skipped": True}
                if i % 10 == 0:
                    _save_cache(cache)
                    rows_sorted = sorted(rows, key=lambda x: x["Company"].lower())
                    xls.write_xlsx(XLSX, rows_sorted)
                continue

            person = str(hit.get("contact_person") or "").strip()
            role = str(hit.get("contact_role") or "").strip()
            li = str(hit.get("linkedin") or "").strip()
            if not person:
                no_match += 1
                _append_progress({"company": co, "status": "no_match"})
                print(f"  [{i}/{len(todo)}] {co[:42]} -> no name", flush=True)
                continue

            # Apply (gap-fill unless replace_all)
            if replace_all or _needs_apollo(r) or _empty(r.get("Contact Person")):
                r["Contact Person"] = person
                if role:
                    r["Role"] = role
                elif _empty(r.get("Role")):
                    r["Role"] = "Not publicly disclosed"
                if li and "/in/" in li:
                    r["LinkedIn"] = li
                elif _empty(r.get("LinkedIn")) or "/company/" in str(r.get("LinkedIn") or ""):
                    # Keep NPD for LinkedIn when Apollo has no person URL
                    if _empty(r.get("LinkedIn")):
                        r["LinkedIn"] = "Not publicly disclosed"
                filled += 1
                if _is_partial(person) or hit.get("apollo_name_obfuscated") == "true":
                    partial_n += 1
                note = " [partial]" if ("*" in person) else ""
                print(
                    f"  [{i}/{len(todo)}] {co[:42]} -> {person}{note}"
                    + (f" | {role}" if role else "")
                    + (f" | LI" if li and '/in/' in li else ""),
                    flush=True,
                )
                _append_progress(
                    {
                        "company": co,
                        "status": "filled",
                        "contact_person": person,
                        "role": role,
                        "linkedin": li,
                        "partial": "*" in person,
                        "source": status,
                    }
                )
            else:
                skipped += 1
                _append_progress({"company": co, "status": "skipped_existing"})

            cache[cache_key] = {
                k: v
                for k, v in hit.items()
                if k in {
                    "contact_person",
                    "contact_role",
                    "linkedin",
                    "email",
                    "phone",
                    "apollo_name_obfuscated",
                    "apollo_person_id",
                }
            }

            if i % 10 == 0:
                _save_cache(cache)
                rows_sorted = sorted(rows, key=lambda x: x["Company"].lower())
                xls.write_xlsx(XLSX, rows_sorted)
                print(f"  … checkpoint saved ({filled} filled so far)", flush=True)

    _save_cache(cache)
    rows.sort(key=lambda x: x["Company"].lower())
    xls.write_xlsx(XLSX, rows)

    named = sum(1 for r in rows if not _empty(r.get("Contact Person")))
    partials = sum(1 for r in rows if "*" in str(r.get("Contact Person") or ""))
    roles = sum(1 for r in rows if not _empty(r.get("Role")))
    lis = sum(1 for r in rows if "/in/" in str(r.get("LinkedIn") or ""))

    summary = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "total_rows": len(rows),
        "todo": len(todo),
        "filled_this_run": filled,
        "partial_this_run": partial_n,
        "no_match": no_match,
        "errors": errors,
        "skipped_existing": skipped,
        "contact_person_filled_total": named,
        "partial_names_total": partials,
        "role_filled_total": roles,
        "person_linkedin_total": lis,
        "note": "Apollo partial/obfuscated names kept as-is (enrich_match=False).",
    }
    SUMMARY.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps(summary, indent=2), flush=True)
    return summary


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=None)
    ap.add_argument("--replace-all", action="store_true")
    ap.add_argument("--no-resume", action="store_true")
    args = ap.parse_args()
    asyncio.run(
        run(
            replace_all=args.replace_all,
            limit=args.limit,
            resume=not args.no_resume,
        )
    )


if __name__ == "__main__":
    main()
