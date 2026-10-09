#!/usr/bin/env python3
"""Re-fill the export columns for companies already accepted in a checkpoint.

Why this exists: the Kompass stage only writes its Excel once every
candidate has been screened, so a long run that is stopped part-way leaves
its accepted distributors sitting in the checkpoint JSON with no output
file. Separately, several column extractors (Headquarters, Contact Person,
Role) were added after some companies had already been screened, so those
rows have gaps that a resumed run will not revisit.

This script reads the accepted companies straight from the checkpoint,
re-runs the field fill for whatever is still missing, writes the enriched
rows back to the checkpoint, and exports an Excel immediately.

Usage:
    .\\.venv\\Scripts\\python.exe refill_accepted.py --query "Wearable Medical Devices"
    .\\.venv\\Scripts\\python.exe refill_accepted.py --query "..." --limit 50
    .\\.venv\\Scripts\\python.exe refill_accepted.py --query "..." --export-only
"""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT / "src") not in sys.path:
    sys.path.insert(0, str(ROOT / "src"))


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--query", required=True, help='e.g. "Wearable Medical Devices"')
    ap.add_argument("--country", default="global")
    ap.add_argument("--out-dir", default=None)
    ap.add_argument("--final", default=None)
    ap.add_argument("--limit", type=int, default=None, help="only process the first N")
    ap.add_argument(
        "--export-only",
        action="store_true",
        help="skip re-filling; just write the Excel from what the checkpoint holds",
    )
    ap.add_argument(
        "--fast-geo",
        action="store_true",
        help="ask each geography column once instead of with every phrasing "
             "(faster, but usually returns fewer countries)",
    )
    ap.add_argument(
        "--deep",
        action="store_true",
        help="after the Google AI pass, run the full gap-fill engine "
             "(company website crawl, LinkedIn, Wikidata, Wikipedia, "
             "Crunchbase) for whatever is still empty",
    )
    ap.add_argument(
        "--deep-only",
        action="store_true",
        help="skip the Google AI pass and run ONLY the gap-fill engine",
    )
    ap.add_argument(
        "--restart",
        action="store_true",
        help="ignore refill progress and re-ask every company from the start",
    )
    ap.add_argument(
        "--all-columns",
        action="store_true",
        help="re-ask every column, not only the ones still empty",
    )
    args = ap.parse_args()

    import run_cursor_expand as rce

    rce._load_env()

    # Ask the geography columns with EVERY phrasing and union the answers.
    # Google AI Overview is non-deterministic — measured, the same query
    # returned 0-7 countries across six runs, union 7 — so a single ask
    # often yields "Italy" where the company actually sells in eight
    # countries. The main run keeps this off because the extra calls triple
    # its runtime across 9,460 companies; for a few hundred it is cheap and
    # it is the difference between one country and a real list.
    if not args.fast_geo:
        os.environ["GOOGLE_AI_GEO_DEEP"] = "true"

    from vendor_intel.discovery.chatgpt_company_screen import (
        ScreenedCompany,
        fill_fields_via_google_ai,
    )
    from vendor_intel.discovery.kompass_distributors import (
        _CHATGPT_FILL_COLUMNS,
        VerifiedDistributor,
        _verified_distributor_from_dict,
        _verified_distributor_to_dict,
        to_slim_row,
    )
    from vendor_intel.pipeline.openai_expand_checkpoint import ExpandCheckpoint
    from vendor_intel.pipeline.web_expand import resolve_final_path, write_final_xlsx

    final_path = resolve_final_path(
        args.query, args.country, args.out_dir and Path(args.out_dir), args.final
    )
    ckpt = ExpandCheckpoint.load(
        ExpandCheckpoint.path_for(final_path.parent, "kompass")
    )
    if ckpt is None:
        print(f"No Kompass checkpoint found next to {final_path}")
        return 1

    accepted_raw = ckpt.data("kompass_accepted", []) or []
    if not accepted_raw:
        print("Checkpoint has no accepted companies yet — nothing to do.")
        return 1

    accepted = [_verified_distributor_from_dict(d) for d in accepted_raw]
    print(f"checkpoint holds {len(accepted)} accepted distributors")

    if not args.export_only:
        subset = accepted[: args.limit] if args.limit else accepted

        # Resume: skip companies a previous run already refilled. --restart
        # ignores that record and does everything again.
        done_urls: set[str] = set(
            [] if args.restart else (ckpt.data("kompass_refilled_urls", []) or [])
        )
        if done_urls:
            remaining = sum(1 for d in subset if d.kompass_profile_url not in done_urls)
            print(f"{len(done_urls)} already refilled — resuming with {remaining}")

        def _save(urls: set[str]) -> None:
            """Persist both the enriched rows and the resume marker together,
            so a stopped run keeps its work instead of discarding it."""
            ckpt.bump(
                "kompass_3_verify",
                "refilling_accepted",
                kompass_accepted=[_verified_distributor_to_dict(x) for x in accepted],
                kompass_refilled_urls=sorted(u for u in urls if u),
                note=f"refill progress: {len(urls)} companies",
            )

        if not args.deep_only:
            print(f"re-filling {len(subset)} companies...\n")
            _refill(
                subset,
                args,
                fill_fields_via_google_ai,
                ScreenedCompany,
                _CHATGPT_FILL_COLUMNS,
                save=_save,
                done_urls=done_urls,
            )

        if args.deep or args.deep_only:
            # Tracked under its OWN key: a company can be finished for the
            # Google-AI pass but not yet crawled, so sharing one marker
            # would make stage 3 skip companies it never actually visited.
            crawled: set[str] = set(
                [] if args.restart else (ckpt.data("kompass_crawled_urls", []) or [])
            )

            def _save_crawl(urls: set[str]) -> None:
                ckpt.bump(
                    "kompass_3_verify",
                    "deep_filling_accepted",
                    kompass_accepted=[_verified_distributor_to_dict(x) for x in accepted],
                    kompass_crawled_urls=sorted(u for u in urls if u),
                    note=f"website crawl progress: {len(urls)} companies",
                )

            remaining = sum(1 for d in subset if d.kompass_profile_url not in crawled)
            if crawled:
                print(f"{len(crawled)} already crawled — resuming with {remaining}")
            print(f"\ndeep gap-fill for {remaining} companies...\n")
            _deep_fill(subset, save=_save_crawl, done_urls=crawled)

        ckpt.bump(
            "kompass_3_verify",
            "refilled_accepted",
            kompass_accepted=[_verified_distributor_to_dict(d) for d in accepted],
            note=f"re-filled columns for {len(subset)} accepted companies",
        )
        print("\ncheckpoint updated")

    rows = [to_slim_row(d) for d in accepted]
    write_final_xlsx(
        final_path,
        rows,
        sheet_name=args.query[:31],
        audit={"query": args.query, "backend": "kompass", "source": "refill_accepted"},
    )
    print(f"\nwrote {len(rows)} rows -> {final_path}")
    _report(rows)
    return 0


_REFILL_SAVE_EVERY = 25  # persist progress this often, not only at the end


def _refill(
    subset,
    args,
    fill_fields_via_google_ai,
    ScreenedCompany,
    all_columns,
    *,
    save=None,
    done_urls=None,
) -> None:
    """Ask Google AI Overview for whatever each company is still missing.

    `done_urls` is the set of companies a previous run already refilled; it
    is mutated as work completes and persisted by `save`. Without it a
    stopped run restarted from company 1 and re-asked Google about hundreds
    of companies it had already paid for — and because the checkpoint was
    only written after ALL companies finished, stopping part-way threw the
    completed work away entirely.
    """
    # Map the export column back onto the dataclass field that holds it, so
    # a filled value lands where to_slim_row() will actually look for it.
    DIRECT = {
        "Website": "website",
        "Founded": "founded_year",
        "Headquarters": "headquarters",
        "Employees": "employees",
        "Contact Person": "contact_person",
        "Email": "email",
        "Office No.": "phone",
        "Key Brands Represented": "brands",
    }

    done_urls = done_urls if done_urls is not None else set()
    processed_since_save = 0

    for i, d in enumerate(subset, 1):
        if d.kompass_profile_url in done_urls:
            continue

        if args.all_columns:
            gaps = list(all_columns)
        else:
            gaps = [c for c in all_columns if _is_blank(d, c, DIRECT)]
        if not gaps:
            print(f"[{i}/{len(subset)}] {d.company_name[:44]} — already complete")
            done_urls.add(d.kompass_profile_url)
            continue

        sc = ScreenedCompany(
            name=d.company_name, profile_url=d.kompass_profile_url, country_hint=d.country
        )
        fill_fields_via_google_ai([sc], columns=gaps, market=d.market)

        for col, value in sc.fields.items():
            value = str(value or "").strip()
            if not value:
                continue
            if col in DIRECT:
                setattr(d, DIRECT[col], value)
            else:
                d.extra_columns = {**(d.extra_columns or {}), col: value}

        print(f"[{i}/{len(subset)}] {d.company_name[:44]} — filled {len(sc.fields)}/{len(gaps)}")

        done_urls.add(d.kompass_profile_url)
        processed_since_save += 1
        if save and processed_since_save >= _REFILL_SAVE_EVERY:
            save(done_urls)
            processed_since_save = 0
            print(f"    ...progress saved ({len(done_urls)} refilled)")

    if save and processed_since_save:
        save(done_urls)


def _deep_fill(subset, *, save=None, done_urls=None) -> None:
    """Run the existing gap-fill engine over the accepted distributors.

    Everything here is orchestration: the extraction itself lives in
    enrichment/gap_fill/, which already crawls the company website's
    /about, /products and /contact pages and consults LinkedIn, Wikidata,
    Wikipedia and Crunchbase. The Kompass stage simply never invoked it,
    which is why nine columns sat near zero despite the capability existing.
    """
    import asyncio

    from vendor_intel.clients.search_router import FreeSearchRouter
    from vendor_intel.config import Settings
    from vendor_intel.enrichment.gap_fill.orchestrator import gap_fill_row
    from vendor_intel.pipeline.kompass_gap_adapter import apply_gap_row, to_gap_row

    # Deliberately NOT the engine's DEFAULT_SOURCES. That list leads with
    # "google_ai", which is given a 900s budget per company because it fires
    # ~22 sequential Overview queries — and the _refill stage above has
    # already asked Google exactly those questions. Including it here made a
    # 3-company smoke test exceed 10 minutes with nothing written.
    #
    # These sources read the company's own website and structured databases,
    # which is precisely what Google could not answer for small private
    # firms, so this is the complement of the Google pass rather than a
    # repeat of it.
    # LinkedIn, Crunchbase and Wikipedia are deliberately excluded. LinkedIn
    # blocks anonymous fetches and reaches it by GUESSING a company slug from
    # the name, which for these firms produces URLs like
    # ".../company/instrumentacin-y-componentes-sainycom-sa" — measured, it
    # was the slowest source and filled 1 value across 5 companies.
    # Crunchbase/Wikipedia cover funded startups and notable firms, not small
    # private European distributors.
    #
    # What remains reads the company's OWN website, which is the one place
    # these facts are reliably published.
    sources = (
        "contact_page",   # /contact -> email, office email, phone
        "leadership",     # /about, /team -> contact person + role
        "backfill",       # DDGS/SearXNG -> website for the 316 rows lacking one
        "wikidata",       # cheap structured lookup; founded/HQ when present
    )
    # Per-source ceiling: without this the engine's 90s default applies to
    # each of the 7 sources, so one dead domain could stall a company for
    # over 10 minutes.
    os.environ.setdefault("GAP_FILL_SOURCE_TIMEOUT_SEC", "45")

    router = FreeSearchRouter(Settings.load())
    total = len(subset)
    # A whole company is capped too: 7 sources x 45s is still 5 minutes in
    # the worst case, and no single row is worth that.
    per_company_budget = 120.0

    tracked = done_urls if done_urls is not None else set()

    async def run() -> None:
        filled_any = 0
        since_save = 0
        for i, d in enumerate(subset, 1):
            # Resume: a website crawl is the slowest stage in the pipeline,
            # so re-doing completed companies after a stop is expensive.
            if d.kompass_profile_url in tracked:
                continue
            row = to_gap_row(d)
            try:
                await asyncio.wait_for(
                    gap_fill_row(
                        row,
                        router=router,
                        owler_client=None,
                        owler_model="",
                        sources=sources,
                        idx=i,
                        total=total,
                    ),
                    timeout=per_company_budget,
                )
            except asyncio.TimeoutError:
                # Keep whatever the sources managed to write before the cap.
                print(f"[{i}/{total}] {d.company_name[:44]} - timed out, keeping partial")
            except Exception as exc:  # one bad company must not end the run
                print(f"[{i}/{total}] {d.company_name[:44]} - gap-fill failed: {exc}")
                continue
            changed = apply_gap_row(d, row)
            if changed:
                filled_any += 1
            label = ", ".join(changed[:4]) if changed else "nothing new"
            print(f"[{i}/{total}] {d.company_name[:44]} - {label}")

            tracked.add(d.kompass_profile_url)
            since_save += 1
            if save and since_save >= _REFILL_SAVE_EVERY:
                save(tracked)
                since_save = 0
                print(f"    ...progress saved ({len(tracked)} crawled)")

        if save and since_save:
            save(tracked)
        print(f"\ndeep gap-fill improved {filled_any}/{total} companies")

    asyncio.run(run())


def _is_blank(d, column: str, direct: dict[str, str]) -> bool:
    """True when a column has no REAL value yet.

    Kompass writes "Not publicly disclosed" as a display placeholder for
    fields it never found, and Core Categories defaults to the market name.
    Treating those as filled is why a first pass only closed ~1 of 7 gaps
    per company — they blocked the re-ask despite carrying no information.
    """
    if column in direct:
        value = getattr(d, direct[column], "")
    else:
        value = (d.extra_columns or {}).get(column, "")
    value = str(value or "").strip()
    if not value or value.lower() == "not publicly disclosed":
        return True

    # The geography columns are supposed to NAME countries. A region-level
    # answer ("Global", "Europe", "Worldwide") is what the old prompt
    # wording produced and is useless in a spreadsheet cell, so treat it as
    # unfilled and ask again with the current wording.
    if column in _GEO_COLUMNS and value.lower() in _VAGUE_GEO:
        return True

    # Core Categories seeded with the bare market name is a placeholder too.
    return column == "Core Categories" and value.strip().lower() == (d.market or "").strip().lower()


_GEO_COLUMNS = ("Continent / Geography", "Operational Presence")
_VAGUE_GEO = {
    "global", "worldwide", "international", "europe", "asia", "africa",
    "north america", "south america", "latin america", "americas", "oceania",
    "emea", "apac", "eu", "middle east", "n/a", "various", "multiple countries",
}


def _report(rows: list[dict]) -> None:
    from vendor_intel.export.landscape_slim import SLIM_COLUMNS

    n = len(rows) or 1
    # Count only REAL values. "Not publicly disclosed" is a display
    # placeholder, so counting it as filled reported columns at 100% that
    # in fact held no information.
    print("\ncolumn fill rate (placeholders not counted):")
    for col in SLIM_COLUMNS:
        filled = sum(
            1
            for r in rows
            if str(r.get(col) or "").strip()
            and str(r.get(col)).strip().lower() != "not publicly disclosed"
        )
        print(f"  {col:<26} {filled:>4}/{len(rows)}  {filled/n*100:5.1f}%")


if __name__ == "__main__":
    raise SystemExit(main())
