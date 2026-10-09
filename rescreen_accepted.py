#!/usr/bin/env python3
"""Re-run the market-relevance screen over already-accepted distributors.

Why this exists: the screen originally asked only "is this company a
distributor?". A company could pass by being a genuine distributor of a
DIFFERENT industry. Fritz Minke GmbH & Co. KG reached the Wearable Medical
Devices export that way — Google's own answer was "No, Fritz Minke is not a
distributor of wearable medical devices... they are a B2B wholesaler for
orthopedic shoemaking", but the rejection sat in one sentence and the word
"wholesaler" in the next, so the sentence-scoped negation check missed it.

The market gate that fixes this only applies to companies screened from now
on. This script applies it retroactively to rows already in the checkpoint.

Nothing is deleted. Rejected companies move to `kompass_rescreen_rejected`
so the decision stays auditable and reversible — a wrong rejection here
would silently drop a real distributor from the deliverable.

Usage:
    .\\.venv\\Scripts\\python.exe rescreen_accepted.py --query "Wearable Medical Devices"
    .\\.venv\\Scripts\\python.exe rescreen_accepted.py --query "..." --limit 25   # sample first
    .\\.venv\\Scripts\\python.exe rescreen_accepted.py --query "..." --apply      # write changes
    .\\.venv\\Scripts\\python.exe rescreen_accepted.py --query "..." --report     # show quarantine
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT / "src") not in sys.path:
    sys.path.insert(0, str(ROOT / "src"))

_SAVE_EVERY = 25  # a stopped run must not lose completed screening work


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--query", required=True, help='e.g. "Wearable Medical Devices"')
    ap.add_argument("--country", default="global")
    ap.add_argument("--out-dir", default=None)
    ap.add_argument("--final", default=None)
    ap.add_argument("--limit", type=int, default=None, help="only screen the first N")
    ap.add_argument(
        "--apply",
        action="store_true",
        help="actually move failing companies out of the accepted list "
             "(without this the run is a dry run and changes nothing)",
    )
    ap.add_argument(
        "--report",
        action="store_true",
        help="just print what a previous --apply quarantined, and exit",
    )
    ap.add_argument(
        "--restore",
        action="store_true",
        help="move everything in the quarantine back into accepted",
    )
    args = ap.parse_args()

    import run_cursor_expand as rce

    rce._load_env()

    from vendor_intel.discovery.kompass_distributors import (
        _verified_distributor_from_dict,
        _verified_distributor_to_dict,
    )
    from vendor_intel.pipeline.openai_expand_checkpoint import ExpandCheckpoint
    from vendor_intel.pipeline.web_expand import resolve_final_path

    final_path = resolve_final_path(
        args.query, args.country, args.out_dir and Path(args.out_dir), args.final
    )
    ckpt = ExpandCheckpoint.load(ExpandCheckpoint.path_for(final_path.parent, "kompass"))
    if ckpt is None:
        print(f"No Kompass checkpoint found next to {final_path}")
        return 1

    quarantine = ckpt.data("kompass_rescreen_rejected", []) or []

    if args.report:
        return _report(quarantine)

    if args.restore:
        if not quarantine:
            print("Quarantine is empty — nothing to restore.")
            return 0
        accepted = (ckpt.data("kompass_accepted", []) or []) + quarantine
        ckpt.bump(
            "kompass_3_verify",
            "rescreen_restored",
            kompass_accepted=accepted,
            kompass_rescreen_rejected=[],
            note=f"restored {len(quarantine)} companies from rescreen quarantine",
        )
        print(f"restored {len(quarantine)} companies -> accepted is now {len(accepted)}")
        return 0

    accepted_raw = ckpt.data("kompass_accepted", []) or []
    if not accepted_raw:
        print("Checkpoint has no accepted companies yet — nothing to do.")
        return 1

    # Resume: skip anything a previous run already screened.
    done = set(ckpt.data("kompass_rescreened_urls", []) or [])
    todo = [d for d in accepted_raw if d.get("kompass_profile_url") not in done]
    if args.limit:
        todo = todo[: args.limit]

    print(f"checkpoint holds {len(accepted_raw)} accepted distributors")
    if done:
        print(f"{len(done)} already re-screened — resuming with {len(todo)}")
    if not todo:
        print("Everything has been re-screened already.")
        return _report(quarantine)
    if not args.apply:
        print("DRY RUN — no changes will be written. Re-run with --apply to act.\n")

    failures = _rescreen(todo, args.query, ckpt, done, accepted_raw, args.apply)

    print(f"\n{'-' * 62}")
    print(f"screened {len(todo)}  |  failed the market gate: {len(failures)}")
    for name, reason in failures[:20]:
        print(f"  REJECT  {name[:44]:<46} {reason[:60]}")
    if len(failures) > 20:
        print(f"  ... and {len(failures) - 20} more")

    if not args.apply:
        print("\nDry run — nothing written. Re-run with --apply to quarantine these.")
        return 0

    print(f"\nquarantined {len(failures)} companies into kompass_rescreen_rejected")
    print("re-export with:")
    print(f'  .\\.venv\\Scripts\\python.exe refill_accepted.py --query "{args.query}" --export-only')
    return 0


def _rescreen(todo, market, ckpt, done, accepted_raw, apply_changes):
    """Ask Google AI Overview about each company and apply the market gate."""
    import asyncio

    from vendor_intel.clients.google_ai_scraper import google_ai_ask, google_ai_ping
    from vendor_intel.discovery.chatgpt_company_screen import (
        _verdict_from_overview,
        google_ai_concurrency,
    )

    failures: list[tuple[str, str]] = []
    rejected_urls: set[str] = set()
    blocked = {"n": 0}
    total = len(todo)
    workers = google_ai_concurrency()

    async def run() -> None:
        if not await google_ai_ping():
            print("Google AI scraper backend is not responding — start it first.")
            return

        sem = asyncio.Semaphore(workers)
        counter = {"n": 0}

        async def one(rec: dict) -> None:
            async with sem:
                name = rec.get("company_name") or ""
                query = (
                    f"Is {name} a distributor (reseller of other brands) "
                    f"or a manufacturer in the {market} market?"
                )
                try:
                    data = await google_ai_ask(query)
                    verdict, reason = _verdict_from_overview(
                        data.get("markdown") or "", market=market
                    )
                except Exception as exc:
                    # An error is NOT evidence against the company — leave it
                    # accepted rather than dropping a real distributor.
                    verdict, reason = "unknown", f"{type(exc).__name__}"

                counter["n"] += 1
                i = counter["n"]

                # Google rate-limits bursts of queries. Once it starts
                # serving CAPTCHA pages every remaining answer is garbage,
                # so stop rather than record hundreds of "unknown" verdicts
                # that would look like real screening work on resume.
                if "captcha" in reason.lower() or "blocked" in reason.lower():
                    blocked["n"] += 1
                    done.discard(rec.get("kompass_profile_url"))
                    print(f"[{i}/{total}] BLOCKED {name[:40]:<41} (Google bot detection)")
                    return

                if verdict == "manufacturer":
                    failures.append((name, reason))
                    rejected_urls.add(rec.get("kompass_profile_url"))
                    # Keep Google's own words on the record so a human can
                    # judge the rejection without re-running the query.
                    rec["rejection_reason"] = reason
                    rec["is_actual_market_distributor"] = False
                    print(f"[{i}/{total}] REJECT {name[:40]:<42} {reason[:50]}")
                else:
                    print(f"[{i}/{total}] keep   {name[:40]:<42} {verdict}")

                done.add(rec.get("kompass_profile_url"))
                if apply_changes and i % _SAVE_EVERY == 0:
                    _save(ckpt, accepted_raw, rejected_urls, done, market)

        await asyncio.gather(*(one(r) for r in todo))

    asyncio.run(run())

    if apply_changes:
        _save(ckpt, accepted_raw, rejected_urls, done, market)
    if blocked["n"]:
        print(
            f"\n!! {blocked['n']} of {len(todo)} queries hit Google bot detection."
            "\n   Those companies were NOT screened and stay accepted; re-run"
            "\n   later to finish them. Lower GOOGLE_AI_SCREEN_CONCURRENCY to"
            "\n   reduce blocking."
        )
    return failures


def _save(ckpt, accepted_raw, rejected_urls, done, market) -> None:
    """Move rejects into quarantine and record progress. Never deletes."""
    keep = [d for d in accepted_raw if d.get("kompass_profile_url") not in rejected_urls]
    moved = [d for d in accepted_raw if d.get("kompass_profile_url") in rejected_urls]
    existing = ckpt.data("kompass_rescreen_rejected", []) or []
    seen = {d.get("kompass_profile_url") for d in existing}
    existing = existing + [d for d in moved if d.get("kompass_profile_url") not in seen]

    ckpt.bump(
        "kompass_3_verify",
        "rescreened_market_relevance",
        kompass_accepted=keep,
        kompass_rescreen_rejected=existing,
        kompass_rescreened_urls=sorted(u for u in done if u),
        note=f"market-relevance re-screen for {market}: {len(existing)} quarantined",
    )


def _report(quarantine) -> int:
    if not quarantine:
        print("Quarantine is empty — no companies have been rejected.")
        return 0
    print(f"{len(quarantine)} companies quarantined by the market-relevance screen:\n")
    for d in quarantine:
        print(f"  {d.get('company_name', '')[:50]:<52} {str(d.get('rejection_reason', ''))[:60]}")
    print("\nRestore them all with --restore if this looks wrong.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
