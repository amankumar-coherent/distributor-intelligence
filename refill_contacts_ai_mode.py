#!/usr/bin/env python3
"""Re-fill the six contact columns from Google AI Mode — real data only.

Re-asks AI Mode, 15 companies per query, for each company's REAL published
contact details with explicit permission to answer "unknown". Anything it
declines (or that looks constructed) is left BLANK rather than guessed.

The six columns: Contact Person, Role, Email, LinkedIn, Office No.,
Office Email.

READ THIS BEFORE TRUSTING THE OUTPUT. AI Mode is the same source that
produced the original guessed values, so this is not independent
verification — it can repeat a plausible invention. What makes it worth
running is that the prompt now REWARDS "unknown" instead of forbidding it,
so fields the model never really knew tend to come back empty. For
independent confirmation against the companies' own websites, use
verify_contact_fields.py, which fetches the real pages over HTTP.

Usage:
    # dry run: show what WOULD change, write nothing (do this first)
    python refill_contacts_ai_mode.py --market "..." --limit 15 --dry-run

    # apply to a sample, then to everything (state is backed up first)
    python refill_contacts_ai_mode.py --market "..." --limit 15
    python refill_contacts_ai_mode.py --market "..."

    # keep generic info@ addresses that AI Mode reports (default: reject)
    python refill_contacts_ai_mode.py --market "..." --allow-generic-email

Progress prints per batch. Interrupting is safe: the state file is only
rewritten at the end, and always after a timestamped backup.
"""
from __future__ import annotations

import argparse
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

from vendor_intel.placeholders.load_keys import apply_env_overrides  # noqa: E402

apply_env_overrides()

from vendor_intel.pipeline.market_research import contact_lookup, schema  # noqa: E402
from vendor_intel.pipeline.market_research.export_xlsx import export_xlsx  # noqa: E402
from vendor_intel.pipeline.market_research.state import (  # noqa: E402
    STATE_DIR,
    MarketResearchState,
    market_slug,
)

_FIELD_TO_COLUMN = {
    "contact_person": "Contact Person",
    "role": "Role",
    "email": "Email",
    "linkedin": "LinkedIn",
    "office_no": "Office No.",
    "office_email": "Office Email",
}


def _log(msg: str) -> None:
    print(f"  [refill] {msg}", flush=True)


def main() -> None:
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    ap.add_argument("--market", required=True)
    ap.add_argument("--limit", type=int, default=0, help="only the first N companies (0 = all)")
    ap.add_argument(
        "--batch-size", type=int, default=contact_lookup.BATCH_SIZE,
        help=f"companies per AI Mode query (default {contact_lookup.BATCH_SIZE})",
    )
    ap.add_argument(
        "--fields", default="",
        help="comma-separated columns to refill (default: all six). Only these "
             "are written; every other contact column is left exactly as-is.",
    )
    ap.add_argument(
        "--only-empty", action="store_true",
        help="skip companies that already have a value in every requested "
             "field, and never overwrite a field that is already filled",
    )
    ap.add_argument("--dry-run", action="store_true", help="report changes, write nothing")
    ap.add_argument(
        "--allow-generic-email", action="store_true",
        help="keep info@/sales@ at the company's own domain (default: reject "
             "as indistinguishable from the original fabrication)",
    )
    args = ap.parse_args()

    # Which columns this run is allowed to touch. Restricting them matters:
    # a field left out is not merely skipped, it is never written, so an
    # already-good value in another column cannot be blanked by this run.
    if args.fields.strip():
        wanted = [f.strip() for f in args.fields.split(",") if f.strip()]
        by_col = {c.casefold(): c for c in _FIELD_TO_COLUMN.values()}
        columns: list[str] = []
        for w in wanted:
            col = by_col.get(w.casefold())
            if col is None:
                print(f"Unknown field {w!r}. Valid: {', '.join(_FIELD_TO_COLUMN.values())}")
                return
            columns.append(col)
    else:
        columns = list(_FIELD_TO_COLUMN.values())
    args._columns = columns

    state = MarketResearchState.load(args.market)
    rows = list(state.companies)

    total = len(rows)
    if args.only_empty:
        # Ask only about companies that are actually missing something.
        # Without this every company is re-queried, which wastes queries
        # (each one risks a CAPTCHA) and re-asks about data already known.
        rows = [
            r for r in rows
            if any(not str(r.get(c, "") or "").strip() for c in columns)
        ]
        print(f"--only-empty: {len(rows)} of {total} companies have a gap in "
              f"{len(columns)} requested field(s)")

    if args.limit:
        rows = rows[: args.limit]
    if not rows:
        print("Nothing to do — every company already has all requested fields.")
        return

    batch_size = max(1, min(20, args.batch_size))
    batches = [rows[i : i + batch_size] for i in range(0, len(rows), batch_size)]

    print()
    print("=" * 66)
    print("Re-asking Google AI Mode for REAL contact data only.")
    print(f"{len(rows)} companies | {len(batches)} queries of up to {batch_size}")
    print(f"fields: {', '.join(columns)}")
    if args.only_empty:
        print("only-empty: existing values are preserved, never overwritten.")
    print("Unknown values will be left BLANK, never guessed.")
    if args.dry_run:
        print("DRY RUN — nothing will be written.")
    print("=" * 66)
    print()

    results: dict[str, dict[str, str]] = {}
    for n, batch in enumerate(batches, 1):
        _log(f"batch {n}/{len(batches)}: asking about {len(batch)} companies...")
        found = contact_lookup.lookup_contacts(
            batch, strict=not args.allow_generic_email
        )
        results.update(found)
        filled = sum(1 for v in found.values() for x in v.values() if x)
        _log(f"batch {n}: {len(found)} companies answered, {filled} real values returned")

    _apply(args, rows, results, state)


def _apply(args, rows: list[dict], results: dict, state) -> None:
    stats = {"kept": 0, "blanked": 0, "unanswered": 0, "preserved": 0}
    samples: list[str] = []
    allowed = set(getattr(args, "_columns", list(_FIELD_TO_COLUMN.values())))

    for row in rows:
        name = row.get("Company", "")
        answer = results.get(name)
        if answer is None:
            # The batch failed (CAPTCHA/parse error). Leave the row alone
            # rather than blanking on the strength of a failed query — a
            # failed lookup is not evidence the existing value is wrong.
            stats["unanswered"] += 1
            continue
        for field, column in _FIELD_TO_COLUMN.items():
            if column not in allowed:
                continue  # not requested: leave completely untouched
            new = answer.get(field, "")
            old = str(row.get(column, "") or "")

            # In --only-empty mode an existing value is authoritative: it
            # was either verified earlier or supplied by the user, so a
            # fresh lookup must not replace or erase it.
            if args.only_empty and old:
                stats["preserved"] += 1
                continue

            if new:
                if new != old and len(samples) < 12:
                    samples.append(f"{name[:28]:28s} {column:15s} {old[:24]!r} -> {new[:34]!r}")
                row[column] = new
                stats["kept"] += 1
            elif args.only_empty:
                # Was empty, still unknown — leave it empty, but do not
                # count it as a blanking (nothing was lost).
                row[column] = ""
            else:
                if old and len(samples) < 12:
                    samples.append(f"{name[:28]:28s} {column:15s} {old[:24]!r} -> BLANK")
                row[column] = ""
                if old:
                    stats["blanked"] += 1

    print()
    print("=" * 66)
    print(f"real values kept/updated : {stats['kept']}")
    print(f"unverified -> blanked    : {stats['blanked']}")
    print(f"companies not answered   : {stats['unanswered']} (left unchanged)")
    if stats["preserved"]:
        print(f"already filled, preserved: {stats['preserved']}")
    print("=" * 66)
    if samples:
        print("Sample changes:")
        for s in samples:
            print(f"  {s}")
        print()

    if args.dry_run:
        print("DRY RUN — state and Excel were NOT modified.")
        return

    state_path = STATE_DIR / f"{market_slug(args.market)}.json"
    if not state_path.is_file():
        _log(f"state file missing, cannot write: {state_path}")
        return

    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup = state_path.with_name(f"{state_path.stem}.before-refill-{stamp}.json")
    shutil.copy2(state_path, backup)
    _log(f"state backed up to {backup.name}")

    # Merge back by company name: `rows` may be a --limit slice, so the
    # untouched companies must keep their existing values.
    data = json.loads(state_path.read_text(encoding="utf-8"))
    updated = {r.get("Company", ""): r for r in rows}
    for i, existing in enumerate(data.get("companies", [])):
        replacement = updated.get(existing.get("Company", ""))
        if replacement is not None:
            data["companies"][i] = replacement

    state_path.write_text(json.dumps(data, indent=1, ensure_ascii=False), encoding="utf-8")
    _log(f"state written ({len(data.get('companies', []))} companies, none dropped)")

    out = export_xlsx(args.market)
    _log(f"re-exported: {out}")
    print()
    print("Blank cells in those six columns now mean 'not known', not 'missing'.")


if __name__ == "__main__":
    main()
