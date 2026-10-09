#!/usr/bin/env python3
"""Global Market Research & Company Database pipeline CLI.

Usage:
    python run_market_research.py --market "Global Semiconductor Distributor Market"
    python run_market_research.py --market "Global Semiconductor Distributor Market" --more
    python run_market_research.py --market "Global Semiconductor Distributor Market" --target 500
    python run_market_research.py --market "Global Semiconductor Distributor Market" --export-xlsx

First run for a market creates fresh state and returns batch 1 (companies
1-20). Passing --more (or re-running the same --market again) continues
numbering and excludes every company already returned for that market,
mirroring the "give me more N" continuation from the spec. State persists
under data/market_research/<market-slug>.json.

--target N automatically repeats the "give me N more, unique from above"
loop internally until N total verified companies are collected (or nothing
new can be found), instead of requiring a separate --more call per batch.

--export-xlsx writes every company accumulated so far for the market to
output/market_research/<market-slug>.xlsx, after a final dedup +
completeness re-check across the whole list (any duplicate or row with a
missing/NPD field is dropped and logged to the workbook's "Export Audit"
sheet rather than exported).

BACKEND: Google AI Mode (udm=50), driven by Patchright via
vendor_intel.scraping.google_ai_mode. No account, no login and no API key
is needed — it runs anonymously against Google. Requires the real Chrome
install (used via channel="chrome") and the `patchright` package from
requirements.txt; plain `playwright` is NOT a substitute (Google instantly
CAPTCHAs it). Set GOOGLE_AI_MODE_ENABLED=false to disable, in which case
the AI Overview extension relay is used instead.
"""
import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

# Force UTF-8 console output.
#
# Windows consoles default to cp1252, which cannot encode most non-Western
# characters — and print() then RAISES UnicodeEncodeError, killing the
# whole run. Observed live: a Turkish distributor whose name contains "Ş"
# (U+015E) crashed a run mid-batch at 272 companies, losing that batch.
#
# This pipeline is explicitly global (the prompt now steers toward Latin
# America, the Middle East, Africa and Asia), so non-ASCII company names
# are the NORM, not an edge case. Reconfiguring stdout is the one change
# that fixes every print site at once; errors="replace" guarantees that
# even an unmappable glyph degrades to "?" rather than aborting.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass  # already UTF-8, or a stream that cannot be reconfigured

from vendor_intel.placeholders.load_keys import apply_env_overrides
from vendor_intel.pipeline.market_research.batch_runner import run_batch, run_until_target
from vendor_intel.pipeline.market_research.export_xlsx import export_xlsx


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--market", required=True, help='e.g. "Global Semiconductor Distributor Market"')
    parser.add_argument("--more", action="store_true", help="Continue an existing market's batches (explicit form; also implied automatically once state exists)")
    parser.add_argument("--batch-size", type=int, default=20, help="Companies requested per prompt round (capped at 20)")
    parser.add_argument("--target", type=int, default=None, help="Keep looping batches until this many total verified companies are collected")
    parser.add_argument("--max-batches", type=int, default=None, help="Safety cap on how many batches --target will run (default: unlimited)")
    parser.add_argument("--export-xlsx", action="store_true", help="Write the full accumulated company list to output/market_research/<slug>.xlsx")
    parser.add_argument("--verify", action="store_true", help="After export, run the KEEP/DROP verification pass over the discovered companies (writes a separate *_verified_distributors.xlsx)")
    args = parser.parse_args()

    apply_env_overrides()

    # Write the market definition to the Excel BEFORE any searching.
    #
    # Requested explicitly: define the market, put it in the sheet, and
    # only then look for companies. Previously the file appeared only
    # after a run finished, so during a multi-hour run there was nothing
    # to look at and no record of which inclusion rule was in force. The
    # header is rewritten again after discovery, so the finished file is
    # identical either way - this only makes the definition visible up
    # front rather than at the end.
    if args.export_xlsx:
        header_path = export_xlsx(args.market)
        print(f"Market definition written to: {header_path}")
        print("(companies are appended as batches complete)")
        print()

    if args.target is not None:
        summary = run_until_target(
            args.market, target=args.target, batch_size=args.batch_size, max_batches=args.max_batches
        )
        print(summary)
    else:
        output = run_batch(args.market, batch_size=args.batch_size)
        print(output)

    if args.export_xlsx:
        path = export_xlsx(args.market)
        print(f"\nExcel export written to: {path}")

        if args.verify:
            _run_verification(path, args.market)



def _market_name_for_criteria(market: str) -> str:
    """Market name without the pipeline's "- Distributors Only" suffix."""
    name = market.strip()
    for suffix in (
        "- Distributors Only", "- Distributors", "- Distributor Only",
        "- Distributor",
    ):
        if name.lower().endswith(suffix.lower()):
            return name[: -len(suffix)].rstrip(" -")
    return name


def _run_verification(xlsx_path, market: str) -> None:
    """Second-pass KEEP/DROP verification of the discovered companies.

    Runs AFTER discovery and export, per explicit instruction (when all
    companies are in, then verify). Kept as a separate pass because the
    two answer different questions: discovery asks who is in this market,
    verification asks whether this specific company genuinely operates as
    a distributor here, with the row's own researched categories/focus
    handed back to the model as evidence.

    Non-destructive by design: the discovery export is left untouched and
    KEEP rows go to a separate <name>_verified_distributors.xlsx, so a bad
    verification pass can never destroy the discovery output. Verification
    is checkpointed per chunk of 10, so an interruption resumes instead of
    restarting.
    """
    import subprocess

    print()
    print("=" * 70)
    print("VERIFICATION PASS - judging each company KEEP / DROP")
    print("Writes a SEPARATE *_verified_distributors.xlsx; the export above")
    print("is left unchanged. Progress is checkpointed per chunk of 10.")
    print("=" * 70)

    cmd = [
        sys.executable,
        str(ROOT / "verify_distributors_ai_mode.py"),
        "--file", str(xlsx_path),
        # Strip the pipeline's filter suffix: it describes the FILTER, not
        # the market, and fed into the criteria prompt it read as
        # "...Market - Distributors Only market".
        "--market-name", _market_name_for_criteria(market),
    ]
    # A subprocess, not an import: the verifier drives its own Patchright
    # session and calls shutdown() when done. Running it in this process
    # would contend with the discovery run's own browser handle for the
    # same persistent profile lock - exactly the 'Opening in existing
    # browser session' collision seen live.
    # Hand the profile over before the child starts. The ancestor
    # exemption in _claim_profile covers this too, but releasing here is
    # unconditional and does not depend on walking the process tree --
    # and Windows recycles PIDs aggressively, so the walk is not something
    # to rely on alone.
    try:
        from vendor_intel.scraping import google_ai_mode as _gam

        _gam.shutdown()
    except Exception:  # noqa: BLE001
        pass

    result = subprocess.run(cmd, cwd=str(ROOT))
    if result.returncode != 0:
        print(f"verification exited with code {result.returncode} - "
              "progress is checkpointed, re-run to resume")


if __name__ == "__main__":
    main()
