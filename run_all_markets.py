#!/usr/bin/env python3
"""Run discovery + export + verification for several markets, one after another.

One command for the whole queue. Each market runs to completion (discover
-> export -> KEEP/DROP verify) before the next one starts, because they
all drive the SAME persistent browser profile: two at once collide on the
profile lock, and two concurrent query streams from one IP is what gets
Google to rate-limit the session.

Why a script rather than chaining commands in the shell: a plain `&&`
chain aborts the whole queue when one market fails, and `;` hides which
ones failed. Here every market is independent — a crash, a CAPTCHA block
or a genuinely exhausted market stops that market only, the rest still
run, and the summary at the end says exactly what happened to each.

Everything is resumable. Discovery keeps its own per-market state file and
verification checkpoints every chunk of 10, so re-running this command
after an interruption continues where it stopped instead of restarting.

Usage:
    # the whole default queue
    python run_all_markets.py

    # a different target, or specific markets only
    python run_all_markets.py --target 300
    python run_all_markets.py --only silicon --only cpaas

    # see the plan without running anything
    python run_all_markets.py --dry-run

    # skip the verification pass
    python run_all_markets.py --no-verify
"""
from __future__ import annotations

import argparse
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent

# Windows consoles are cp1252 and print() RAISES on an unmappable
# character. This queue collects companies worldwide, so non-ASCII names
# are normal input — a Turkish name once killed an entire run this way.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

# The market strings are the pipeline's identity for a market: the state
# file, the Excel filename and the discovery prompt all derive from them,
# so changing one starts a NEW market rather than resuming an existing.
#
# CPaaS is deliberately last: its distributor definition is the inverted
# one (channel partners ARE the distributors there), so if anything about
# the definition routing is wrong it shows up after the straightforward
# markets have already banked their results.
MARKETS: tuple[str, ...] = (
    "Global Silicon Carbide Market - Distributors Only",
    "Global Wearable Glucometer Market - Distributors Only",
    "Global Dry Eyes Treatment Market - Distributors Only",
    "Global Solar Rooftop Market - Distributors Only",
    "Global Polyurethane Market - Distributors Only",
    "Global CPaaS (Communication Platform as a Service) Market - Distributors Only",
)


def _log(msg: str) -> None:
    print(msg, flush=True)


def _saved_count(market: str) -> int:
    """Companies already collected, so progress is visible across restarts."""
    try:
        import json

        sys.path.insert(0, str(ROOT / "src"))
        from vendor_intel.pipeline.market_research.state import STATE_DIR, market_slug

        path = STATE_DIR / f"{market_slug(market)}.json"
        if not path.is_file():
            return 0
        return len(json.loads(path.read_text(encoding="utf-8")).get("companies", []))
    except Exception:
        return 0


def run_market(market: str, *, target: int, verify: bool, batch_size: int) -> dict:
    """Run one market to completion. Never raises — the queue must go on."""
    started = time.time()
    before = _saved_count(market)

    cmd = [
        sys.executable,
        str(ROOT / "run_market_research.py"),
        "--market", market,
        "--target", str(target),
        "--batch-size", str(batch_size),
        "--export-xlsx",
    ]
    if verify:
        cmd.append("--verify")

    _log("")
    _log("#" * 78)
    _log(f"# {market}")
    _log(f"# target {target} | already saved {before} | verify {'on' if verify else 'off'}")
    _log(f"# started {datetime.now():%H:%M:%S}")
    _log("#" * 78)

    try:
        code = subprocess.run(cmd, cwd=str(ROOT)).returncode
    except KeyboardInterrupt:
        # Ctrl+C should stop the QUEUE, not be swallowed as a market
        # failure — progress is already checkpointed either way.
        raise
    except Exception as exc:  # noqa: BLE001
        code = -1
        _log(f"! {type(exc).__name__}: {exc}")

    after = _saved_count(market)
    mins = (time.time() - started) / 60
    status = "ok" if code == 0 else f"exit {code}"
    _log("")
    _log(f"--> {market}: {status} | {before} -> {after} companies | {mins:.1f} min")
    return {
        "market": market, "status": status, "before": before,
        "after": after, "minutes": mins,
    }


def main() -> None:
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--target", type=int, default=400,
                    help="companies to collect per market (default 400)")
    ap.add_argument("--batch-size", type=int, default=20)
    ap.add_argument("--no-verify", action="store_true",
                    help="skip the KEEP/DROP verification pass")
    ap.add_argument("--only", action="append", default=[], metavar="TEXT",
                    help="run only markets whose name contains TEXT (repeatable)")
    ap.add_argument("--dry-run", action="store_true",
                    help="print the queue and exit without running")
    args = ap.parse_args()

    queue = list(MARKETS)
    if args.only:
        needles = [n.lower() for n in args.only]
        queue = [m for m in queue if any(n in m.lower() for n in needles)]
        if not queue:
            _log(f"no market matches {args.only} — nothing to do")
            return

    verify = not args.no_verify

    _log("=" * 78)
    _log(f"QUEUE: {len(queue)} market(s), target {args.target} each, "
         f"verification {'ON' if verify else 'OFF'}")
    _log("Markets run one at a time — they share one browser profile and one IP.")
    _log("Everything is resumable: re-run this command to continue after a stop.")
    _log("=" * 78)
    for n, m in enumerate(queue, 1):
        saved = _saved_count(m)
        _log(f"  {n}. {m}" + (f"   [{saved} already saved]" if saved else ""))

    if args.dry_run:
        _log("\nDRY RUN — nothing was executed.")
        return

    results = []
    try:
        for market in queue:
            results.append(run_market(
                market, target=args.target, verify=verify,
                batch_size=args.batch_size,
            ))
    except KeyboardInterrupt:
        _log("\ninterrupted — progress is saved; re-run to resume")

    _log("")
    _log("=" * 78)
    _log("SUMMARY")
    _log("=" * 78)
    _log(f"  {'market':<58}{'status':<10}{'companies':>10}")
    total_new = 0
    for r in results:
        gained = r["after"] - r["before"]
        total_new += gained
        _log(f"  {r['market'][:56]:<58}{r['status']:<10}"
             f"{r['after']:>6} (+{gained})")
    skipped = len(queue) - len(results)
    _log("")
    _log(f"  {total_new} companies added across {len(results)} market(s)"
         + (f", {skipped} not reached" if skipped else ""))
    _log("")
    _log("  Discovery output : output/market_research/<market-slug>.xlsx")
    if verify:
        _log("  Verified output  : output/market_research/"
             "<market-slug>_verified_distributors.xlsx")
        _log("  Audit log        : output/market_research/"
             "<market-slug>_verify_log.json")


if __name__ == "__main__":
    main()
