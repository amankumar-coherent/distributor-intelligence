#!/usr/bin/env python3
"""Deep audit with web search verification — thin CLI over market_audit.

Run per-market:
  python final_audit_v2.py --market avocado_oil
  python final_audit_v2.py --market technical_foam

Or all preset markets:
  python final_audit_v2.py

Each company: live web search + GPT verdict (keep / review / remove).
Removes "remove" rows from CSV/XLSX; writes audit_cache.json and audit_report.md.
"""
from __future__ import annotations

import argparse
import asyncio
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

from vendor_intel.enrichment.market_audit import audit_market_exports  # noqa: E402

MARKETS = [
    "avocado_oil",
    "technical_foam",
    "spectral_sensor",
    "gel_ice_pack",
    "cyber_physical_systems_security",
    "high_energy_laser",
]

MARKET_DESCRIPTIONS = {
    "avocado_oil": (
        "avocado oil production, processing, refining, bottling, or distribution "
        "as a food ingredient or cosmetic raw material"
    ),
    "technical_foam": (
        "technical or industrial foam manufacturing — polyurethane, acoustic, thermal, "
        "EVA, silicone, neoprene, melamine foam for industrial/commercial applications"
    ),
    "spectral_sensor": (
        "spectral sensors, hyperspectral imaging, NIR/SWIR sensors, Raman spectrometers, "
        "FTIR instruments, multispectral cameras for industrial or scientific use"
    ),
    "gel_ice_pack": (
        "gel ice packs, cold packs, reusable cold therapy packs, PCM cold chain packs, "
        "or temperature-controlled packaging manufacturing"
    ),
    "cyber_physical_systems_security": (
        "cybersecurity for cyber-physical systems (CPS), OT/ICS security, SCADA protection, "
        "industrial control system security, critical infrastructure protection"
    ),
    "high_energy_laser": (
        "high energy laser systems, directed energy weapons, high-power fiber/solid-state/disk "
        "lasers (>1kW), or laser weapon systems for defense or heavy industry"
    ),
}


async def audit_market(market: str, *, concurrent: int = 8) -> dict:
    output_dir = ROOT / "output" / market
    market_desc = MARKET_DESCRIPTIONS.get(market, market.replace("_", " "))
    market_query = market.replace("_", " ").title()

    print(f"\n{'=' * 65}")
    print(f"  AUDITING: {market_query.upper()}")
    print(f"{'=' * 65}")

    stats = await audit_market_exports(
        output_dir=output_dir,
        market_query=market_query,
        market_desc=market_desc,
        concurrent=concurrent,
    )
    if stats.get("skipped"):
        return {"market": market, "error": stats.get("reason", "skipped")}

    return {
        "market": market,
        "total": stats.get("total", 0),
        "keep": stats.get("keep", 0),
        "review": stats.get("review", 0),
        "remove": stats.get("remove", 0),
        "removed_names": stats.get("removed_names", []),
    }


async def main(target_markets: list[str], *, concurrent: int = 8) -> None:
    print("=" * 65)
    print("  FINAL AUDIT v2 — Web Search + GPT Verification")
    print("  Each company: Live web search evidence + AI verdict")
    print("  Actions: keep (confirmed) / review (flag) / remove (delete)")
    print("=" * 65)
    print(f"  Markets: {', '.join(target_markets)}")

    results = []
    for market in target_markets:
        try:
            results.append(await audit_market(market, concurrent=concurrent))
        except Exception as exc:
            import traceback

            print(f"  [ERROR] {market}: {exc}")
            traceback.print_exc()
            results.append({"market": market, "error": str(exc)})

    print("\n" + "=" * 65)
    print("  AUDIT COMPLETE — GRAND SUMMARY")
    print("=" * 65)
    print(f"\n  {'Market':<42} {'Total':>6} {'Keep':>6} {'Review':>7} {'Remove':>7}")
    print(f"  {'-' * 70}")
    grand_total = grand_keep = grand_review = grand_remove = 0
    for r in results:
        if "error" not in r:
            print(
                f"  {r['market']:<42} {r['total']:>6} {r['keep']:>6} "
                f"{r['review']:>7} {r['remove']:>7}"
            )
            grand_total += r["total"]
            grand_keep += r["keep"]
            grand_review += r["review"]
            grand_remove += r["remove"]
    print(f"  {'-' * 70}")
    print(
        f"  {'TOTAL':<42} {grand_total:>6} {grand_keep:>6} "
        f"{grand_review:>7} {grand_remove:>7}"
    )
    print("=" * 65)

    summary_path = ROOT / "output" / "final_audit_v2_summary.json"
    summary_path.parent.mkdir(parents=True, exist_ok=True)
    summary_path.write_text(json.dumps(results, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\n  Summary saved: {summary_path}")
    print("  Per-market reports in: output/<market>/audit_report.md")
    print("  Per-market cache in:   output/<market>/audit_cache.json")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Final audit with web search verification")
    parser.add_argument(
        "--market",
        choices=MARKETS,
        default=None,
        help="Audit a specific market only (default: all preset markets)",
    )
    parser.add_argument(
        "--concurrent",
        type=int,
        default=8,
        help="Parallel company audits (default: 8)",
    )
    args = parser.parse_args()
    target_markets = [args.market] if args.market else MARKETS
    asyncio.run(main(target_markets, concurrent=args.concurrent))
