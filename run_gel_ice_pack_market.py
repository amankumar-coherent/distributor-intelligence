#!/usr/bin/env python3
"""
Global Gel Ice Pack Market — discovery + FMCG Excel export (30 columns).

Usage:
  .venv\\Scripts\\python.exe run_gel_ice_pack_market.py
  .venv\\Scripts\\python.exe run_gel_ice_pack_market.py --mock
  .venv\\Scripts\\python.exe run_gel_ice_pack_market.py --full-crawl --profile recall --cap maximum

Then deep-enrich all companies:
  .venv\\Scripts\\python.exe run_gel_ice_pack_scrape.py --limit 10
  .venv\\Scripts\\python.exe run_gel_ice_pack_scrape.py
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

QUERY = "Global Gel Ice Pack Market"
COUNTRY = "global"
OUTPUT_SUBDIR = ROOT / "output" / "gel_ice_pack"
SEEDS_FILE = ROOT / "queries" / "seeds" / "global_gel_ice_pack_market.txt"

SECTIONS = [
    "Manufacturers of Gel Ice Packs",
    "Cold Chain / Logistics",
    "Medical / Healthcare Distributors",
    "Retail / Consumer Brands",
    "OEM / Private Label",
    "Suppliers / Raw Materials",
    "Other",
]


def _slug(text: str) -> str:
    return "".join(c if c.isalnum() else "_" for c in text.lower()).strip("_")[:80]


def _load_seeds(path: Path) -> tuple[list[str], dict[str, str]]:
    seed_sections: dict[str, str] = {}
    seed_names: list[str] = []
    if not path.is_file():
        return seed_names, seed_sections
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if "|" in line:
            name, sec = line.split("|", 1)
            name, sec = name.strip(), sec.strip()
            if name:
                seed_names.append(name)
                if sec:
                    seed_sections[name] = sec
        else:
            seed_names.append(line)
    return seed_names, seed_sections


def main() -> None:
    parser = argparse.ArgumentParser(description="Global Gel Ice Pack Market pipeline + Excel export")
    parser.add_argument("--mock", action="store_true", help="Mock mode (no live search)")
    parser.add_argument(
        "--profile",
        choices=("quality", "balanced", "recall"),
        default="recall",
        help="Pipeline profile (default: recall)",
    )
    parser.add_argument(
        "--cap",
        choices=("focused", "standard", "broad", "maximum"),
        default="maximum",
        help="Company count cap (default: maximum)",
    )
    parser.add_argument("--seeds-file", default=str(SEEDS_FILE))
    parser.add_argument(
        "--full-crawl",
        action="store_true",
        help="Use smart_crawl instead of fast SSC (slower, richer Excel columns)",
    )
    args = parser.parse_args()

    from vendor_intel.config import Settings
    from vendor_intel.export.fmcg_landscape import export_fmcg_xlsx
    from vendor_intel.pipeline.orchestrator import run_pipeline_sync, save_pipeline_csv
    from vendor_intel.placeholders.load_keys import apply_env_overrides

    apply_env_overrides()
    settings = Settings.load()
    if args.mock:
        settings = settings.model_copy(update={"use_mock_data": True, "mock_mode": True})
    else:
        settings = settings.model_copy(update={"use_mock_data": False, "mock_mode": False})

    from vendor_intel.pipeline.cap_profiles import apply_cap, cap_describe

    updates: dict = {
        "pipeline_profile": args.profile,
        "pipeline_recall_mode": args.profile == "recall",
    }
    if args.full_crawl:
        updates["pipeline_use_ssc"] = False
    settings = settings.model_copy(update=updates)
    settings = apply_cap(settings, args.cap)

    seed_names, seed_sections = _load_seeds(Path(args.seeds_file))

    query_context = {
        "industry": QUERY,
        "country": COUNTRY,
        "functions": [
            "manufacturer",
            "distributor",
            "retailer",
            "brand",
            "exporter",
            "importer",
            "logistics",
        ],
        "sections": SECTIONS,
        "seed_companies": list(dict.fromkeys(seed_names)),
        "seed_sections": seed_sections,
    }

    print(f"\n{'=' * 60}")
    print(f"  Market : {QUERY}")
    print(f"  Geo    : {COUNTRY}")
    print(f"  Profile: {args.profile} | Cap: {cap_describe(args.cap)}")
    print(f"  Crawl  : {'smart_crawl (deep)' if args.full_crawl else 'SSC (fast)'}")
    print(f"  Seeds  : {len(seed_names)} companies")
    print(f"{'=' * 60}\n")

    t0 = time.perf_counter()
    result = run_pipeline_sync(query_context, settings)
    elapsed = time.perf_counter() - t0

    slug = _slug(f"{QUERY}_{COUNTRY}")
    OUTPUT_SUBDIR.mkdir(parents=True, exist_ok=True)
    csv_path = OUTPUT_SUBDIR / f"{slug}.csv"
    xlsx_path = OUTPUT_SUBDIR / f"{slug}.xlsx"
    json_path = OUTPUT_SUBDIR / f"{slug}.json"

    save_pipeline_csv(result, str(csv_path))
    xlsx_out = export_fmcg_xlsx(result, xlsx_path)
    json_path.write_text(json.dumps(result, indent=2, default=str), encoding="utf-8")

    n = len(result.get("relevant_companies") or [])
    llm = result.get("llm_usage") or {}
    print(f"\n{'=' * 60}")
    print(f"  Done in {elapsed / 60:.1f} min")
    print(f"  Companies : {n}")
    print(f"  LLM calls : {llm.get('llm_calls_total', '?')}")
    print(f"  CSV       : {csv_path.resolve()}")
    print(f"  Excel     : {xlsx_out}")
    print(f"  JSON      : {json_path.resolve()}")
    print(f"\n  Next: .venv\\Scripts\\python.exe run_gel_ice_pack_scrape.py --limit 10")
    print(f"{'=' * 60}\n")


if __name__ == "__main__":
    main()
