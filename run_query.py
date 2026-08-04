#!/usr/bin/env python3
"""
run_query.py — Interactive Vendor Intelligence query runner.

Run one market query at a time: type a query, get a CSV, then run another or quit.

Usage:
  python run_query.py
  python run_query.py --query "Atomic Clock Market"
  python run_query.py --query "Atomic Clock Market" --country Germany
  python run_query.py --query "Atomic Clock Market" --mode bulk
  python run_query.py --file my_queries.txt --mode bulk --no-post-enrich
  python run_query.py --profile recall

Output (default demo folder — override with MARKET_QUERY_OUTPUT_DIR in .env):
  output/demo/<slug>/<slug>.csv
  output/demo/<slug>/<slug>.json
  output/demo/<slug>/<slug>_founded_and_hq_v2.xlsx  (post-enrich)
  output/demo/<slug>/_neon_filtered_export.json     (post-enrich)
  output/demo/session_log.json

After export, post-enrich runs automatically (gap-fill → LinkedIn verify Founded+HQ → Apollo
→ contacts → brand fix → market audit → Founded+HQ Excel → Neon). Use --no-post-enrich to skip.

Resume after Ctrl+C (skips finished steps + already-done companies):
  python run_query.py --resume --query "Global Smartwatch Market" --push-neon
  python run_query.py --resume --market-dir output/demo/global_smartwatch_market_global
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

def _refresh_output_dir() -> Path:
    global OUTPUT_DIR
    from vendor_intel.pipeline.output_paths import market_query_output_dir

    OUTPUT_DIR = market_query_output_dir(ROOT)
    return OUTPUT_DIR


OUTPUT_DIR = ROOT / "output" / "demo"


def _validate_env() -> list[str]:
    warnings: list[str] = []
    mock = os.getenv("USE_MOCK_DATA", "true").strip().lower()
    if mock not in ("false", "0", "no", "off"):
        warnings.append(
            "USE_MOCK_DATA is not 'false' — results will be empty/fake.\n"
            "  Fix: open .env and set USE_MOCK_DATA=false"
        )
    provider = os.getenv("LLM_PROVIDER", "anthropic").strip().lower()
    key_map = {
        "anthropic": "ANTHROPIC_API_KEY",
        "opencode": "OPENCODE_API_KEY",
        "gemini": "GEMINI_API_KEY",
        "groq": "GROQ_API_KEY",
        "openai": "OPENAI_API_KEY",
    }
    key_var = key_map.get(provider)
    if key_var and not os.getenv(key_var, "").strip():
        warnings.append(
            f"LLM_PROVIDER={provider} but {key_var} is not set.\n"
            f"  Fix: open .env and add {key_var}=your_key_here"
        )
    return warnings


# Shared with run_batch.py — see src/vendor_intel/pipeline/run_and_export.py.
from vendor_intel.pipeline.run_and_export import (
    auto_seed_names as _auto_seed_names,
    execute_and_export,
    load_seeds_file as _load_seeds_file,
    parse_market_line as _parse_input,
    parse_sections as _parse_sections,
    slugify as _slug,
)


def _load_settings(profile: str):
    from vendor_intel.placeholders.load_keys import apply_env_overrides
    from vendor_intel.config import Settings

    apply_env_overrides()
    settings = Settings.load()
    return settings.model_copy(
        update={
            "use_mock_data": False,
            "mock_mode": False,
            "pipeline_profile": profile,
            "pipeline_recall_mode": profile == "recall",
            "pipeline_use_ssc": True,
            "pipeline_min_export_confidence": 0.50,
        }
    )


def _push_result_to_neon(result: dict) -> None:
    """Upsert exported companies into Neon after a successful run."""
    from vendor_intel.neon_push import push_result_to_neon

    push_result_to_neon(result, label="neon")


def _should_push_to_neon(*, push_neon: bool, no_push_neon: bool) -> bool:
    """Auto-push when DATABASE_URL is set unless explicitly disabled."""
    if no_push_neon:
        return False
    if push_neon:
        return True
    env = os.getenv("PUSH_TO_NEON", "").strip().lower()
    if env in ("0", "false", "no", "off"):
        return False
    if env in ("1", "true", "yes", "on"):
        return True
    return bool(os.getenv("DATABASE_URL", "").strip())


def _maybe_push_to_neon(result: dict, *, push_neon: bool, no_push_neon: bool) -> None:
    if _should_push_to_neon(push_neon=push_neon, no_push_neon=no_push_neon):
        _push_result_to_neon(result)


def _run_post_enrich(
    result: dict,
    *,
    push_neon: bool,
    no_push_neon: bool,
    post_market_kwargs: dict | None = None,
) -> dict | None:
    """Gap-fill → Apollo names → resolve partials → Excel → Neon (Founded+HQ)."""
    market_dir = result.get("_market_dir") or (
        str(Path(result["_json_path"]).parent) if result.get("_json_path") else ""
    )
    if not market_dir:
        print("  [post-enrich] Skipped — no export folder in result", flush=True)
        return None
    from vendor_intel.pipeline.post_market import run_post_market_pipeline_sync

    do_push = _should_push_to_neon(push_neon=push_neon, no_push_neon=no_push_neon)
    kwargs = dict(post_market_kwargs or {})
    kwargs.setdefault("push_neon", do_push)
    kwargs.setdefault("resume", True)
    return run_post_market_pipeline_sync(Path(market_dir), **kwargs)


def _print_post_enrich(stats: dict | None) -> None:
    if not stats:
        return
    if stats.get("error"):
        print(f"  [post-enrich] FAILED: {stats['error']}", flush=True)
        return
    bar = "-" * 60
    print(f"\n{bar}")
    print("  Post-enrich complete")
    print(f"  Time               : {stats.get('elapsed_sec', '?')}s")
    if stats.get("filtered_companies") is not None:
        print(f"  Founded+HQ rows    : {stats['filtered_companies']}")
    excel = stats.get("excel") or {}
    if excel.get("file"):
        print(f"  Excel              : {excel['file']}")
    if stats.get("neon_json"):
        print(f"  Neon export JSON   : {stats['neon_json']}")
    brand = stats.get("brand_resolve") or {}
    if brand.get("corrected"):
        print(f"  Brand/domain fixed : {brand['corrected']} companies")
    audit = stats.get("market_audit") or {}
    if audit.get("queued") is not None:
        print(
            f"  Market audit       : {audit.get('kept', 0)} keep, "
            f"{audit.get('review', 0)} review, {audit.get('removed', 0)} removed"
        )
    li_verify = stats.get("linkedin_firmographic_verify") or {}
    if li_verify.get("queued") is not None:
        print(
            f"  LinkedIn verify    : {li_verify.get('updated', 0)} updated, "
            f"{li_verify.get('verified_both', 0)} Founded+HQ confirmed "
            f"(delay ~{li_verify.get('delay_sec', '?')}s/call)"
        )
    if audit.get("report_path"):
        print(f"  Audit report       : {audit['report_path']}")
    neon = stats.get("neon_push") or {}
    if neon.get("vendor_intel_companies"):
        print(f"  Neon pushed        : {neon['vendor_intel_companies']} companies")
    elif neon.get("skipped"):
        print(f"  Neon push          : skipped ({neon.get('reason', 'disabled')})")
    print(f"{bar}\n")


def run_one_query(
    query: str,
    country: str,
    settings,
    *,
    enrich_limit: int | None = None,
    classify_limit: int | None = None,
    sections: list[str] | None = None,
    seeds: list[str] | None = None,
    out_name: str | None = None,
    fmcg_xlsx: bool = False,
    resume_checkpoint: bool = False,
) -> dict:
    return execute_and_export(
        query,
        country,
        settings,
        OUTPUT_DIR,
        enrich_limit=enrich_limit,
        classify_limit=classify_limit,
        sections=sections,
        seeds=seeds,
        out_name=out_name,
        fmcg_xlsx=fmcg_xlsx,
        resume_checkpoint=resume_checkpoint,
    )


def _print_result(query: str, country: str, result: dict, elapsed: float) -> None:
    companies = len(result.get("relevant_companies") or [])
    llm = result.get("llm_usage") or {}
    bar = "-" * 60
    print(f"\n{bar}")
    print(f"  Done: {query} ({country})")
    print(f"  Companies exported : {companies}")
    print(f"  Time               : {elapsed / 60:.1f} min ({elapsed:.0f}s)")
    print(f"  LLM calls          : {llm.get('llm_calls_total', '?')}")
    print(f"  Est. cost          : ${llm.get('estimated_cost_usd', '?')}")
    print(f"  CSV                : {result.get('_csv_path', '')}")
    print(f"  Word doc           : {result.get('_docx_path', '')}")
    print(f"  Excel (FMCG)       : {result.get('_xlsx_path', '')}")
    print(f"  JSON               : {result.get('_json_path', '')}")
    print(f"{bar}\n")


def _load_session_log() -> list[dict]:
    log_path = OUTPUT_DIR / "session_log.json"
    if log_path.exists():
        try:
            return json.loads(log_path.read_text(encoding="utf-8"))
        except Exception:
            pass
    return []


def _save_session_log(log: list[dict]) -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    (OUTPUT_DIR / "session_log.json").write_text(
        json.dumps(log, indent=2, default=str), encoding="utf-8"
    )


def _parse_query_and_country(raw: str) -> tuple[str, str]:
    if "|" in raw:
        parts = raw.split("|", 1)
        return parts[0].strip(), (parts[1].strip() or "global")
    return raw.strip(), "global"


def _print_history(log: list[dict]) -> None:
    if not log:
        print("\n  No queries run yet.\n")
        return
    print(f"\n  Session history ({len(log)} queries):")
    for i, entry in enumerate(log, 1):
        status = "OK  " if entry.get("status") == "ok" else "FAIL"
        q = entry.get("query", "?")
        c = entry.get("country", "global")
        n = entry.get("companies_exported", "?")
        t = entry.get("elapsed_minutes", "?")
        ran = (entry.get("ran_at") or "")[:16].replace("T", " ")
        print(f"    {i:2}. [{status}] {q} ({c}) — {n} rows, {t} min  [{ran}]")
    print()


def interactive_mode(
    settings, enrich_cap: int, classify_cap: int,
    sections: list[str] | None = None, seeds: list[str] | None = None,
    *, push_neon: bool = False, no_push_neon: bool = False,
    post_enrich: bool = True,
    post_market_kwargs: dict | None = None,
) -> None:
    session_log = _load_session_log()
    print("\n" + "=" * 60)
    print("  INTERACTIVE QUERY MODE")
    print("  Type a market query and press Enter.")
    print("  With geography : \"Atomic Clock Market | Germany\"")
    print("  With sections  : \"Drone Battery Market | global | "
          "Battery Pack Manufacturers; OEM Battery Manufacturers; BMS Manufacturers\"")
    print("  Or keyword     : \"Drone Battery Market ; functionality: "
          "Battery Pack Manufacturers, BMS Manufacturers, Contract Manufacturer\"")
    print("  Commands: history, quit")
    print("=" * 60)

    while True:
        print()
        try:
            raw = input("  Query: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n  Exiting.")
            break

        if not raw:
            continue
        if raw.lower() in ("quit", "exit", "q"):
            break
        if raw.lower() == "history":
            _print_history(session_log)
            continue

        query, country, typed_sections = _parse_input(raw)
        q_sections = typed_sections or list(sections or [])
        geo = f"({country})" if country != "global" else "(global)"
        sect_note = f" | sections: {', '.join(q_sections)}" if q_sections else ""
        print(f'\n  Will run: "{query}" {geo}{sect_note}')
        print(f"  Profile: {settings.pipeline_profile} | ~15-25 min | output -> {OUTPUT_DIR}")
        try:
            confirm = input("  Press Enter to start, or 'skip': ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            print("\n  Exiting.")
            break
        if confirm in ("skip", "s", "n", "no"):
            continue

        t0 = time.perf_counter()
        print(f"\n  [{datetime.now().strftime('%H:%M:%S')}] Starting pipeline...")
        try:
            result = run_one_query(
                query, country, settings,
                enrich_limit=enrich_cap,
                classify_limit=classify_cap,
                sections=q_sections,
                seeds=seeds,
            )
            elapsed = round(time.perf_counter() - t0, 1)
            _print_result(query, country, result, elapsed)
            if post_enrich:
                _print_post_enrich(
                    _run_post_enrich(
                        result,
                        push_neon=push_neon,
                        no_push_neon=no_push_neon,
                        post_market_kwargs=post_market_kwargs,
                    )
                )
            else:
                _maybe_push_to_neon(result, push_neon=push_neon, no_push_neon=no_push_neon)
            llm = result.get("llm_usage") or {}
            log_entry = {
                "query": query,
                "country": country,
                "status": "ok",
                "companies_exported": len(result.get("relevant_companies") or []),
                "elapsed_seconds": elapsed,
                "elapsed_minutes": round(elapsed / 60, 2),
                "llm_calls": llm.get("llm_calls_total"),
                "estimated_cost_usd": llm.get("estimated_cost_usd"),
                "csv_file": result.get("_csv_path", ""),
                "ran_at": datetime.now(timezone.utc).isoformat(),
            }
        except Exception as exc:
            elapsed = round(time.perf_counter() - t0, 1)
            print(f"\n  FAILED after {elapsed:.0f}s: {exc}")
            import traceback
            traceback.print_exc()
            log_entry = {
                "query": query,
                "country": country,
                "status": "error",
                "error": str(exc),
                "elapsed_seconds": elapsed,
                "ran_at": datetime.now(timezone.utc).isoformat(),
            }

        session_log.append(log_entry)
        _save_session_log(session_log)
        print("  [Enter] another query  |  [q] quit  |  [h] history")
        try:
            nxt = input("  > ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            break
        if nxt in ("q", "quit", "exit"):
            break
        if nxt in ("h", "history"):
            _print_history(session_log)


def batch_mode(
    query_file: str, settings, enrich_cap: int, classify_cap: int,
    sections: list[str] | None = None, seeds: list[str] | None = None,
    *, push_neon: bool = False, no_push_neon: bool = False,
    post_enrich: bool = True,
    post_market_kwargs: dict | None = None,
    fmcg_xlsx: bool = False,
) -> None:
    path = Path(query_file)
    if not path.exists():
        print(f"  ERROR: File not found: {query_file}")
        sys.exit(1)
    # Line format: "Query | Country | Section1; Section2; ..."  (country/sections optional)
    # or keyword form: "Query ; functionality: Section1, Section2, ..."
    queries: list[tuple[str, str, list[str]]] = []
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        q, c, line_sections = _parse_input(line)
        queries.append((q, c, line_sections or list(sections or [])))
    if not queries:
        print(f"  No queries in {query_file}")
        sys.exit(1)

    session_log = _load_session_log()
    for idx, (query, country, line_sections) in enumerate(queries, 1):
        print(f"\n{'=' * 60}\n  [{idx}/{len(queries)}] {query} ({country})\n{'=' * 60}")
        t0 = time.perf_counter()
        try:
            result = run_one_query(
                query, country, settings,
                enrich_limit=enrich_cap,
                classify_limit=classify_cap,
                sections=line_sections,
                seeds=seeds,
                fmcg_xlsx=fmcg_xlsx,
            )
            elapsed = round(time.perf_counter() - t0, 1)
            _print_result(query, country, result, elapsed)
            if post_enrich:
                _print_post_enrich(
                    _run_post_enrich(
                        result,
                        push_neon=push_neon,
                        no_push_neon=no_push_neon,
                        post_market_kwargs=post_market_kwargs,
                    )
                )
            else:
                _maybe_push_to_neon(result, push_neon=push_neon, no_push_neon=no_push_neon)
            llm = result.get("llm_usage") or {}
            session_log.append({
                "query": query,
                "country": country,
                "status": "ok",
                "companies_exported": len(result.get("relevant_companies") or []),
                "elapsed_seconds": elapsed,
                "elapsed_minutes": round(elapsed / 60, 2),
                "llm_calls": llm.get("llm_calls_total"),
                "estimated_cost_usd": llm.get("estimated_cost_usd"),
                "csv_file": result.get("_csv_path", ""),
                "ran_at": datetime.now(timezone.utc).isoformat(),
            })
        except Exception as exc:
            elapsed = round(time.perf_counter() - t0, 1)
            print(f"  FAILED: {exc}")
            import traceback
            traceback.print_exc()
            session_log.append({
                "query": query,
                "country": country,
                "status": "error",
                "error": str(exc),
                "elapsed_seconds": elapsed,
                "ran_at": datetime.now(timezone.utc).isoformat(),
            })
        _save_session_log(session_log)


def main() -> None:
    parser = argparse.ArgumentParser(description="Vendor Intelligence — interactive query runner")
    parser.add_argument("--query", "-q", help="Single query (non-interactive)")
    parser.add_argument("--country", "-c", default="global", help="Geography (default: global)")
    parser.add_argument("--file", "-f", help="Batch: one query per line")
    parser.add_argument(
        "--mode", "-m",
        choices=("bulk", "standard", "quality", "full"),
        default=None,
        help="Pipeline preset: bulk (fast) | standard | quality (default post-enrich) | full",
    )
    parser.add_argument(
        "--profile", "-p",
        choices=("quality", "balanced", "recall"),
        default=None,
        help="quality (default) | recall = more rows, noisier",
    )
    parser.add_argument(
        "--sections", "-s",
        default=None,
        help="Exact CEO section categories, ';'- or ','-separated. "
        "e.g. \"Battery Pack Manufacturers; OEM Battery Manufacturers; BMS Manufacturers; Contract Manufacturer\"",
    )
    parser.add_argument(
        "--seeds", "-S",
        default=None,
        help="Known companies to guarantee in the output (from a report), ';'- or ','-separated. "
        "e.g. \"FUNKE; Wieland Onda; Hexonic; Xylem; Kaori\"",
    )
    parser.add_argument(
        "--seeds-file",
        default=None,
        help="File with must-have company names (one per line; '#' comments ok). "
        "Keeps the run command short for presentations.",
    )
    parser.add_argument(
        "--cap",
        choices=("bulk", "focused", "standard", "broad", "maximum"),
        default=None,
        help="Result cap: bulk ~15-35 (fastest) | focused ~25-50 | standard ~50-90 | broad | maximum",
    )
    parser.add_argument(
        "--out",
        default=None,
        help="Output basename (no extension), e.g. waste_oil_market_global_v2. Single-query mode only.",
    )
    parser.add_argument(
        "--push-neon",
        action="store_true",
        help="After export, upsert full company export into Neon (companies / markets / metadata).",
    )
    parser.add_argument(
        "--no-push-neon",
        action="store_true",
        help="Skip Neon push even when DATABASE_URL is set (overrides auto-push).",
    )
    parser.add_argument(
        "--no-post-enrich",
        action="store_true",
        help="Skip post-export pipeline (gap-fill, Apollo names, name resolve, Excel, Neon).",
    )
    parser.add_argument(
        "--no-name-resolve",
        action="store_true",
        help="After Apollo, skip partial-name upgrade (LinkedIn/Google/RocketReach waterfall). "
        "Keeps Apollo first-name / obfuscated contacts as-is. Same as SKIP_NAME_RESOLVE=true.",
    )
    parser.add_argument(
        "--no-linkedin-verify",
        action="store_true",
        help="Skip post-market step 2 LinkedIn Founded+HQ verify. "
        "Same as SKIP_LINKEDIN_FIRMOGRAPHIC_VERIFY=true.",
    )
    parser.add_argument(
        "--no-wikidata-firmographic",
        action="store_true",
        help="Skip Wikidata Founded+HQ fallback (step 2b). Same as SKIP_WIKIDATA_FIRMOGRAPHIC=true.",
    )
    parser.add_argument(
        "--no-wikipedia-firmographic",
        action="store_true",
        help="Skip Wikipedia Founded+HQ fallback (step 2c). Same as SKIP_WIKIPEDIA_FIRMOGRAPHIC=true.",
    )
    parser.add_argument(
        "--resume",
        action="store_true",
        help="Resume POST-ENRICH only on an existing market export folder (needs prior CSV/JSON). "
        "Does NOT skip Phase 2 discovery - use --resume-checkpoint for that.",
    )
    parser.add_argument(
        "--resume-checkpoint",
        action="store_true",
        help="Skip Phase 1/2 discovery - reuse saved phase1 plan + phase2 company list, "
        "continue Phase 3 enrich / classify / export. Use after discovery finished / Ctrl+C.",
    )
    parser.add_argument(
        "--market-dir",
        default=None,
        help="With --resume: explicit market folder path (overrides slug from --query).",
    )
    parser.add_argument(
        "--no-geo-filter",
        action="store_true",
        help="Keep global players that operate in the target geography (don't drop foreign-HQ majors). "
        "Use for country markets that include international suppliers (e.g. 'Brazil X Market').",
    )
    parser.add_argument(
        "--shuffle",
        action="store_true",
        help="Randomize company order within each section (no quality/seed clustering).",
    )
    parser.add_argument(
        "--fmcg-xlsx",
        action="store_true",
        help="Also export 30-column FMCG landscape Excel (.xlsx) alongside CSV.",
    )
    args = parser.parse_args()
    cli_sections = _parse_sections(args.sections)
    cli_seeds = _parse_sections(args.seeds) + [n for n, _dom, _sec in _load_seeds_file(args.seeds_file)]
    # de-dupe preserving order
    cli_seeds = list(dict.fromkeys(cli_seeds))

    from vendor_intel.placeholders.load_keys import apply_env_overrides
    from vendor_intel.utils.output_filter import install_stderr_filter

    install_stderr_filter()
    apply_env_overrides()
    _refresh_output_dir()

    for w in _validate_env():
        print(f"  WARNING: {w}")

    from vendor_intel.pipeline.pipeline_modes import mode_describe, resolve_batch_options, apply_mode_settings

    profile_arg = args.profile or os.getenv("PIPELINE_PROFILE", "quality").strip()
    batch_opts = resolve_batch_options(
        args.mode,
        max_concurrent=None,
        cap=args.cap,
        profile=profile_arg,
        fmcg_xlsx=True if args.fmcg_xlsx else None,
        full_enrich=False,
        with_audit=False,
        gap_fill=False,
        push_neon=args.push_neon,
        search_backfill=False,
    )
    profile = batch_opts["profile"]
    fmcg_xlsx = args.fmcg_xlsx or batch_opts.get("fmcg_xlsx", False)
    settings = _load_settings(profile)
    if args.mode:
        settings = apply_mode_settings(settings, args.mode, cap_override=batch_opts["cap"])
        print(f"  Mode: {args.mode} — {mode_describe(args.mode)}")
    elif args.cap:
        from vendor_intel.pipeline.cap_profiles import apply_cap, cap_describe

        settings = apply_cap(settings, args.cap)
        print(f"  Cap: {cap_describe(args.cap)}")
    if args.no_geo_filter:
        settings = settings.model_copy(update={"pipeline_strict_geo": False})
        print("  Geo filter: OFF - keeping global players active in the target geography")
    if args.shuffle:
        settings = settings.model_copy(update={"pipeline_shuffle_export": True})
        print("  Order: randomized within sections")
    from vendor_intel.pipeline.geo_limits import pipeline_limits

    lim = pipeline_limits(settings, recall=(profile == "recall"), country="global")
    enrich_cap = int(lim["enrich"])
    classify_cap = int(lim["discover"])
    print(f"  Profile: {profile} | discover cap: {classify_cap} | enrich cap: {enrich_cap}")
    try:
        from vendor_intel.integrations.agent_reach_client import log_agent_reach_status

        log_agent_reach_status()
    except Exception:
        pass
    def _env_true(name: str) -> bool:
        return (os.getenv(name) or "").strip().lower() in ("1", "true", "yes", "on")

    post_enrich = not args.no_post_enrich
    post_market_kwargs = batch_opts.get("post_market")
    if args.mode in ("bulk", "standard"):
        post_enrich = False
    if args.mode in ("quality", "full") and not args.no_post_enrich:
        post_enrich = True
    skip_notes: list[str] = []
    if args.no_name_resolve or _env_true("SKIP_NAME_RESOLVE"):
        post_market_kwargs = dict(post_market_kwargs or {})
        post_market_kwargs["resolve_partials"] = False
        skip_notes.append("name_resolve")
    if args.no_linkedin_verify or _env_true("SKIP_LINKEDIN_FIRMOGRAPHIC_VERIFY"):
        post_market_kwargs = dict(post_market_kwargs or {})
        post_market_kwargs["linkedin_firmographic_verify"] = False
        skip_notes.append("linkedin_verify")
    if args.no_wikidata_firmographic or _env_true("SKIP_WIKIDATA_FIRMOGRAPHIC"):
        post_market_kwargs = dict(post_market_kwargs or {})
        post_market_kwargs["wikidata_firmographic"] = False
        skip_notes.append("wikidata")
    if args.no_wikipedia_firmographic or _env_true("SKIP_WIKIPEDIA_FIRMOGRAPHIC"):
        post_market_kwargs = dict(post_market_kwargs or {})
        post_market_kwargs["wikipedia_firmographic"] = False
        skip_notes.append("wikipedia")
    if args.no_push_neon or _env_true("SKIP_NEON_PUSH"):
        post_market_kwargs = dict(post_market_kwargs or {})
        post_market_kwargs["push_neon"] = False
        skip_notes.append("neon")
    if skip_notes:
        print(f"  Post-enrich skips: {', '.join(skip_notes)}")
    if post_enrich and post_market_kwargs:
        print("  Post-enrich: ON (mode preset — gap-fill → LinkedIn verify → Apollo → audit → Excel → Neon)")
    elif post_enrich:
        print("  Post-enrich: ON (gap-fill → LinkedIn verify → Apollo → contacts → brand → audit → Excel → Neon)")
    else:
        print("  Post-enrich: OFF")

    try:
        if args.resume:
            from vendor_intel.pipeline.run_and_export import slugify
            from vendor_intel.pipeline.post_market import run_post_market_pipeline_sync

            if args.market_dir:
                market_dir = Path(args.market_dir)
            elif args.query:
                query, country, _ = _parse_input(args.query)
                if args.country and args.country != "global":
                    country = args.country
                slug = slugify(query, country)
                market_dir = OUTPUT_DIR / slug
            else:
                print("ERROR: --resume needs --query or --market-dir")
                return
            if not market_dir.is_dir():
                print(f"ERROR: market folder not found: {market_dir}")
                print("  Run discovery first, or pass --market-dir path")
                return
            print(f"  RESUME post-enrich only: {market_dir}")
            print("  (skips finished steps + already-processed companies)")
            do_push = bool(args.push_neon) or (
                not args.no_push_neon and bool(os.getenv("DATABASE_URL", "").strip())
            )
            kwargs = dict(post_market_kwargs or {})
            kwargs["resume"] = True
            kwargs.setdefault("push_neon", do_push)
            _print_post_enrich(run_post_market_pipeline_sync(market_dir, **kwargs))
        elif args.file:
            batch_mode(
                args.file, settings, enrich_cap, classify_cap,
                sections=cli_sections, seeds=cli_seeds,
                push_neon=args.push_neon, no_push_neon=args.no_push_neon,
                post_enrich=post_enrich,
                post_market_kwargs=post_market_kwargs,
                fmcg_xlsx=fmcg_xlsx,
            )
        elif args.query:
            query, country, q_sections = _parse_input(args.query)
            if args.country and args.country != "global":
                country = args.country
            final_sections = q_sections or cli_sections
            t0 = time.perf_counter()
            result = run_one_query(
                query, country, settings,
                enrich_limit=enrich_cap,
                classify_limit=classify_cap,
                sections=final_sections,
                seeds=cli_seeds,
                out_name=args.out,
                fmcg_xlsx=fmcg_xlsx,
                resume_checkpoint=bool(args.resume_checkpoint),
            )
            _print_result(query, country, result, round(time.perf_counter() - t0, 1))
            if post_enrich:
                pm = dict(post_market_kwargs or {})
                pm.setdefault("resume", True)
                _print_post_enrich(
                    _run_post_enrich(
                        result,
                        push_neon=args.push_neon,
                        no_push_neon=args.no_push_neon,
                        post_market_kwargs=pm,
                    )
                )
            else:
                _maybe_push_to_neon(result, push_neon=args.push_neon, no_push_neon=args.no_push_neon)
        else:
            interactive_mode(
                settings, enrich_cap, classify_cap,
                sections=cli_sections, seeds=cli_seeds,
                push_neon=args.push_neon, no_push_neon=args.no_push_neon,
                post_enrich=post_enrich,
                post_market_kwargs=post_market_kwargs,
            )
    finally:
        try:
            from vendor_intel.clients.ddg_worker_pool import shutdown_ddg_pool
            shutdown_ddg_pool(wait=True)
        except Exception:
            pass


if __name__ == "__main__":
    main()
