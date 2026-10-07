#!/usr/bin/env python3
"""One-command market expand: ChatGPT (OpenAI SDK) OR Cursor SDK OR harvest-only.

Examples (from D:\\vendor-intel):

  # Single market
  .\\.venv\\Scripts\\python.exe run_cursor_expand.py --backend chatgpt --query "Global Avocado Oil Market" --batch all

  # 1…10k markets (same steps 0–7 each; skip-done + resume):
  .\\.venv\\Scripts\\python.exe run_chatgpt_expand_batch.py --input queries/neon_markets_all.csv --batch-id expand-10k --yes
"""
from __future__ import annotations

import argparse
import asyncio
import json
import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT / "src") not in sys.path:
    sys.path.insert(0, str(ROOT / "src"))


def _load_env() -> None:
    try:
        from dotenv import dotenv_values, load_dotenv
    except ImportError:
        return
    load_dotenv(ROOT / ".env")
    cmi_env = Path(r"C:\Users\aman.kumar\Desktop\Cmi\agents\vendor-intel\.env")
    if cmi_env.exists():
        # Prefer Cmi LLM keys (DeepSeek) over stale D:\ OpenAI keys
        vals = dotenv_values(cmi_env)
        for k, v in (vals or {}).items():
            if not v:
                continue
            if k.startswith(
                ("OPENAI_", "LLM_", "COMPILER_", "CLASSIFIER_", "MARKET_MAP_")
            ):
                os.environ[k] = v
        load_dotenv(cmi_env, override=False)

    # os.environ is now populated from .env, but several modules (notably
    # the ChatGPT web-UI scraper) cache config into module-level globals at
    # import time and only re-read os.environ when explicitly told to via
    # apply_env_overrides() — without this call CHATGPT_SCRAPER_ENABLED (and
    # the selenium/LLM key overrides) silently stay at their hardcoded
    # defaults no matter what .env says.
    try:
        from vendor_intel.placeholders.load_keys import apply_env_overrides

        apply_env_overrides()
    except Exception:
        pass


def _resolve_final(query: str, final: str | None) -> str:
    """Generic per-market FINAL under output/web_expand/<slug>/ (no demo hardcoding)."""
    if final:
        return str(Path(final).resolve())
    from vendor_intel.pipeline.web_expand import default_output_dir, resolve_final_path

    out = default_output_dir(query, "global")
    return str(resolve_final_path(query, "global", out, None))


def build_prompt(
    *,
    query: str,
    country: str,
    batch: str,
    target: int,
    final_path: str,
) -> str:
    return f"""You are running as a Cursor SDK local agent inside the vendor-intel repo.

TASK: Web expand own — market research fill for this market.

Market query: {query}
Country / geo: {country}
Batch: {batch}  (a = first half toward target, b = remaining, all = full)
Target companies: {target}
FINAL Excel path (merge/update this file): {final_path}
Repo root (cwd): {ROOT}

MANDATORY WORKFLOW:
1. Load existing FINAL Excel if present; count companies; keep good rows.
2. Do LIVE web search yourself (WebSearch / WebFetch / or project search tools).
3. Extract real channel companies; filter OEMs/junk/duplicates.
4. Progress toward {target} for this batch.
5. Fill ALL 35 landscape columns. Contacts: public only or "Not publicly disclosed".
6. Specialty Grocers + Foodservice / HoReCa = "Not applicable" for IT/wearables/green chem.
7. Save wrap-formatted Excel + audit JSON. Report counts at the end.

Do the work now. Prefer half-batches. Do not ask questions unless blocked.
"""


def _run_harvest(args: argparse.Namespace) -> dict:
    from vendor_intel.pipeline.web_expand import run_web_expand

    return asyncio.run(
        run_web_expand(
            args.query,
            country=args.country,
            target=args.target,
            batch=args.batch,
            max_queries=args.max_queries,
            merge_existing=not args.no_merge,
            final_path=args.final,
            out_dir=args.out_dir,
        )
    )


def _run_openai_pipeline(args: argparse.Namespace) -> int:
    """Deep enrich via existing run_query.py (uses OPENAI_API_KEY / LLM_PROVIDER)."""
    cmd = [
        sys.executable,
        str(ROOT / "run_query.py"),
        "--query",
        args.query,
        "--country",
        args.country,
        "--mode",
        "quality",
        "--cap",
        args.cap,
        "--fmcg-xlsx",
        "--profile",
        "recall",
        "--no-push-neon",
    ]
    print("  [openai] launching pipeline:", " ".join(cmd), flush=True)
    return subprocess.call(cmd, cwd=str(ROOT))


def run_chatgpt_backend(args: argparse.Namespace) -> int:
    """Web search + ChatGPT/DeepSeek (OpenAI SDK) extract + fill all columns."""
    openai_key = (os.environ.get("OPENAI_API_KEY") or "").strip()
    if not openai_key:
        print(
            "ERROR: OPENAI_API_KEY not set in .env.\n"
            "  Add your DeepSeek (or OpenAI) key as OPENAI_API_KEY in\n"
            "  D:\\vendor-intel\\.env or Cmi\\agents\\vendor-intel\\.env\n"
            "  With DeepSeek also set:\n"
            "    OPENAI_BASE_URL=https://api.deepseek.com/v1\n"
            "    OPENAI_MODEL=deepseek-chat",
            file=sys.stderr,
        )
        return 1

    from vendor_intel.pipeline.chatgpt_expand import run_chatgpt_expand, _is_deepseek, _llm_label

    provider = _llm_label()
    print(f"=== {provider} SDK expand (NO run_query.py) ===", flush=True)
    if _is_deepseek():
        print(
            "  DeepSeek mode: search-stack fill + ddgs harvest + chat residual "
            "(no OpenAI hosted web_search)",
            flush=True,
        )
    if getattr(args, "seeds_only", False):
        print(
            f"  steps: curated seeds → {provider} verify → search-stack fill → final verify",
            flush=True,
        )
    else:
        print(
            f"  steps: {provider} find+verify → search-stack fill "
            f"(Google AI → DDGS/SearXNG) → {provider} residual gaps → final verify",
            flush=True,
        )
    print(f"  query  : {args.query}", flush=True)
    print(f"  batch  : {args.batch}", flush=True)
    print(f"  target : {args.target}", flush=True)
    print(f"  model  : {os.environ.get('OPENAI_MODEL', 'gpt-4o-mini')}", flush=True)
    print(f"  base   : {os.environ.get('OPENAI_BASE_URL', 'https://api.openai.com/v1')}", flush=True)
    if getattr(args, "seeds_only", False):
        print("  mode   : seeds-only (uses queries/seeds/*.txt)", flush=True)
    if not getattr(args, "no_web_fill", False):
        fb = str(getattr(args, "fill_backend", "search") or "search").lower()
        if fb in ("search", "searx", "ddgs", "stack") or _is_deepseek():
            print(
                f"  fill   : Google AI → DDGS/SearXNG first, then {provider} chat for leftover NPD/gaps",
                flush=True,
            )
        else:
            print(
                "  fill   : OpenAI Responses API web_search per company",
                flush=True,
            )
    else:
        print(f"  fill   : {provider} knowledge only (--no-web-fill)", flush=True)
    if getattr(args, "fresh", False):
        print("  ckpt   : fresh start (ignore checkpoint)", flush=True)
    else:
        print("  ckpt   : auto-resume from checkpoint if present", flush=True)
    if str(getattr(args, "fill_backend", "search") or "search").lower() in (
        "search",
        "searx",
        "ddgs",
        "stack",
    ) or _is_deepseek():
        print(
            f"  fill   : Google AI + DDGS + SearXNG + Exa + Wikipedia + Owler, "
            f"then {provider} residual for remaining gaps (SEARCH_STACK_OPENAI_RESIDUAL)",
            flush=True,
        )
        if not getattr(args, "no_apollo", False):
            print("  apollo : contact person/role", flush=True)
    else:
        if not getattr(args, "no_linkedin_enrich", False):
            print(
                "  linkedin: OpenAI web_search only (site:linkedin.com — no MCP)",
                flush=True,
            )
        if not getattr(args, "no_apollo", False):
            print(
                "  apollo : contact person/role → full name via LinkedIn/web",
                flush=True,
            )

    if getattr(args, "enrich", False):
        print(
            f"  NOTE: --enrich ignored for chatgpt backend ({provider} SDK does all steps).",
            flush=True,
        )

    try:
        audit = asyncio.run(
            run_chatgpt_expand(
                args.query,
                country=args.country,
                target=args.target,
                batch=args.batch,
                max_queries=args.max_queries,
                merge_existing=not args.no_merge,
                final_path=args.final,
                out_dir=args.out_dir,
                fill_existing=args.fill_existing,
                max_fill=args.max_fill,
                seeds_only=bool(getattr(args, "seeds_only", False)),
                seeds_path=getattr(args, "seeds", None),
                skip_recall=bool(getattr(args, "skip_recall", False)),
                web_fill=not bool(getattr(args, "no_web_fill", False)),
                openai_discover=not bool(getattr(args, "no_openai_discover", False)),
                ddgs_harvest=bool(getattr(args, "ddgs_harvest", False)),
                resume=not bool(getattr(args, "fresh", False)),
                use_linkedin_enrich=not bool(getattr(args, "no_linkedin_enrich", False)),
                use_apollo=not bool(getattr(args, "no_apollo", False)),
                fill_backend=str(getattr(args, "fill_backend", "search") or "search"),
                force=bool(getattr(args, "fresh", False)),
            )
        )
    except Exception as err:  # noqa: BLE001
        print(f"ERROR: chatgpt expand failed: {type(err).__name__}: {err}", file=sys.stderr)
        return 1

    print(json.dumps(audit, indent=2), flush=True)

    total = int(audit.get("total_after") or 0)
    if total < args.target:
        print(
            f"\n  NOTE: {total}/{args.target} companies. "
            f"Run again with --batch b or raise --max-queries / --max-fill.",
            flush=True,
        )
    else:
        print(f"\n  OK: {total} >= {args.target}", flush=True)
    return 0


def run_kompass_backend(args: argparse.Namespace) -> int:
    """Global Kompass distributor scraper, with automatic ChatGPT/DeepSeek
    expand-pipeline fallback to reach a minimum company count.

    Step 1: Kompass — finds companies Kompass explicitly classifies as
    distributors, with market-specific evidence (not just a keyword match) —
    see vendor_intel.discovery.kompass_distributors. No upper limit by
    default (MAX_COMPANIES_PER_MARKET=0).
    Step 2: if the total is still below MIN_COMPANIES_PER_MARKET (default
    500) — including when Kompass finds the market page but not enough
    distributors, OR can't find/access Kompass for this market at all —
    fall back to the SAME steps as --backend chatgpt (DeepSeek/OpenAI SDK
    find+verify+fill, same as run_chatgpt_expand) to top up the remainder
    into the SAME final Excel. Existing Kompass rows are excluded by name so
    the fallback only discovers genuinely new companies.
    """
    from vendor_intel.discovery.kompass_distributors import (
        apply_final_distributor_gate,
        market_result_to_rows,
        min_companies_per_market,
        run_kompass_distributor_search,
    )
    from vendor_intel.pipeline.openai_expand_checkpoint import ExpandCheckpoint
    from vendor_intel.pipeline.web_expand import (
        read_final_rows,
        resolve_final_path,
        write_final_xlsx,
    )

    min_total = min_companies_per_market()
    final_path = resolve_final_path(args.query, args.country, args.out_dir and Path(args.out_dir), args.final)
    existing = [] if args.no_merge else read_final_rows(final_path)

    ckpt, resumed = ExpandCheckpoint.load_or_create(
        final_path.parent,
        query=args.query,
        country=args.country,
        batch="kompass",
        target=min_total,
        resume=not getattr(args, "fresh", False),
    )

    print(f"=== Kompass distributor scraper: {args.query} (min target: {min_total}) ===", flush=True)
    if resumed and ckpt.data("kompass_step_done"):
        print("  [kompass] checkpoint: Kompass discovery already completed — reusing saved result", flush=True)
        market_status = ckpt.data("kompass_result_market_status", "kompass_market_not_found")
        kompass_new_rows = list(ckpt.data("kompass_result_rows", []))
        kompass_audit = {"market_status": market_status}
        if market_status == "ok":
            kompass_audit.update(
                {
                    "kompass_product_id": ckpt.data("kompass_result_product_id"),
                    "kompass_product_url": ckpt.data("kompass_result_product_url"),
                    "verified_distributors": len(kompass_new_rows),
                    "rejected_candidates": ckpt.data("kompass_result_rejected_count", 0),
                }
            )
        else:
            print(f"  [kompass] market_status={market_status} — 0 companies from Kompass", flush=True)
    else:
        if resumed:
            print("  [kompass] checkpoint: resuming Kompass discovery from last saved progress...", flush=True)
        ckpt.bump("kompass_0_start", "begin", note="starting kompass distributor search")
        result = run_kompass_distributor_search(args.query, ckpt=ckpt)

        kompass_new_rows = []
        kompass_audit = {"market_status": result.market_status}
        if result.market_status == "ok":
            kompass_new_rows = market_result_to_rows(result)
            kompass_audit.update(
                {
                    "kompass_product_id": result.kompass_product_id,
                    "kompass_product_url": result.kompass_product_url,
                    "verified_distributors": len(kompass_new_rows),
                    "rejected_candidates": result.rejected_count,
                }
            )
        else:
            print(f"  [kompass] market_status={result.market_status} — 0 companies from Kompass", flush=True)

        ckpt.bump(
            "kompass_4_done", "search_complete",
            kompass_step_done=True,
            kompass_result_market_status=result.market_status,
            kompass_result_product_id=result.kompass_product_id,
            kompass_result_product_url=result.kompass_product_url,
            kompass_result_rejected_count=result.rejected_count,
            kompass_result_rows=kompass_new_rows,
            note="kompass distributor search complete",
        )

    existing_names = {row.get("Company", "").strip().lower() for row in existing}
    merged = list(existing) + [
        r for r in kompass_new_rows if r["Company"].strip().lower() not in existing_names
    ]

    write_final_xlsx(
        final_path,
        merged,
        sheet_name=args.query[:31],
        audit={"query": args.query, "backend": "kompass", **kompass_audit},
    )
    market_status = kompass_audit.get("market_status")
    print(
        f"  [kompass] {len(merged)} total rows after Kompass "
        f"({len(kompass_new_rows)} new verified distributors"
        + (f", {kompass_audit.get('rejected_candidates', 0)} rejected)" if market_status == "ok" else ")"),
        flush=True,
    )

    if len(merged) >= min_total:
        print(f"  [kompass] OK — {len(merged)} >= minimum {min_total}, no fallback needed", flush=True)
        _apply_gate_and_rewrite(final_path, merged, args.query, kompass_audit, used_fallback=False)
        ckpt.mark_done()
        return 0

    print(
        f"  [kompass] {len(merged)} < minimum {min_total} — falling back to ChatGPT/DeepSeek "
        f"expand pipeline (same steps as --backend chatgpt) to top up",
        flush=True,
    )
    ckpt.bump("kompass_5_fallback", "dispatching", note="dispatching chatgpt/deepseek fallback")
    fallback_args = argparse.Namespace(**vars(args))
    fallback_args.target = min_total
    fallback_args.no_merge = False  # must merge on top of the Kompass rows just written
    rc = run_chatgpt_backend(fallback_args)
    if rc != 0:
        print(
            f"  [kompass] fallback pipeline failed (exit {rc}) — "
            f"{len(merged)} Kompass-only rows remain saved at {final_path}",
            flush=True,
        )
        return rc

    # Final verification gate — applies to ALL rows now in the file,
    # regardless of whether Kompass or the ChatGPT/DeepSeek fallback
    # produced them (spec: pure distributors only, no manufacturer/OEM,
    # enforced deterministically at the end, not just by an LLM prompt).
    all_rows = read_final_rows(final_path)
    _apply_gate_and_rewrite(final_path, all_rows, args.query, kompass_audit, used_fallback=True)
    ckpt.mark_done()
    return 0


def _apply_gate_and_rewrite(
    final_path: Path,
    rows: list[dict[str, str]],
    query: str,
    audit: dict[str, object],
    *,
    used_fallback: bool,
) -> None:
    from vendor_intel.discovery.kompass_distributors import apply_final_distributor_gate
    from vendor_intel.pipeline.web_expand import write_final_xlsx
    from vendor_intel.pipeline.website_verification import verify_and_fix_website

    kept, rejected = apply_final_distributor_gate(rows)

    # Website reachability + identity verification — every row, regardless
    # of source (Kompass or the ChatGPT/DeepSeek fallback), gets its Website
    # column checked/fixed here, same as company_enrichment.py's per-company
    # path does for rows that go through that orchestrator instead.
    websites_fixed = 0
    for row in kept:
        name = str(row.get("Company") or "").strip()
        if not name:
            continue
        try:
            result = verify_and_fix_website(name, str(row.get("Headquarters") or ""), str(row.get("Website") or ""))
        except Exception as exc:  # noqa: BLE001
            print(f"    -> website verify error for {name!r}: {type(exc).__name__}: {exc}", flush=True)
            continue
        if result.was_corrected or (not result.website and row.get("Website")):
            row["Website"] = result.website
            websites_fixed += 1

    write_final_xlsx(
        final_path,
        kept,
        sheet_name=query[:31],
        audit={
            **audit,
            "backend": "kompass",
            "final_gate_rejected": len(rejected),
            "websites_fixed": websites_fixed,
        },
    )
    print(
        f"  [kompass] FINAL VERIFY: {len(kept)} distributor rows kept, "
        f"{len(rejected)} manufacturer/OEM rows rejected, "
        f"{websites_fixed} websites fixed/cleared -> {final_path}",
        flush=True,
    )
    print(
        json.dumps(
            {
                **audit,
                "total_after": len(kept),
                "used_fallback": used_fallback,
                "final_gate_rejected": len(rejected),
                "websites_fixed": websites_fixed,
            },
            indent=2,
        )
    )


def run_harvest_backend(args: argparse.Namespace) -> int:
    print("=== Web harvest expand (no ChatGPT) ===", flush=True)
    try:
        audit = _run_harvest(args)
    except Exception as err:  # noqa: BLE001
        print(f"ERROR: harvest failed: {type(err).__name__}: {err}", file=sys.stderr)
        return 1
    print(json.dumps(audit, indent=2), flush=True)
    return 0


def run_cursor_backend(args: argparse.Namespace) -> int:
    api_key = (os.environ.get("CURSOR_API_KEY") or "").strip()
    if not api_key:
        print(
            "ERROR: CURSOR_API_KEY not set.\n"
            "  Use OpenAI instead:\n"
            "    .\\.venv\\Scripts\\python.exe run_cursor_expand.py --backend openai "
            f'--query "{args.query}" --batch {args.batch}\n'
            "  Or add CURSOR_API_KEY=crsr_... to .env for --backend cursor.",
            file=sys.stderr,
        )
        return 1

    try:
        from cursor_sdk import (  # type: ignore
            Agent,
            AgentOptions,
            CursorAgentError,
            LocalAgentOptions,
        )
    except ImportError:
        print(
            "ERROR: cursor-sdk not installed.\n"
            "  .\\.venv\\Scripts\\python.exe -m pip install cursor-sdk\n"
            "  Or use: --backend openai",
            file=sys.stderr,
        )
        return 1

    final_path = _resolve_final(args.query, args.final)
    prompt = build_prompt(
        query=args.query,
        country=args.country,
        batch=args.batch,
        target=args.target,
        final_path=final_path,
    )

    print("=== Cursor SDK expand ===", flush=True)
    print(f"  query : {args.query}", flush=True)
    print(f"  batch : {args.batch}", flush=True)
    print(f"  target: {args.target}", flush=True)
    print(f"  model : {args.model}", flush=True)
    print(f"  final : {final_path}", flush=True)
    print("  Starting Cursor agent…\n", flush=True)

    options = AgentOptions(
        api_key=api_key,
        model=args.model,
        local=LocalAgentOptions(cwd=str(ROOT)),
    )

    try:
        if args.no_stream:
            result = Agent.prompt(prompt, options)
            status = getattr(result, "status", None)
            text = getattr(result, "result", None)
            usage = getattr(result, "usage", None)
            print("\n=== Agent finished ===", flush=True)
            print(f"  status: {status}", flush=True)
            if text:
                print(text, flush=True)
            if usage is not None:
                print(
                    f"  tokens: total={getattr(usage, 'total_tokens', None)} "
                    f"in={getattr(usage, 'input_tokens', None)} "
                    f"out={getattr(usage, 'output_tokens', None)}",
                    flush=True,
                )
            return 0 if status in (None, "finished") else 2

        with Agent.create(
            model=args.model,
            api_key=api_key,
            local=LocalAgentOptions(cwd=str(ROOT)),
        ) as agent:
            run = agent.send(prompt)
            for message in run.messages():
                mtype = getattr(message, "type", None)
                if mtype == "assistant":
                    content = getattr(message, "message", None)
                    blocks = getattr(content, "content", None) if content else None
                    if blocks is None and hasattr(message, "content"):
                        blocks = message.content
                    if blocks:
                        for block in blocks:
                            if getattr(block, "type", None) == "text":
                                print(getattr(block, "text", ""), end="", flush=True)
                elif mtype == "tool_call":
                    print(f"\n[tool] {getattr(message, 'name', 'tool')}", flush=True)
                elif mtype == "status":
                    print(f"\n[status] {getattr(message, 'status', '')}", flush=True)
            result = run.wait()
            print("\n\n=== Agent finished ===", flush=True)
            print(f"  status: {result.status}", flush=True)
            usage = getattr(result, "usage", None)
            if usage is not None:
                print(
                    f"  tokens: total={getattr(usage, 'total_tokens', None)} "
                    f"in={getattr(usage, 'input_tokens', None)} "
                    f"out={getattr(usage, 'output_tokens', None)}",
                    flush=True,
                )
            out_dir = Path(final_path).parent
            out_dir.mkdir(parents=True, exist_ok=True)
            (out_dir / f"cursor_expand_batch_{args.batch}_run.json").write_text(
                json.dumps(
                    {
                        "status": result.status,
                        "query": args.query,
                        "batch": args.batch,
                        "target": args.target,
                        "final": final_path,
                        "model": args.model,
                        "backend": "cursor",
                    },
                    indent=2,
                ),
                encoding="utf-8",
            )
            return 0 if result.status == "finished" else 2

    except CursorAgentError as err:
        print(
            f"ERROR: agent failed to start: {getattr(err, 'message', err)} "
            f"(retryable={getattr(err, 'is_retryable', None)})",
            file=sys.stderr,
        )
        return 1
    except Exception as err:  # noqa: BLE001
        print(f"ERROR: {type(err).__name__}: {err}", file=sys.stderr)
        return 1


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Market expand: --backend chatgpt (OPENAI_API_KEY) or cursor (CURSOR_API_KEY)"
    )
    parser.add_argument("--query", "-q", required=True, help='e.g. "Global Firewall Market"')
    parser.add_argument("--country", "-c", default="global")
    parser.add_argument("--batch", "-b", choices=("a", "b", "all"), default="a")
    parser.add_argument("--target", "-t", type=int, default=500)
    parser.add_argument("--final", default=None, help="Override FINAL xlsx path")
    parser.add_argument(
        "--backend",
        choices=("chatgpt", "openai", "cursor", "harvest", "kompass"),
        default="chatgpt",
        help="chatgpt/openai=full OpenAI/DeepSeek SDK find+verify+fill (default, NO run_query.py); "
        "DeepSeek uses search-stack + chat (no hosted web_search); "
        "cursor=Cursor SDK; harvest=web search only; "
        "kompass=global Kompass distributor scraper (market-verified distributors only)",
    )
    parser.add_argument(
        "--enrich",
        action="store_true",
        help="Ignored for chatgpt (kept for compatibility). ChatGPT backend never calls run_query.py.",
    )
    parser.add_argument(
        "--max-fill",
        type=int,
        default=500,
        help="Fill every table row up to this cap (default 500, not 20)",
    )
    parser.add_argument(
        "--fill-existing",
        action="store_true",
        help="Ask ChatGPT to re-fill existing FINAL rows (not only new ones)",
    )
    parser.add_argument(
        "--cap",
        default="maximum",
        choices=("bulk", "focused", "standard", "broad", "maximum"),
        help="Pipeline cap when --enrich (default: maximum)",
    )
    parser.add_argument(
        "--max-queries",
        type=int,
        default=120,
        help="Web search query count (use 120–180 for 300+ companies)",
    )
    parser.add_argument("--out-dir", default=None)
    parser.add_argument("--no-merge", action="store_true")
    parser.add_argument(
        "--seeds-only",
        action="store_true",
        help="Skip web harvest; load curated queries/seeds/*.txt then verify+fill (use for 300+ fill)",
    )
    parser.add_argument(
        "--seeds",
        default=None,
        help="Optional seed file path (default: auto-match queries/seeds/<market>.txt)",
    )
    parser.add_argument(
        "--skip-recall",
        action="store_true",
        help="Skip ChatGPT recall step (useful with --seeds-only when seed list is full)",
    )
    parser.add_argument(
        "--no-web-fill",
        action="store_true",
        help="Disable OpenAI web_search during fill (faster/cheaper; uses model knowledge only)",
    )
    parser.add_argument(
        "--no-openai-discover",
        action="store_true",
        help="Disable OpenAI web_search company discovery (Step 2)",
    )
    parser.add_argument(
        "--ddgs-harvest",
        action="store_true",
        help="Also run free ddgs/SearXNG harvest (optional; OpenAI discover is default)",
    )
    parser.add_argument(
        "--fresh",
        action="store_true",
        help="Ignore existing checkpoint and start a new run (default: auto-resume)",
    )
    parser.add_argument(
        "--resume",
        action="store_true",
        help="Resume from checkpoint if present (default behavior; flag kept for clarity)",
    )
    parser.add_argument(
        "--no-linkedin-enrich",
        action="store_true",
        help="Skip LinkedIn gap-fill via OpenAI web_search (site:linkedin.com)",
    )
    parser.add_argument(
        "--no-apollo",
        action="store_true",
        help="Skip Apollo contact person/role lookup + full-name resolve",
    )
    parser.add_argument(
        "--fill-backend",
        choices=("search", "openai", "enrichment_chain"),
        default="search",
        help="search=DDGS/SearXNG/Google AI scraper (default); "
        "openai=OpenAI web_search fill (costly); "
        "enrichment_chain=per-field Kompass → ChatGPT scraper → Google AI → "
        "DeepSeek → Smart Crawl fallback chain",
    )
    parser.add_argument(
        "--model",
        default=os.environ.get("CURSOR_EXPAND_MODEL", "composer-2.5"),
        help="Cursor model id (cursor backend only)",
    )
    parser.add_argument(
        "--no-stream",
        action="store_true",
        help="Cursor backend: Agent.prompt one-shot",
    )
    args = parser.parse_args()

    _load_env()
    os.chdir(ROOT)

    if args.backend in ("chatgpt", "openai"):
        return run_chatgpt_backend(args)
    if args.backend == "harvest":
        return run_harvest_backend(args)
    if args.backend == "kompass":
        return run_kompass_backend(args)
    return run_cursor_backend(args)


if __name__ == "__main__":
    raise SystemExit(main())
