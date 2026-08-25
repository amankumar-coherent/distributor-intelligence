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
        choices=("chatgpt", "openai", "cursor", "harvest"),
        default="chatgpt",
        help="chatgpt/openai=full OpenAI/DeepSeek SDK find+verify+fill (default, NO run_query.py); "
        "DeepSeek uses search-stack + chat (no hosted web_search); "
        "cursor=Cursor SDK; harvest=web search only",
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
        choices=("search", "openai"),
        default="search",
        help="search=DDGS/SearXNG/Google AI scraper (default); "
        "openai=OpenAI web_search fill (costly)",
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
    return run_cursor_backend(args)


if __name__ == "__main__":
    raise SystemExit(main())
