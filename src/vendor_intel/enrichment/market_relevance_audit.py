"""Post-export market relevance audit — web search + OpenAI keep/review/remove.

Runs on Founded+HQ companies before Excel/Neon export. Sets is_relevant=false on
clear mismatches; flags partial matches for review (excluded from Neon push).
"""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

from vendor_intel.enrichment.gap_fill.gaps import norm
from vendor_intel.enrichment.gap_fill.sources import _search_snippets


def _company_key(row: dict[str, Any]) -> str:
    return norm(row.get("company") or row.get("brand")).lower()


def _load_cache(path: Path) -> dict[str, dict]:
    if not path.is_file():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _save_cache(path: Path, cache: dict[str, dict]) -> None:
    path.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")


async def _gather_evidence(
    router: Any,
    name: str,
    website: str,
    geo: str,
) -> str:
    queries = [
        f'"{name}" company products services market',
        (f'"{name}" {website} official what we do').strip(),
    ]
    parts: list[str] = []
    for q in queries:
        try:
            snippets = await _search_snippets(router, q, max_hits=5)
        except Exception:
            snippets = ""
        if snippets:
            parts.append(f"Query: {q}\n{snippets}")
    return "\n\n".join(parts)[:6000]


async def _llm_audit_verdict(
    row: dict[str, Any],
    *,
    market_label: str,
    market_def: str,
    include_kw: list[str],
    exclude_kw: list[str],
    evidence: str,
) -> dict[str, Any]:
    from backend.config import get_settings, sync_opencode_to_openai_env
    from openai import AsyncOpenAI

    sync_opencode_to_openai_env()
    cfg = get_settings()
    if not cfg.openai_api_key:
        return {"action": "keep", "reason": "OPENAI_API_KEY not set", "skipped": True}

    name = norm(row.get("company") or row.get("brand"))
    profile = (
        f"Company: {name}\n"
        f"Website: {row.get('website') or row.get('domain') or ''}\n"
        f"Role: {row.get('role') or ''}\n"
        f"Products: {str(row.get('key_products') or '')[:200]}\n"
        f"Summary: {str(row.get('company_summary') or row.get('summary') or '')[:300]}\n"
        f"Stored confidence: {row.get('confidence', '')}"
    )

    client = AsyncOpenAI(api_key=cfg.openai_api_key)
    model = (os.getenv("OPENAI_MODEL") or cfg.model_extract or "gpt-4o-mini").strip()

    inc = ", ".join(include_kw[:10]) or "n/a"
    exc = ", ".join(exclude_kw[:8]) or "n/a"

    prompt = f"""You are a strict market research auditor.

MARKET: {market_label}
MARKET DEFINITION: {market_def or market_label}
INCLUDE KEYWORDS: {inc}
EXCLUDE KEYWORDS: {exc}

STORED COMPANY RECORD:
{profile}

WEB SEARCH EVIDENCE:
{evidence or "(none — be conservative; use review if uncertain)"}

Decide if this company genuinely operates in this market (products/services match).

Return JSON:
{{
  "action": "keep|review|remove",
  "is_relevant": true/false,
  "relevance_score": 0-5,
  "market_relevant": "yes|partial|no|unclear",
  "reason": "brief justification citing evidence",
  "confidence": "high|medium|low"
}}

Rules:
- "keep" = clear evidence they participate in this market
- "review" = partial match or insufficient evidence
- "remove" = wrong industry, directory, or clearly not in this market
- is_relevant=false when action is remove
- relevance_score 0-2 → usually remove; 3 → review; 4-5 → keep
- Use web evidence over stored profile when they conflict
"""

    try:
        resp = await client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"},
            temperature=0,
        )
        text = (resp.choices[0].message.content or "").strip()
        data = json.loads(text) if text else {}
        if not isinstance(data, dict):
            return {"action": "keep", "reason": "invalid_llm_response"}
        action = str(data.get("action") or "keep").lower()
        if action not in ("keep", "review", "remove"):
            action = "keep"
        score = int(data.get("relevance_score") or 0)
        if score <= 2 and action == "keep":
            action = "remove" if score <= 1 else "review"
        if action == "remove":
            data["is_relevant"] = False
        elif action == "review":
            data["is_relevant"] = True
        else:
            data["is_relevant"] = bool(data.get("is_relevant", True))
        data["action"] = action
        return data
    except Exception as exc:
        return {"action": "keep", "reason": f"audit_error: {exc}"}


def apply_audit_verdict(row: dict[str, Any], verdict: dict[str, Any]) -> str:
    """Apply audit result to row. Returns action applied."""
    action = str(verdict.get("action") or "keep").lower()
    reason = norm(verdict.get("reason"))[:500]
    score = verdict.get("relevance_score")

    row["_market_audit"] = {
        "action": action,
        "reason": reason,
        "relevance_score": score,
        "market_relevant": verdict.get("market_relevant"),
        "confidence": verdict.get("confidence"),
    }

    if score is not None:
        try:
            row["audit_score"] = int(score)
            row["confidence"] = round(min(1.0, max(0.0, int(score) / 5.0)), 2)
        except (TypeError, ValueError):
            pass

    if action == "remove":
        row["is_relevant"] = False
        row["audit_reasoning"] = reason
        row["reject_reason"] = f"post_market_audit:{reason[:120]}"
        row.pop("_audit_review", None)
        return "remove"

    if action == "review":
        row["is_relevant"] = True
        row["_audit_review"] = True
        row["audit_reasoning"] = reason
        return "review"

    row["is_relevant"] = bool(verdict.get("is_relevant", True))
    row.pop("_audit_review", None)
    if reason:
        row["audit_reasoning"] = reason
    row["post_market_audited"] = True
    return "keep"


def _write_report(
    path: Path,
    *,
    market_label: str,
    results: list[dict[str, Any]],
) -> None:
    kept = sum(1 for r in results if r.get("action") == "keep")
    review = sum(1 for r in results if r.get("action") == "review")
    removed = sum(1 for r in results if r.get("action") == "remove")

    lines = [
        f"# Market Relevance Audit — {market_label}\n\n",
        f"| Verdict | Count |\n|---------|-------|\n",
        f"| Keep | {kept} |\n| Review | {review} |\n| Remove | {removed} |\n\n",
        "| Company | Action | Score | Reason |\n|---------|--------|-------|--------|\n",
    ]
    for r in results:
        lines.append(
            f"| {r.get('company', '?')} | {r.get('action', '?')} | "
            f"{r.get('relevance_score', '-')} | {str(r.get('reason', ''))[:80]} |\n"
        )
    path.write_text("".join(lines), encoding="utf-8")


async def audit_market_relevance_for_rows(
    rows: list[dict[str, Any]],
    *,
    query_context: dict[str, Any],
    scope: dict[str, Any] | None = None,
    market_dir: Path | None = None,
    concurrent: int = 3,
    force: bool = False,
) -> dict[str, Any]:
    import asyncio

    from vendor_intel.clients.search_router import FreeSearchRouter
    from vendor_intel.config import Settings
    from vendor_intel.pipeline.market_relevance import market_context_summary

    if not rows:
        return {"queued": 0, "kept": 0, "review": 0, "removed": 0}

    cache_path = (market_dir or Path(".")) / "market_relevance_audit_cache.json"
    report_path = (market_dir or Path(".")) / "market_relevance_audit_report.md"
    cache = _load_cache(cache_path)

    ctx = market_context_summary(scope, query_context)
    market_label = str(ctx.get("market") or query_context.get("industry") or "market")
    market_def = str(ctx.get("market_definition") or "")
    include_kw = list(ctx.get("include_keywords") or [])
    exclude_kw = list(ctx.get("exclude_keywords") or [])
    geo = str(query_context.get("country") or "")

    router = FreeSearchRouter(Settings.load())
    sem = asyncio.Semaphore(max(1, concurrent))

    todo: list[dict[str, Any]] = []
    for row in rows:
        # Only analyst seeds skip the audit; auto-enumerated market players and
        # discovered companies are audited for market relevance.
        if row.get("is_analyst_seed"):
            continue
        key = _company_key(row)
        if not force and key in cache and not cache[key].get("error"):
            cached_action = cache[key].get("action")
            if cached_action in ("keep", "review", "remove"):
                apply_audit_verdict(row, cache[key])
                continue
        todo.append(row)

    kept = removed = review = errors = 0
    report_rows: list[dict[str, Any]] = []
    cache_lock = asyncio.Lock()

    async def _one(row: dict[str, Any], idx: int) -> None:
        nonlocal kept, review, removed, errors
        async with sem:
            name = str(row.get("company") or row.get("brand") or "?")[:50]
            key = _company_key(row)
            website = str(row.get("website") or row.get("domain") or "")
            try:
                evidence = await _gather_evidence(router, name, website, geo)
                verdict = await _llm_audit_verdict(
                    row,
                    market_label=market_label,
                    market_def=market_def,
                    include_kw=include_kw,
                    exclude_kw=exclude_kw,
                    evidence=evidence,
                )
                if verdict.get("skipped"):
                    return
                action = apply_audit_verdict(row, verdict)
                async with cache_lock:
                    cache[key] = {
                        **verdict,
                        "company": name,
                        "action": action,
                    }
                    report_rows.append({**cache[key]})
                    # Per-company checkpoint so Ctrl+C mid-audit can resume.
                    _save_cache(cache_path, cache)
                if action == "keep":
                    kept += 1
                elif action == "review":
                    review += 1
                else:
                    removed += 1
                print(
                    f"  [market_audit] [{idx}/{len(todo)}] {name} -> {action} "
                    f"(score={verdict.get('relevance_score', '?')})",
                    flush=True,
                )
            except Exception as exc:
                errors += 1
                async with cache_lock:
                    cache[key] = {"error": str(exc), "action": "keep", "company": name}
                    _save_cache(cache_path, cache)
                print(f"  [market_audit] [{idx}/{len(todo)}] {name} ERROR: {exc}", flush=True)

    if todo:
        print(f"  [market_audit] Auditing {len(todo)} companies (web + OpenAI)...", flush=True)
        await asyncio.gather(*[_one(r, i) for i, r in enumerate(todo, 1)])

    _save_cache(cache_path, cache)

    all_report = []
    for row in rows:
        key = _company_key(row)
        if key in cache:
            all_report.append(cache[key])
    _write_report(report_path, market_label=market_label, results=all_report)

    for row in rows:
        audit = row.get("_market_audit") or {}
        act = audit.get("action")
        if act == "remove":
            removed += 1
        elif act == "review":
            review += 1
        elif act == "keep":
            kept += 1

    return {
        "queued": len(todo),
        "kept": kept,
        "review": review,
        "removed": removed,
        "errors": errors,
        "cache_path": str(cache_path),
        "report_path": str(report_path),
    }


def exportable_after_audit(row: dict[str, Any]) -> bool:
    """Founded+HQ row eligible for Neon push (relevant, not flagged review)."""
    if not row.get("is_relevant", True):
        return False
    if row.get("_audit_review"):
        return False
    return True
