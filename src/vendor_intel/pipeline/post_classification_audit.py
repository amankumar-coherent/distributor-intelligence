"""Post-classification LLM audit — keep/review/remove verdict per exported
company, generalized from final_audit_v2.py's hand-written per-market
MARKET_DESCRIPTIONS dict into a market description pulled from the pipeline's
own scope/market plan (pipeline.market_relevance.market_context_summary).

Runs on the FINAL export-ready company list (not every classified
candidate), matching what the original script actually audited (the CSV,
i.e. what already shipped) — this keeps cost bounded to what actually
exports instead of every raw candidate.

No new review queue: "remove" verdicts drop into export_rejected (the same
bucket/shape the rest of the pipeline already uses), "review" verdicts
demote into unverified_companies (the existing "could not confirm" bucket).
Both mechanisms already exist in pipeline/orchestrator.py for exactly this
purpose.
"""
from __future__ import annotations

import asyncio
from typing import Any


async def audit_final_companies(
    final: list[dict[str, Any]],
    query_context: dict[str, Any],
    scope: dict[str, Any] | None,
    settings: Any,
    *,
    router: Any = None,
    max_concurrent: int = 8,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    """Audit every row in `final`. Returns (kept, removed, reviewed):
    - kept: rows to leave in `final` unchanged
    - removed: rows to append to export_rejected (has "export_reject" set)
    - reviewed: rows in the unverified_companies shape (company/domain/role/reason)
    """
    if not final:
        return [], [], []

    from vendor_intel.clients.claude import ClaudeClient

    client = ClaudeClient(settings)
    if not client.available:
        return list(final), [], []

    from vendor_intel.clients.search_router import FreeSearchRouter
    from vendor_intel.pipeline.market_relevance import market_context_summary

    router = router or FreeSearchRouter(settings)
    market_ctx = market_context_summary(scope, query_context)
    market_label = str(market_ctx.get("market") or "this market")
    market_def = str(market_ctx.get("market_definition") or "")
    geo = str(query_context.get("country") or "")

    sem = asyncio.Semaphore(max(1, max_concurrent))

    async def _audit_one(row: dict[str, Any]) -> dict[str, Any]:
        name = str(row.get("company") or row.get("brand") or "")
        website = str(row.get("website") or row.get("domain") or "")
        async with sem:
            evidence = await _gather_evidence(router, name, website, geo)
            verdict = await _llm_verdict(client, name, row, market_label, market_def, evidence)
        return {**row, "_audit": verdict}

    audited = await asyncio.gather(*[_audit_one(r) for r in final], return_exceptions=True)

    kept: list[dict[str, Any]] = []
    removed: list[dict[str, Any]] = []
    reviewed: list[dict[str, Any]] = []
    for row, result in zip(final, audited):
        if isinstance(result, BaseException) or not isinstance(result, dict):
            kept.append(row)
            continue
        verdict = result.get("_audit") or {}
        action = str(verdict.get("action") or "keep").lower()
        reason = str(verdict.get("reason") or "")[:150]
        if action == "remove":
            out = dict(row)
            out["export_reject"] = f"post_audit:{reason or 'not confirmed'}"
            removed.append(out)
        elif action == "review":
            reviewed.append(
                {
                    "company": row.get("company") or row.get("brand") or "",
                    "domain": str(row.get("domain") or row.get("website") or ""),
                    "role": row.get("role") or "",
                    "reason": f"post-audit flagged for review: {reason}" if reason else "post-audit flagged for review",
                }
            )
        else:
            kept.append(row)

    if removed or reviewed:
        print(
            f"  [post_audit] {len(kept)} keep, {len(reviewed)} review, "
            f"{len(removed)} remove (of {len(final)} final)",
            flush=True,
        )
    return kept, removed, reviewed


async def _gather_evidence(router: Any, name: str, website: str, geo: str) -> str:
    queries = [
        f'"{name}" company what do they do products services',
        (f'"{name}" {website} official').strip(),
    ]
    snippets: list[str] = []
    for q in queries:
        try:
            results = await router.search(q, geo=geo, validation_mode=True)
        except Exception:
            results = []
        for r in results[:4]:
            snippets.append(f"[{r.title}] {r.link}\n{r.snippet[:200]}")
    return "\n\n".join(snippets)[:3000]


async def _llm_verdict(
    client: Any, name: str, row: dict[str, Any], market_label: str, market_def: str, evidence: str
) -> dict[str, Any]:
    profile = (
        f"Company: {name}\n"
        f"Website: {row.get('website') or row.get('domain') or ''}\n"
        f"Role: {row.get('role') or ''}\n"
        f"Summary: {str(row.get('company_summary') or row.get('summary') or '')[:300]}"
    )
    system = (
        "You are a strict market research auditor. Verify if this company truly "
        "operates in the stated market. Respond with JSON only."
    )
    user = f"""MARKET: {market_label}
MARKET DEFINITION: {market_def}

COMPANY PROFILE (from our database):
{profile}

WEB EVIDENCE (from live search):
{evidence}

STRICT RULES:
- "keep" = clear evidence company operates in this market
- "review" = partial match or insufficient evidence — needs manual check
- "remove" = wrong industry, not a real company, or clearly irrelevant
- Use web evidence to override the profile if they contradict

Return JSON: {{"action": "keep|review|remove", "reason": "brief justification"}}"""
    try:
        raw = await asyncio.to_thread(client.complete_json, system, user, max_tokens=200)
        if isinstance(raw, dict) and raw.get("action") in ("keep", "review", "remove"):
            return raw
    except Exception:
        pass
    return {"action": "keep", "reason": "audit_unavailable"}
