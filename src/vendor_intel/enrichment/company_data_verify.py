"""OpenAI + web evidence: verify market fit and stored company fields."""
from __future__ import annotations

import asyncio
import json
import os
import time
from pathlib import Path
from typing import Any

from vendor_intel.enrichment.market_audit import gather_company_evidence


def _crawl(row: dict[str, Any]) -> dict[str, Any]:
    c = row.get("_crawl")
    return c if isinstance(c, dict) else {}


def _str(v: object) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    if s.lower() in ("not publicly disclosed", "none", "null", "n/a", "-"):
        return ""
    return s


def company_profile_from_row(row: dict[str, Any]) -> dict[str, str]:
    crawl = _crawl(row)
    return {
        "name": _str(row.get("company") or row.get("brand")),
        "brand": _str(row.get("brand")),
        "website": _str(row.get("website") or row.get("domain")),
        "role": _str(row.get("role")),
        "key_products": _str(row.get("key_products")),
        "summary": _str(row.get("company_summary") or row.get("summary"))[:400],
        "founded": _str(crawl.get("founded_year") or row.get("founded_year")),
        "headquarters": _str(crawl.get("headquarters") or crawl.get("contact_address")),
        "offices": _str(crawl.get("offices") or crawl.get("manufacturing_locations")),
        "employees": _str(crawl.get("employee_count")),
        "revenue": _str(crawl.get("revenue")),
        "email": _str(crawl.get("email")),
        "phone": _str(crawl.get("phone")),
        "contact_person": _str(crawl.get("contact_person")),
    }


def _local_evidence(row: dict[str, Any]) -> str:
    crawl = _crawl(row)
    parts: list[str] = []
    for key in (
        "crawl_text",
        "headquarters",
        "contact_address",
        "founded_year",
        "employee_count",
        "revenue",
        "offices",
        "products",
        "email",
        "phone",
        "contact_person",
        "contact_role",
    ):
        val = _str(crawl.get(key))
        if val:
            parts.append(f"{key}: {val[:800]}")
    raw = row.get("_crawl_raw")
    if isinstance(raw, dict):
        data = raw.get("data")
        if isinstance(data, dict):
            parts.append(f"crawl_raw_snippet: {json.dumps(data, default=str)[:1200]}")
    return "\n".join(parts)[:4000]


async def verify_one_company(
    row: dict[str, Any],
    *,
    market_query: str,
    market_desc: str,
    client: Any,
    model: str,
    sem: asyncio.Semaphore,
    use_web_search: bool = True,
) -> dict[str, Any]:
    async with sem:
        profile = company_profile_from_row(row)
        name = profile["name"] or "Unknown"
        local = _local_evidence(row)
        web = ""
        if use_web_search:
            web = await gather_company_evidence(
                {
                    "name": name,
                    "website": profile["website"],
                    "role": profile["role"],
                    "key_products": profile["key_products"],
                    "summary": profile["summary"],
                    "headquarters": profile["headquarters"],
                }
            )

        prompt = f"""You are a strict data-quality auditor for a market intelligence database.

MARKET: {market_query}
MARKET SCOPE: {market_desc}

STORED COMPANY RECORD:
{json.dumps(profile, indent=2)}

LOCAL CRAWL / STORED TEXT (from our pipeline — treat as candidate evidence, may be wrong):
{local or "(none)"}

WEB SEARCH SNIPPETS (external evidence):
{web or "(none — rely on local crawl only)"}

Tasks:
1. Does this company genuinely operate in this market scope (products/services match)?
2. Is the brand name consistent with the website/domain?
3. For each stored field below, say if web/local evidence SUPPORTS it, CONTRADICTS it, or is UNCLEAR.
   Only mark SUPPORTED if evidence explicitly mentions the value (do not guess).

Return JSON:
{{
  "company": "{name}",
  "market_relevant": "yes|partial|no|unclear",
  "brand_product_match": "yes|partial|no|unclear",
  "overall_verdict": "verified|review|reject",
  "fields": {{
    "founded": "supported|contradicted|unclear|empty",
    "headquarters": "supported|contradicted|unclear|empty",
    "key_products": "supported|contradicted|unclear|empty",
    "employees": "supported|contradicted|unclear|empty",
    "offices": "supported|contradicted|unclear|empty"
  }},
  "what_they_do": "1-2 sentences from best evidence",
  "issues": "comma-separated list or none",
  "confidence": "high|medium|low",
  "reason": "brief justification citing evidence"
}}"""

        for attempt in range(3):
            try:
                resp = await client.chat.completions.create(
                    model=model,
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0,
                    max_tokens=500,
                    response_format={"type": "json_object"},
                )
                result = json.loads(resp.choices[0].message.content or "{}")
                result["company"] = name
                result["website"] = profile["website"]
                return result
            except Exception as exc:
                if attempt < 2:
                    await asyncio.sleep(2**attempt)
                else:
                    return {
                        "company": name,
                        "market_relevant": "unclear",
                        "brand_product_match": "unclear",
                        "overall_verdict": "review",
                        "fields": {},
                        "what_they_do": "",
                        "issues": f"verify_error: {exc}",
                        "confidence": "low",
                        "reason": "OpenAI verification failed",
                        "website": profile["website"],
                    }
        return {"company": name, "overall_verdict": "review", "reason": "unknown"}


def write_verify_report(
    report_path: Path,
    *,
    market_query: str,
    market_desc: str,
    results: list[dict[str, Any]],
) -> None:
    verified = [r for r in results if r.get("overall_verdict") == "verified"]
    review = [r for r in results if r.get("overall_verdict") == "review"]
    reject = [r for r in results if r.get("overall_verdict") == "reject"]

    lines = [
        f"# Data Verification Report — {market_query}\n\n",
        f"**Market scope:** {market_desc}\n\n",
        "| Verdict | Count |\n|---------|-------|\n",
        f"| Verified | {len(verified)} |\n",
        f"| Review | {len(review)} |\n",
        f"| Reject | {len(reject)} |\n",
        f"| **Total** | **{len(results)}** |\n\n",
        "## Companies\n\n",
        "| # | Company | Verdict | Market | Brand/Products | HQ | Founded | Issues |\n",
        "|---|---------|---------|--------|----------------|----|---------|--------|\n",
    ]
    for i, r in enumerate(results, 1):
        fields = r.get("fields") or {}
        lines.append(
            f"| {i} | **{r.get('company', '?')}** | {r.get('overall_verdict', '?')} | "
            f"{r.get('market_relevant', '?')} | {r.get('brand_product_match', '?')} | "
            f"{fields.get('headquarters', '-')} | {fields.get('founded', '-')} | "
            f"{str(r.get('issues', 'none'))[:50]} |\n"
        )
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text("".join(lines), encoding="utf-8")


async def verify_companies(
    rows: list[dict[str, Any]],
    *,
    market_query: str,
    market_desc: str,
    output_dir: Path,
    concurrent: int = 3,
    use_web_search: bool = True,
    force: bool = False,
) -> dict[str, Any]:
    from backend.config import get_settings, sync_opencode_to_openai_env
    from openai import AsyncOpenAI

    sync_opencode_to_openai_env()
    cfg = get_settings()
    if not cfg.openai_api_key:
        return {"skipped": True, "reason": "OPENAI_API_KEY not set"}

    output_dir = Path(output_dir)
    cache_path = output_dir / "verify_cache.json"
    report_path = output_dir / "verify_report.md"

    cached: dict[str, dict] = {}
    if cache_path.is_file() and not force:
        try:
            cached_list = json.loads(cache_path.read_text(encoding="utf-8"))
            cached = {r["company"].lower(): r for r in cached_list if r.get("company")}
        except Exception:
            cached = {}

    todo = [r for r in rows if company_profile_from_row(r)["name"].lower() not in cached]
    if todo:
        client = AsyncOpenAI(api_key=cfg.openai_api_key)
        model = (os.getenv("OPENAI_MODEL") or cfg.model_extract or "gpt-4o-mini").strip()
        sem = asyncio.Semaphore(max(1, concurrent))
        t0 = time.perf_counter()
        total = len(todo)
        print(f"  [verify] Auditing {total} companies (OpenAI + web)...", flush=True)
        print(
            f"  [verify] First line may take 30–90s (web search + OpenAI). "
            f"Progress prints after each company.",
            flush=True,
        )

        lock = asyncio.Lock()
        done = 0

        def _persist_cache() -> None:
            cache_path.write_text(
                json.dumps(list(cached.values()), indent=2, ensure_ascii=False),
                encoding="utf-8",
            )

        async def _run_one(row: dict[str, Any]) -> dict[str, Any]:
            nonlocal done
            name = company_profile_from_row(row)["name"] or "Unknown"
            print(f"  [verify] ... {name} (web + OpenAI)", flush=True)
            result = await verify_one_company(
                row,
                market_query=market_query,
                market_desc=market_desc,
                client=client,
                model=model,
                sem=sem,
                use_web_search=use_web_search,
            )
            verdict = result.get("overall_verdict", "?")
            issues = str(result.get("issues") or "")
            err = " [API error]" if "verify_error" in issues else ""
            async with lock:
                done += 1
                if result.get("company"):
                    cached[result["company"].lower()] = result
                _persist_cache()
                partial: list[dict[str, Any]] = []
                for row in rows:
                    key = company_profile_from_row(row)["name"].lower()
                    if key in cached:
                        partial.append(cached[key])
                write_verify_report(
                    report_path,
                    market_query=market_query,
                    market_desc=market_desc,
                    results=partial,
                )
            print(
                f"  [verify] [{done}/{total}] {name} -> {verdict}{err} ({time.perf_counter() - t0:.0f}s)",
                flush=True,
            )
            return result

        new_results = await asyncio.gather(*[_run_one(r) for r in todo])
        print(f"  [verify] Done in {time.perf_counter() - t0:.1f}s", flush=True)
        for r in new_results:
            if r.get("company"):
                cached[r["company"].lower()] = r
        _persist_cache()

    all_results = []
    for r in rows:
        name = company_profile_from_row(r)["name"].lower()
        if name in cached:
            all_results.append(cached[name])

    write_verify_report(report_path, market_query=market_query, market_desc=market_desc, results=all_results)

    verified = sum(1 for r in all_results if r.get("overall_verdict") == "verified")
    review = sum(1 for r in all_results if r.get("overall_verdict") == "review")
    reject = sum(1 for r in all_results if r.get("overall_verdict") == "reject")

    stats = {
        "total": len(all_results),
        "verified": verified,
        "review": review,
        "reject": reject,
        "verify_cache_path": str(cache_path),
        "verify_report_path": str(report_path),
    }
    print(
        f"  [verify] Results: {verified} verified | {review} review | {reject} reject",
        flush=True,
    )
    print(f"  [verify] Report: {report_path}", flush=True)
    return stats
