"""Inject Phase 1 LLM seed companies (priority, correct domains)."""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from vendor_intel.discovery.company_registry import normalize_seed_companies
from vendor_intel.funnel.scope_schema import normalize_run_scope


def load_scope_from_plan(plan_path: str | None) -> dict[str, Any]:
    if not plan_path:
        return {}
    p = Path(plan_path)
    if not p.is_file():
        return {}
    try:
        plan = json.loads(p.read_text(encoding="utf-8"))
        return normalize_run_scope(dict(plan.get("scope") or {}), str(plan.get("query") or ""))
    except Exception:
        return {}


def seed_companies_from_scope(scope: dict[str, Any]) -> list[dict[str, str]]:
    # Analyst seeds (from the curated seed file / --seeds) are guaranteed and bypass
    # the channel gate. Auto-enumerated market players are seeds for recall only and
    # are gated normally, so manufacturers never leak into the export.
    analyst_doms = {
        str(d).lower().removeprefix("www.")
        for d in (scope.get("_analyst_seed_domains") or [])
        if str(d).strip()
    }
    rows: list[dict[str, Any]] = []
    for display, domain, fn in normalize_seed_companies(scope.get("seed_companies")):
        if not display or not domain:
            continue
        dom = domain.lower().removeprefix("www.")
        rows.append(
            {
                "name": display,
                "domain": dom,
                "discovery_source": "phase1_seed",
                "company_function": fn or "",
                "is_seed": True,
                "is_analyst_seed": dom in analyst_doms,
            }
        )
    return rows


async def resolve_user_seeds(
    names: list[str],
    settings: Any,
    country: str = "global",
    *,
    pinned_domains: dict[str, str] | None = None,
) -> tuple[list[dict[str, str]], list[str]]:
    """Resolve analyst-provided company names to {canonical_name, primary_domain}.

    Returns (resolved, unresolved_names). Names present in ``pinned_domains`` skip
    the web search entirely and use the supplied official domain directly (robust
    when search backends are rate-limited). For the rest, one quick search per name
    finds the official domain; names with no findable official site are returned in
    unresolved_names (reported to the user).
    """
    from vendor_intel.clients.search_router import FreeSearchRouter
    from vendor_intel.discovery.company_registry import (
        is_blocklisted_domain,
        resolve_official_domain,
    )
    from vendor_intel.utils.domains import domain_from_url

    clean = [str(n).strip() for n in (names or []) if str(n).strip()]
    if not clean:
        return [], []
    pins = {
        str(k).strip().lower(): domain_from_url(str(v)) or str(v).strip().lower().removeprefix("www.")
        for k, v in (pinned_domains or {}).items()
        if str(k).strip() and str(v).strip()
    }
    router = FreeSearchRouter(settings)
    geo = (country or "global").strip()
    out: list[dict[str, str]] = []
    unresolved: list[str] = []
    for name in clean:
        pinned = pins.get(name.strip().lower())
        if pinned and not is_blocklisted_domain(pinned):
            out.append(
                {"canonical_name": name, "primary_domain": pinned, "company_function": "", "source": "analyst"}
            )
            print(f"  [seeds] pinned '{name}' -> {pinned} (from seed file)", flush=True)
            continue
        q = f"{name} official website" + (f" {geo}" if geo and geo.lower() != "global" else "")
        try:
            rows = await router.search(q, market=name, geo=geo, search_topic=name, discovery_mode=False)
        except Exception:
            rows = []
        cands: list[str] = []
        for r in rows[:8]:
            d = domain_from_url(getattr(r, "link", "") or "")
            if d and not is_blocklisted_domain(d) and d not in cands:
                cands.append(d)
        # Only accept a domain that actually matches the company name — never grab the top
        # unrelated result (that put mercadopago.com.br under 'V-TEX', a directory under
        # 'Donadon'). Unmatched names are reported as unresolved (honest), not mis-attached.
        dom = resolve_official_domain(name, cands)
        if dom:
            out.append(
                {"canonical_name": name, "primary_domain": dom, "company_function": "", "source": "analyst"}
            )
            print(f"  [seeds] resolved '{name}' -> {dom}", flush=True)
        else:
            unresolved.append(name)
            print(f"  [seeds] could not resolve an official domain for '{name}'", flush=True)
    return out, unresolved


def merge_user_seeds_into_scope(
    scope: dict[str, Any],
    resolved: list[dict[str, str]],
    *,
    analyst: bool = False,
) -> dict[str, Any]:
    """Prepend seeds to scope.seed_companies (deduped by domain).

    When ``analyst=True`` the resolved domains are recorded in
    ``scope['_analyst_seed_domains']`` so downstream gates can tell true analyst
    seeds (guaranteed, bypass the channel gate) apart from auto-enumerated market
    players (recall only, gated normally).
    """
    if not resolved:
        return scope
    existing = list(scope.get("seed_companies") or [])
    resolved_doms = {r["primary_domain"].lower().removeprefix("www.") for r in resolved}
    if analyst:
        # Analyst seeds take priority — put them first so they survive the seed cap.
        merged = list(resolved)
        for s in existing:
            d = (s.get("primary_domain") if isinstance(s, dict) else "") or ""
            if d.lower().removeprefix("www.") not in resolved_doms:
                merged.append(s)
    else:
        # Auto-enumerated players go AFTER existing (analyst) seeds so they can't
        # push guaranteed seeds past the cap.
        merged = list(existing)
        existing_doms = set()
        for s in existing:
            d = (s.get("primary_domain") if isinstance(s, dict) else "") or ""
            existing_doms.add(d.lower().removeprefix("www."))
        for r in resolved:
            if r["primary_domain"].lower().removeprefix("www.") not in existing_doms:
                merged.append(r)
    scope = dict(scope)
    scope["seed_companies"] = merged
    if analyst:
        analyst_doms = set(scope.get("_analyst_seed_domains") or [])
        for r in resolved:
            d = str(r.get("primary_domain") or "").lower().removeprefix("www.")
            if d:
                analyst_doms.add(d)
        scope["_analyst_seed_domains"] = sorted(analyst_doms)
    return scope


def merge_seeds_first(
    companies: list[dict[str, str]],
    scope: dict[str, Any],
) -> tuple[list[dict[str, str]], set[str]]:
    """Seeds at front; dedupe by domain."""
    seeds = seed_companies_from_scope(scope)
    seed_doms = {s["domain"] for s in seeds if s.get("domain")}
    by_dom: dict[str, dict[str, str]] = {}
    for s in seeds:
        dom = s.get("domain") or ""
        if dom:
            by_dom[dom] = s
    for c in companies:
        dom = (c.get("domain") or "").strip().lower()
        if not dom:
            continue
        if dom not in by_dom:
            by_dom[dom] = {**c, "is_seed": False}
    ordered = list(by_dom.values())
    ordered.sort(key=lambda x: (0 if x.get("is_seed") else 1, x.get("name", "")))
    return ordered, seed_doms
