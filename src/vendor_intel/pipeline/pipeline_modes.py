"""Pipeline run modes — presets for speed vs depth.

Modes unify cap, core-pipeline tuning, post-export steps, and batch defaults so
one flag (`--mode bulk`) replaces guessing dozens of CLI options.

  bulk     — 10k+ markets on 1 PC (~5–12 min/market): small cap, no heavy discovery,
               pipeline + CSV/JSON only; run gap-fill / quality pass later.
  standard — Default batch: pipeline + exports, no post-enrich (same as before).
  quality  — Pipeline + full post_market (gap-fill → Apollo → audit → Excel → Neon).
  full     — quality + deep post-export enrich + batch audit + gap-fill + Neon push.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Literal

from vendor_intel.pipeline.cap_profiles import apply_cap

PipelineMode = Literal["bulk", "standard", "quality", "full"]

MODE_KEYS: tuple[PipelineMode, ...] = ("bulk", "standard", "quality", "full")

# Fast gap-fill sources for bulk-lite post (no Apollo, no audit).
BULK_LITE_GAP_SOURCES = "backfill,linkedin"


@dataclass(frozen=True)
class PipelineModeSpec:
    key: PipelineMode
    label: str
    description: str
    cap: str
    profile: str
    default_max_concurrent: int
    fmcg_xlsx: bool
    settings_overrides: dict[str, Any] = field(default_factory=dict)
    # post_market.py step toggles; None = skip post_market entirely.
    post_market: dict[str, Any] | None = None
    # Legacy batch worker flags (full-enrich / with-audit / gap-fill / push_neon).
    full_enrich: bool = False
    search_backfill: bool = False
    with_audit: bool = False
    gap_fill: bool = False
    push_neon: bool = False


MODES: dict[PipelineMode, PipelineModeSpec] = {
    "bulk": PipelineModeSpec(
        key="bulk",
        label="Bulk",
        description=(
            "Fastest — ~5–12 min/market on 1 PC. Small company cap, skips multi-source "
            "discovery and player enumeration, exports CSV/JSON only. Use gap_fill_batch "
            "or re-run with --mode quality on shortlists later."
        ),
        cap="bulk",
        profile="quality",
        default_max_concurrent=3,
        fmcg_xlsx=False,
        settings_overrides={
            "pipeline_multi_source_discovery_enabled": False,
            "pipeline_enumerate_players": False,
            "pipeline_use_ssc": True,
            "pipeline_post_audit_enabled": False,
            "pipeline_enrich_concurrent": 8,
            "pipeline_classify_concurrent": 8,
            "phase3_agentic_validation": False,
            "widen_if_unique_lt": 40,
            "discovery_force_widen_below": 30,
        },
        post_market=None,
    ),
    "standard": PipelineModeSpec(
        key="standard",
        label="Standard",
        description="Pipeline + CSV/DOCX/JSON/XLSX — no post-enrich (previous batch default).",
        cap="standard",
        profile="quality",
        default_max_concurrent=2,
        fmcg_xlsx=True,
        post_market=None,
    ),
    "quality": PipelineModeSpec(
        key="quality",
        label="Quality",
        description=(
            "Full post_market: gap-fill → LinkedIn verify Founded+HQ → Apollo names → "
            "contact resolve → brand fix → market audit → Founded+HQ Excel → Neon. "
            "Uses maximum cap (≥300 companies via country-wise discovery)."
        ),
        cap="maximum",
        profile="quality",
        default_max_concurrent=1,
        fmcg_xlsx=True,
        settings_overrides={
            "target_unique_companies": 300,
            "widen_if_unique_lt": 300,
            "discovery_yield_min_unique": 300,
            "pipeline_global_export_min_rows": 300,
        },
        post_market={
            "gap_fill": True,
            "linkedin_firmographic_verify": True,
            "linkedin_verify_concurrent": 1,
            "apollo_names": True,
            "resolve_partials": True,
            "brand_domain_resolve": True,
            "market_relevance_audit": True,
            "export_excel": True,
            "push_neon": True,
            "gap_concurrent": 4,
        },
    ),
    "full": PipelineModeSpec(
        key="full",
        label="Full",
        description="Maximum depth: quality post_market + deep enrich + batch audit + gap-fill + Neon.",
        cap="maximum",
        profile="quality",
        default_max_concurrent=1,
        fmcg_xlsx=True,
        settings_overrides={
            "target_unique_companies": 300,
            "widen_if_unique_lt": 300,
            "discovery_yield_min_unique": 300,
            "pipeline_global_export_min_rows": 300,
        },
        post_market={
            "gap_fill": True,
            "linkedin_firmographic_verify": True,
            "linkedin_verify_concurrent": 1,
            "apollo_names": True,
            "resolve_partials": True,
            "brand_domain_resolve": True,
            "market_relevance_audit": True,
            "export_excel": True,
            "push_neon": True,
            "gap_concurrent": 4,
        },
        full_enrich=True,
        search_backfill=True,
        with_audit=True,
        gap_fill=True,
        push_neon=True,
    ),
}


def get_mode(key: str | None) -> PipelineModeSpec | None:
    if not key:
        return None
    return MODES.get(key.strip().lower())  # type: ignore[arg-type]


def mode_describe(key: str) -> str:
    spec = get_mode(key)
    if not spec:
        return ""
    return f"{spec.label}: {spec.description}"


def apply_mode_settings(settings: Any, mode_key: str, *, cap_override: str | None = None) -> Any:
    """Apply cap + core-pipeline overrides for a mode."""
    spec = get_mode(mode_key)
    if not spec:
        return settings
    cap = cap_override or spec.cap
    updated = apply_cap(settings, cap)
    if spec.settings_overrides:
        updated = updated.model_copy(update=spec.settings_overrides)
    return updated.model_copy(
        update={
            "pipeline_profile": spec.profile,
            "pipeline_recall_mode": spec.profile == "recall",
        }
    )


def resolve_batch_options(
    mode_key: str | None,
    *,
    max_concurrent: int | None,
    cap: str | None,
    profile: str | None,
    fmcg_xlsx: bool | None,
    full_enrich: bool,
    with_audit: bool,
    gap_fill: bool,
    push_neon: bool,
    search_backfill: bool,
) -> dict[str, Any]:
    """Merge explicit CLI flags with mode defaults (explicit flags win)."""
    spec = get_mode(mode_key)
    if not spec:
        return {
            "cap": cap or "standard",
            "profile": profile or "quality",
            "max_concurrent": max_concurrent if max_concurrent is not None else 1,
            "fmcg_xlsx": False if fmcg_xlsx is False else True,
            "full_enrich": full_enrich,
            "with_audit": with_audit,
            "gap_fill": gap_fill,
            "push_neon": push_neon,
            "search_backfill": search_backfill,
            "post_market": None,
            "mode": None,
        }
    return {
        "cap": cap or spec.cap,
        "profile": profile or spec.profile,
        "max_concurrent": (
            max_concurrent if max_concurrent is not None else spec.default_max_concurrent
        ),
        "fmcg_xlsx": spec.fmcg_xlsx if fmcg_xlsx is None else fmcg_xlsx,
        "full_enrich": full_enrich or spec.full_enrich,
        "with_audit": with_audit or spec.with_audit,
        "gap_fill": gap_fill or spec.gap_fill,
        "push_neon": push_neon or spec.push_neon,
        "search_backfill": search_backfill or spec.search_backfill,
        "post_market": spec.post_market,
        "mode": spec.key,
    }
