"""User-selectable result caps — how many companies a run aims for.

A cap scales discovery breadth (volume prompts + widen loops), the discover/enrich
ceilings, and the export-row ceiling together, so the user can trade depth for speed.
Actual exported counts still depend on how many real companies pass quality.
"""
from __future__ import annotations

from typing import Any

# key -> {label, approx, time, discover, enrich, export, volume, widen}
CAP_TIERS: dict[str, dict[str, Any]] = {
    "bulk": {
        "label": "Bulk",
        "approx": "~15-35 companies",
        "time": "fastest bulk (~5-12 min)",
        "discover": 60, "enrich": 60, "export": 40, "volume": 12, "widen": 1,
    },
    "focused": {
        "label": "Focused",
        "approx": "~25-50 companies",
        "time": "fastest (~10-15 min)",
        "discover": 120, "enrich": 120, "export": 60, "volume": 20, "widen": 2,
    },
    "standard": {
        "label": "Standard",
        "approx": "~50-90 companies",
        "time": "~15-25 min",
        "discover": 220, "enrich": 220, "export": 120, "volume": 32, "widen": 3,
    },
    "broad": {
        "label": "Broad",
        "approx": "~90-160 companies",
        "time": "~25-40 min",
        "discover": 340, "enrich": 340, "export": 190, "volume": 46, "widen": 4,
    },
    "maximum": {
        "label": "Maximum",
        "approx": "~300+ companies (country sweep)",
        "time": "slowest (~60-180 min with country sweep)",
        "discover": 800, "enrich": 800, "export": 2000, "volume": 60, "widen": 5,
    },
}

DEFAULT_CAP = "broad"


def cap_keys() -> list[str]:
    return list(CAP_TIERS.keys())


def cap_describe(key: str) -> str:
    t = CAP_TIERS.get(key)
    if not t:
        return ""
    return f"{t['label']}: {t['approx']} - {t['time']}"


def apply_cap(settings: Any, cap_key: str | None) -> Any:
    """Return settings with discover/enrich/export ceilings + discovery breadth set by the cap."""
    t = CAP_TIERS.get((cap_key or "").strip().lower())
    if not t:
        return settings
    updates: dict[str, Any] = {
        "pipeline_discover_max": t["discover"],
        "pipeline_global_discover_max": t["discover"],
        "pipeline_enrich_max": t["enrich"],
        "pipeline_global_enrich_max": t["enrich"],
        "pipeline_export_max_rows": t["export"],
        "pipeline_global_export_max_rows": t["export"],
        "volume_prompt_count": t["volume"],
        "pipeline_global_volume_prompt_count": t["volume"],
        "widen_loop_max": t["widen"],
    }
    # Maximum tier aligns with country-sweep ≥300 company target
    if (cap_key or "").strip().lower() == "maximum":
        updates["target_unique_companies"] = max(
            int(getattr(settings, "target_unique_companies", 300) or 300), 300
        )
        updates["pipeline_global_export_min_rows"] = max(
            int(getattr(settings, "pipeline_global_export_min_rows", 0) or 0), 300
        )
        updates["widen_if_unique_lt"] = max(
            int(getattr(settings, "widen_if_unique_lt", 300) or 300), 300
        )
        updates["discovery_yield_min_unique"] = max(
            int(getattr(settings, "discovery_yield_min_unique", 300) or 300), 300
        )
    return settings.model_copy(update=updates)
