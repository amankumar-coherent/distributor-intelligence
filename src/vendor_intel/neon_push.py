"""Push vendor-intel JSON exports to Neon (used after gap-fill and pipeline runs)."""
from __future__ import annotations

import importlib.util
import os
from pathlib import Path
from typing import Any

from dotenv import load_dotenv

from vendor_intel.config import _project_root

load_dotenv(_project_root() / ".env")


def should_push_neon() -> bool:
    return bool((os.environ.get("DATABASE_URL") or "").strip())


def _load_push_module():
    mod_path = _project_root() / "scripts" / "push_export_to_neon.py"
    spec = importlib.util.spec_from_file_location("push_export_to_neon", mod_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load {mod_path}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def push_json_to_neon(json_path: Path | str, *, merge_companies: bool = False) -> dict[str, Any]:
    """Push one market export JSON to Neon. merge_companies=True fills empty DB fields only."""
    path = Path(json_path).resolve()
    if not path.is_file():
        raise FileNotFoundError(path)
    if not should_push_neon():
        raise RuntimeError("DATABASE_URL is not set — cannot push to Neon")
    mod = _load_push_module()
    return mod.push_export(path, merge_companies=merge_companies)


def push_result_to_neon(result: dict, *, label: str = "neon") -> dict[str, Any] | None:
    """Push pipeline result dict to Neon using its _json_path export."""
    json_path = result.get("_json_path")
    if not json_path:
        return None
    if not should_push_neon():
        print(f"  [{label}] Skipped — DATABASE_URL not set", flush=True)
        return None
    path = Path(json_path).resolve()
    print(f"  [{label}] Pushing {path.name} to Neon...", flush=True)
    stats = push_json_to_neon(path, merge_companies=False)
    total = stats.get("vendor_intel_companies", stats.get("companies", 0))
    print(f"  [{label}] Done — {total} companies upserted", flush=True)
    return stats


def push_json_after_gap_fill(json_path: Path | str) -> dict[str, Any] | None:
    """Merge-push gap-filled JSON so the website picks up new HQ / founded / etc."""
    if not should_push_neon():
        return None
    path = Path(json_path).resolve()
    print(f"  [neon] Merging gap-fill into Neon — {path.name}", flush=True)
    stats = push_json_to_neon(path, merge_companies=True)
    merged = stats.get("vendor_intel_companies_merged", 0)
    total = stats.get("vendor_intel_companies", 0)
    print(f"  [neon] Done — {total} companies upserted ({merged} merged with existing)", flush=True)
    return stats
