"""Rebuild _neon_filtered_export.json from current FINAL Founded+HQ set, then optionally push."""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "src"))

from export_founded_hq_v2 import find_main_json  # noqa: E402


def rebuild(market_dir: Path) -> Path:
    # Reuse export_one filter logic without rewriting Excel if locked — call filter directly.
    from export_founded_hq_v2 import has_founded, has_headquarters

    json_path = find_main_json(market_dir)
    data = json.loads(json_path.read_text(encoding="utf-8"))
    rows = [r for r in (data.get("relevant_companies") or []) if r.get("is_relevant", True)]

    _channel_mode = False
    try:
        from vendor_intel.funnel.channel_partner_types import channel_partner_mode_enabled

        _channel_mode = channel_partner_mode_enabled(data.get("query_context") or {})
    except Exception:
        pass
    _keep_flag = "is_analyst_seed" if _channel_mode else "is_seed"
    filtered = [
        r for r in rows if (has_founded(r) and has_headquarters(r)) or r.get(_keep_flag)
    ]
    try:
        from vendor_intel.funnel.channel_partner_types import is_production_exportable

        if _channel_mode:
            filtered = [r for r in filtered if is_production_exportable(r)]
    except Exception:
        pass

    ctx = data.get("query_context") or {}
    market = str(ctx.get("industry") or data.get("query") or market_dir.name)
    country = str(ctx.get("country") or "global")
    out = {
        "query": market,
        "query_context": {"industry": market, "country": country},
        "relevant_companies": filtered,
        "completed_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "filter_note": "Founded + HQ FINAL set for Neon / Distributor Intelligence",
    }
    path = market_dir / "_neon_filtered_export.json"
    path.write_text(json.dumps(out, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"{market_dir.name}: {len(filtered)} companies → {path}")
    return path


def main() -> None:
    base = Path(r"C:\Users\aman.kumar\Desktop\coherent\output\run4")
    slugs = [
        "global_smartwatch_market_global",
        "global_green_chemical_market_global",
        "global_firewall_market_global",
    ]
    push = "--push" in sys.argv
    paths = [rebuild(base / s) for s in slugs]
    if push:
        from vendor_intel.neon_push import push_json_to_neon

        for p in paths:
            print(f"Pushing {p.parent.name}...")
            stats = push_json_to_neon(p, merge_companies=False)
            print(json.dumps(stats, indent=2, default=str))


if __name__ == "__main__":
    main()
