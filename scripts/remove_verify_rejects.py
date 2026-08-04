#!/usr/bin/env python3
"""Mark verify-reject companies as is_relevant=false and re-export Founded+HQ Excel."""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT))

from export_founded_hq_v2 import export_one  # noqa: E402


def main(slug: str = "cyber_physical_systems_security") -> None:
    market_dir = ROOT / "output" / slug
    json_path = market_dir / "global_cyber_physical_systems_security_market_global.json"
    verify_path = market_dir / "verify_cache.json"

    data = json.loads(json_path.read_text(encoding="utf-8"))
    verify = json.loads(verify_path.read_text(encoding="utf-8"))
    rejects = {r["company"].lower(): r for r in verify if r.get("overall_verdict") == "reject"}

    if not rejects:
        print("No reject companies in verify_cache.json")
        return

    print("Reject companies to remove from export:")
    for name, v in rejects.items():
        print(f"  - {v.get('company')} ({v.get('website', '')})")

    removed: list[dict] = []
    for key in ("relevant_companies", "all_classified"):
        for row in data.get(key) or []:
            name = str(row.get("company") or row.get("brand") or "").strip()
            if name.lower() not in rejects:
                continue
            row["is_relevant"] = False
            row["verify_verdict"] = "reject"
            v = rejects[name.lower()]
            row["verify_reason"] = v.get("reason") or v.get("issues") or "verify reject"
            row["verify_removed_at"] = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
            if not any(x["company"] == name for x in removed):
                removed.append(
                    {
                        "company": name,
                        "website": row.get("website"),
                        "reason": str(row["verify_reason"])[:300],
                    }
                )

    data.setdefault("export_rejected", [])
    existing = {x.get("company", "").lower() for x in data["export_rejected"] if isinstance(x, dict)}
    for item in removed:
        if item["company"].lower() not in existing:
            data["export_rejected"].append(item)

    json_path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Marked {len(removed)} companies is_relevant=false in JSON")

    neon_path = market_dir / "_neon_filtered_export.json"
    if neon_path.is_file():
        neon = json.loads(neon_path.read_text(encoding="utf-8"))
        before = len(neon.get("relevant_companies") or [])
        neon["relevant_companies"] = [
            r
            for r in (neon.get("relevant_companies") or [])
            if str(r.get("company") or "").lower() not in rejects
        ]
        neon_path.write_text(json.dumps(neon, indent=2, ensure_ascii=False), encoding="utf-8")
        print(f"Neon filtered export: {before} -> {len(neon['relevant_companies'])} rows")

    stats = export_one(market_dir, suffix="founded_and_hq_v2")
    print(json.dumps(stats, indent=2))


if __name__ == "__main__":
    slug = sys.argv[1] if len(sys.argv) > 1 else "cyber_physical_systems_security"
    main(slug)
