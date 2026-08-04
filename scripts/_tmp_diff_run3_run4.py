"""Diff run3 vs run4 relevant companies for the three markets."""
from __future__ import annotations

import json
import re
from pathlib import Path

ROOT3 = Path(r"C:\Users\aman.kumar\Desktop\coherent\output\run3")
ROOT4 = Path(r"C:\Users\aman.kumar\Desktop\coherent\output\run4")
MARKETS = [
    "global_smartwatch_market_global",
    "global_green_chemical_market_global",
    "global_firewall_market_global",
]


def norm_name(name: str) -> str:
    s = re.sub(r"[^a-z0-9]+", " ", (name or "").lower()).strip()
    return re.sub(r"\s+", " ", s)


def companies(path: Path) -> list[dict]:
    data = json.loads(path.read_text(encoding="utf-8"))
    return [r for r in (data.get("relevant_companies") or []) if r.get("is_relevant", True)]


def main() -> None:
    for mid in MARKETS:
        p3 = ROOT3 / mid / f"{mid}.json"
        p4 = ROOT4 / mid / f"{mid}.json"
        c3 = companies(p3)
        c4 = companies(p4)
        m3 = {norm_name(r.get("company") or r.get("brand") or ""): r for r in c3}
        m4 = {norm_name(r.get("company") or r.get("brand") or ""): r for r in c4}
        only3 = sorted(k for k in m3 if k and k not in m4)
        only4 = sorted(k for k in m4 if k and k not in m3)
        both = sorted(k for k in m3 if k and k in m4)
        print(f"\n=== {mid} ===")
        print(f"  run3={len(c3)} run4={len(c4)} overlap={len(both)} only_run3={len(only3)} only_run4={len(only4)}")
        print("  ONLY IN RUN3 (to add):")
        for k in only3:
            r = m3[k]
            crawl = r.get("_crawl") or {}
            print(
                f"    - {(r.get('company') or '')[:45]:45} | "
                f"role={(r.get('role') or '')[:20]:20} | "
                f"founded={crawl.get('founded_year') or '?'} | "
                f"hq={(str(crawl.get('headquarters') or '')[:30])}"
            )
        if only4[:8]:
            print("  only in run4 (sample):", ", ".join(only4[:8]))


if __name__ == "__main__":
    main()
