# -*- coding: utf-8 -*-
"""Rewrite Desktop/final json heading id as integer newsId (or null)."""
from __future__ import annotations

import json
from pathlib import Path

from _export_seven_json import MARKETS, neon_market_id

OUT_DIR = Path(r"C:\Users\aman.kumar\Desktop\final json")


def main() -> None:
    for m in MARKETS:
        path = OUT_DIR / m["json_name"]
        if not path.exists():
            print("MISSING", path)
            continue
        data = json.loads(path.read_text(encoding="utf-8"))
        mid = neon_market_id(m["market"])
        data["market"] = m["market"]
        data["id"] = mid
        path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
        print(f"{path.name} id={mid} ({type(mid).__name__})")


if __name__ == "__main__":
    main()
