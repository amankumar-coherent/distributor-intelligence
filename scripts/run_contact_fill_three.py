"""Fill contacts for green / firewall / smartphone CLEANED finals only."""
from __future__ import annotations

import asyncio
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import clean_verify_five_finals as cv  # noqa: E402


async def main() -> None:
    keys = os.getenv("FILL_MARKETS", "smartphone,firewall,green").split(",")
    keys = [k.strip() for k in keys if k.strip()]
    cv.MARKETS = [m for m in cv.MARKETS if m["key"] in keys]
    # preserve order: smartphone (smallest empty), firewall, green
    order = {k: i for i, k in enumerate(keys)}
    cv.MARKETS.sort(key=lambda m: order.get(m["key"], 99))
    await cv.main_async()


if __name__ == "__main__":
    asyncio.run(main())
