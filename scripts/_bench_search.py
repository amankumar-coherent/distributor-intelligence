import asyncio
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
from dotenv import load_dotenv

load_dotenv(ROOT / ".env")
from vendor_intel.clients.search_router import FreeSearchRouter
from vendor_intel.config import Settings


async def main() -> None:
    s = Settings.load()
    r = FreeSearchRouter(s)
    t0 = time.time()
    hits = await r.search("Ingram Micro CEO")
    print("elapsed", round(time.time() - t0, 1), "hits", len(hits or []))
    for h in hits or []:
        title = getattr(h, "title", None) or (h.get("title") if isinstance(h, dict) else "")
        print("-", (title or "")[:90])


if __name__ == "__main__":
    asyncio.run(main())
