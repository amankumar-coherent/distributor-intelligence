# -*- coding: utf-8 -*-
import asyncio
import re
import sys
from pathlib import Path

import httpx
from dotenv import load_dotenv

ROOT = Path(r"D:\vendor-intel")
sys.path.insert(0, str(ROOT / "src"))
load_dotenv(ROOT / ".env")

from vendor_intel.clients.google_ai_scraper import google_ai_ask  # noqa: E402


async def main() -> None:
    queries = [
        "Packaging Warehouse Australia packagingwarehouse.com.au LinkedIn company size employees",
        "Verpackungszentrale GmbH Hamburg verpackungszentrale.de Mitarbeiter Anzahl Impressum",
        "site:northdata.de Verpackungszentrale GmbH Mitarbeiter",
    ]
    for q in queries:
        print("===", q)
        data = await google_ai_ask(q, llm_clean=True)
        md = str(data.get("markdown") or "")
        print(md[:700].replace("\n", " "))
        print()

    ua = {"User-Agent": "Mozilla/5.0"}
    urls = [
        "https://www.verpackungszentrale.de/",
        "https://www.verpackungszentrale.de/impressum",
        "https://www.packagingwarehouse.com.au/",
        "https://web.archive.org/cdx/search/cdx?url=packagingwarehouse.com.au&output=json&limit=5",
        "https://www.northdata.de/?q=Verpackungszentrale+GmbH",
    ]
    for u in urls:
        try:
            r = httpx.get(u, headers=ua, verify=False, timeout=25.0, follow_redirects=True)
            t = re.sub(r"(?is)<[^>]+>", " ", r.text or "")
            t = re.sub(r"\s+", " ", t).strip()
            print("FETCH", r.status_code, u, t[:350])
        except Exception as e:
            print("ERR", u, e)


if __name__ == "__main__":
    asyncio.run(main())
