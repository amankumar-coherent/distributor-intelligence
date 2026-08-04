#!/usr/bin/env python3
"""Scrape avocado oil market research from IMARC, SkyQuest, Research and Markets."""
from __future__ import annotations

import argparse
import asyncio
import json
import os
import sys
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

MARKET_SOURCES: tuple[dict[str, str], ...] = (
    {
        "source": "IMARC Group",
        "domain": "imarcgroup.com",
        "query": "site:imarcgroup.com avocado oil market size CAGR",
    },
    {
        "source": "SkyQuest Technology",
        "domain": "skyquestt.com",
        "query": "site:skyquestt.com avocado oil market",
    },
    {
        "source": "Research and Markets",
        "domain": "researchandmarkets.com",
        "query": "site:researchandmarkets.com avocado oil market",
    },
)


async def _extract_report_fields(source: str, title: str, url: str, text: str) -> dict[str, Any]:
    from backend.config import get_settings, sync_opencode_to_openai_env
    from openai import AsyncOpenAI

    sync_opencode_to_openai_env()
    cfg = get_settings()
    if not cfg.openai_api_key:
        return {
            "source": source,
            "title": title,
            "source_url": url,
            "snippet": text[:500],
            "notes": "OPENAI_API_KEY missing — raw snippet only",
        }
    client = AsyncOpenAI(api_key=cfg.openai_api_key)
    model = (os.getenv("OPENAI_MODEL") or cfg.model_extract or "gpt-4o-mini").strip()
    prompt = (
        f"Source: {source}\nURL: {url}\nTitle: {title}\n\n"
        "Extract market research facts from the text. Return JSON with keys: "
        "market_size, cagr, key_players (array), regions (array), segments (array), notes. "
        "Use only stated facts; empty strings/arrays if unknown.\n\n"
        f"{text[:8000]}"
    )
    try:
        resp = await client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"},
            temperature=0,
        )
        raw = json.loads((resp.choices[0].message.content or "{}").strip() or "{}")
    except Exception as exc:
        raw = {"notes": f"extract failed: {exc}"}
    if not isinstance(raw, dict):
        raw = {}
    raw["source"] = source
    raw["title"] = title
    raw["source_url"] = url
    if not raw.get("notes"):
        raw["notes"] = text[:300]
    return raw


async def scrape_market_research(
    *,
    topic: str = "avocado oil",
    output_path: Path | None = None,
    max_per_source: int = 3,
) -> list[dict[str, Any]]:
    from vendor_intel.clients.search_router import FreeSearchRouter
    from vendor_intel.config import Settings
    from vendor_intel.placeholders.load_keys import apply_env_overrides
    from vendor_intel.scraping.ddgs_extract import fetch_page_via_ddgs_extract

    apply_env_overrides()
    settings = Settings.load()
    router = FreeSearchRouter(settings)
    reports: list[dict[str, Any]] = []
    seen_urls: set[str] = set()

    for src in MARKET_SOURCES:
        q = src["query"].replace("avocado oil", topic)
        print(f"[market-research] {src['source']}: {q}", flush=True)
        hits = await router.search(q, discovery_mode=True)
        count = 0
        for hit in hits:
            if count >= max_per_source:
                break
            link = (hit.link or "").split("#")[0].rstrip("/")
            if src["domain"] not in link.lower() or link.lower() in seen_urls:
                continue
            seen_urls.add(link.lower())
            page = await asyncio.to_thread(fetch_page_via_ddgs_extract, link)
            text = (page.text or "").strip() if page.alive else (hit.snippet or "")
            if len(text) < 40:
                text = f"{hit.title}. {hit.snippet}"
            item = await _extract_report_fields(src["source"], hit.title, link, text)
            reports.append(item)
            count += 1
            print(f"  + {hit.title[:60]}", flush=True)

    if output_path:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(json.dumps(reports, indent=2, default=str), encoding="utf-8")
        print(f"[market-research] wrote {output_path.resolve()}", flush=True)
    return reports


def main() -> None:
    parser = argparse.ArgumentParser(description="Scrape market research for avocado oil")
    parser.add_argument(
        "--output",
        default=str(ROOT / "output" / "avocado_oil" / "market_research_avocado_oil.json"),
    )
    parser.add_argument("--topic", default="avocado oil")
    parser.add_argument("--max-per-source", type=int, default=3)
    args = parser.parse_args()
    asyncio.run(
        scrape_market_research(
            topic=args.topic,
            output_path=Path(args.output),
            max_per_source=args.max_per_source,
        )
    )


if __name__ == "__main__":
    main()
