"""Shared page fetch — ddgs extract → httpx/trafilatura → Agent Reach Jina Reader.

Jina (via Agent-Reach-main) reads any public URL as clean markdown, including
LinkedIn company pages, GitHub org pages, and JS-heavy corporate sites.
"""
from __future__ import annotations

import asyncio


async def fetch_page_text(url: str, *, max_chars: int = 14000) -> tuple[str, str]:
    """Return (text, final_url). Empty text on failure."""
    if not (url or "").strip():
        return "", url

    try:
        from vendor_intel.scraping.ddgs_extract import fetch_page_via_ddgs_extract

        page = await asyncio.to_thread(fetch_page_via_ddgs_extract, url)
        if page.alive and len((page.text or "").strip()) >= 40:
            return (page.text or "")[:max_chars], page.final_url or url
    except Exception:
        pass

    try:
        import httpx

        async with httpx.AsyncClient(
            follow_redirects=True,
            timeout=14.0,
            headers={"User-Agent": "Mozilla/5.0 (compatible; VendorIntel/1.0)"},
        ) as client:
            resp = await client.get(url)
            if resp.status_code < 400 and len(resp.text) > 200:
                try:
                    import trafilatura

                    extracted = trafilatura.extract(resp.text, include_comments=False)
                    if extracted and len(extracted.strip()) >= 40:
                        return extracted[:max_chars], str(resp.url)
                except Exception:
                    pass
                return resp.text[:max_chars], str(resp.url)
    except Exception:
        pass

    try:
        from vendor_intel.integrations.agent_reach_client import read_url_markdown

        jina_text = await read_url_markdown(url)
        if jina_text and len(jina_text.strip()) >= 40:
            return jina_text[:max_chars], url
    except Exception:
        pass

    return "", url
