"""Kompass (kompass.com) HTTP client — page fetch, URL parsing, disk caching,
rate limiting. No HTML parsing/business logic here (see
vendor_intel.discovery.kompass_distributors for that) — this module is pure
fetch + URL/ID extraction + cache plumbing, mirroring the conventions in
clients/search_cache_disk.py and clients/rate_limiter.py.

Kompass URL shapes (verified against real Kompass URLs, not assumed):
  Company profile : https://www.kompass.com/z/ww/c/<slug>/<company-id>/
  Product/market   : https://www.kompass.com/.../<slug>/<product-id>/
Company IDs and product/classification IDs use different formats and must
never be confused — see extract_company_id() / extract_product_id().
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

import httpx

from vendor_intel.clients import rate_limiter as _rate_limiter

RATE_ENGINE = "kompass"
_DEFAULT_RATE = (2.0, 0.3)  # conservative — a real commercial B2B directory, not a search engine

_BROWSER_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"
)
_DEFAULT_HEADERS = {
    "User-Agent": _BROWSER_UA,
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Connection": "keep-alive",
}

# Company ID: alpha prefix (2+ letters) + digits, e.g. MD001143, US123456.
_COMPANY_ID_RE = re.compile(r"/([A-Za-z]{2,4}\d{4,})/?(?:$|[?#])")
# Product/classification ID: pure numeric segment near the end of a Kompass
# product/market URL, e.g. .../smartphones/4066041/
_PRODUCT_ID_RE = re.compile(r"/(\d{4,})/?(?:$|[?#])")


def _log(msg: str) -> None:
    print(f"  [kompass] {msg}", flush=True)


def kompass_enabled() -> bool:
    return (os.getenv("KOMPASS_SCRAPER_ENABLED", "true") or "true").strip().lower() not in (
        "0",
        "false",
        "no",
        "off",
    )


def request_delay_seconds() -> float:
    try:
        return max(0.0, float(os.getenv("REQUEST_DELAY_SECONDS", "2")))
    except ValueError:
        return 2.0


def max_retries() -> int:
    try:
        return max(0, int(os.getenv("MAX_RETRIES", "3")))
    except ValueError:
        return 3


def is_kompass_url(url: str) -> bool:
    try:
        host = urlparse(url).netloc.lower()
    except Exception:
        return False
    return host.endswith("kompass.com")


def is_kompass_company_url(url: str) -> bool:
    """True only for an actual Kompass company PROFILE url (contains /c/)."""
    if not is_kompass_url(url):
        return False
    return "/c/" in urlparse(url).path.lower()


def extract_company_id(url: str) -> str | None:
    """Extract the Kompass COMPANY id from a company profile URL, e.g.
    https://www.kompass.com/z/ww/c/accent-electronic-sa/md001143/ -> MD001143.
    Returns None if the URL isn't a recognizable company-profile URL — never
    guesses, and never returns a numeric-only product id here."""
    if not is_kompass_company_url(url):
        return None
    path = urlparse(url).path.rstrip("/")
    m = _COMPANY_ID_RE.search(path + "/")
    if not m:
        return None
    return m.group(1).upper()


def extract_company_slug(url: str) -> str | None:
    """Extract the company slug segment, e.g. 'accent-electronic-sa'."""
    if not is_kompass_company_url(url):
        return None
    parts = [p for p in urlparse(url).path.split("/") if p]
    company_id = extract_company_id(url)
    if not company_id or company_id.lower() not in {p.lower() for p in parts}:
        return None
    idx = [p.lower() for p in parts].index(company_id.lower())
    return parts[idx - 1] if idx > 0 else None


def extract_product_id(url: str) -> str | None:
    """Extract a Kompass PRODUCT/CLASSIFICATION id (numeric) from a market/
    product page URL. Distinct from extract_company_id() — a product id is
    purely numeric, a company id is alpha-prefixed. Returns None if the URL
    isn't a Kompass URL or has no trailing numeric id segment, or if it looks
    like a company profile URL (to avoid ever conflating the two ID kinds)."""
    if not is_kompass_url(url) or is_kompass_company_url(url):
        return None
    path = urlparse(url).path.rstrip("/")
    m = _PRODUCT_ID_RE.search(path + "/")
    if not m:
        return None
    return m.group(1)


def extract_global_region_code(url: str) -> str | None:
    """Extract the region segment from a Kompass global URL, e.g. 'ww' from
    /z/ww/c/... Returns None if not present (e.g. a country-specific portal)."""
    parts = [p for p in urlparse(url).path.split("/") if p]
    if len(parts) >= 2 and parts[0] == "z":
        return parts[1]
    return None


@dataclass
class KompassPage:
    url: str
    final_url: str
    alive: bool
    html: str = ""
    status_code: int = 0
    error: str = ""
    from_cache: bool = False


def _cache_dir() -> Path:
    from vendor_intel.config import _project_root

    raw = (os.getenv("KOMPASS_CACHE_DIR") or "data/kompass_cache").strip()
    base = Path(raw)
    if not base.is_absolute():
        base = _project_root() / base
    base.mkdir(parents=True, exist_ok=True)
    return base


def _cache_key(url: str) -> str:
    return hashlib.sha256(url.strip().lower().encode("utf-8")).hexdigest()[:40]


def _cache_ttl_sec() -> int:
    try:
        return max(3600, int(os.getenv("KOMPASS_CACHE_TTL_SEC", str(14 * 24 * 3600))))
    except ValueError:
        return 14 * 24 * 3600


def _cache_enabled() -> bool:
    return (os.getenv("KOMPASS_CACHE_ENABLED", "true") or "true").strip().lower() not in (
        "0",
        "false",
        "no",
        "off",
    )


def _get_cached_page(url: str) -> KompassPage | None:
    if not _cache_enabled():
        return None
    path = _cache_dir() / f"{_cache_key(url)}.json"
    if not path.is_file():
        return None
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        saved_at = float(payload.get("saved_at") or 0)
        if time.time() - saved_at > _cache_ttl_sec():
            path.unlink(missing_ok=True)
            return None
        return KompassPage(
            url=payload["url"],
            final_url=payload["final_url"],
            alive=bool(payload["alive"]),
            html=payload.get("html", ""),
            status_code=int(payload.get("status_code") or 0),
            error=payload.get("error", ""),
            from_cache=True,
        )
    except Exception:
        return None


def _put_cached_page(url: str, page: KompassPage) -> None:
    if not _cache_enabled() or not page.alive:
        return
    path = _cache_dir() / f"{_cache_key(url)}.json"
    try:
        path.write_text(
            json.dumps(
                {
                    "url": page.url,
                    "final_url": page.final_url,
                    "alive": page.alive,
                    "html": page.html,
                    "status_code": page.status_code,
                    "error": page.error,
                    "saved_at": time.time(),
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
    except Exception:
        pass  # best-effort — never fail a fetch over a cache write


def _fetch_via_real_browser(url: str) -> KompassPage | None:
    """Fallback for pages httpx can't get past DataDome on: the project's
    shared Selenium Chrome instance (same one used for general web
    scraping — see vendor_intel.scraping.selenium_browser), forced
    non-headless since DataDome specifically fingerprints and blocks
    headless Chrome. Returns None on failure — caller treats that the same
    as the original 403.

    NOTE: driving the user's own real, logged-in Chrome profile was tried
    and does NOT work here — this machine's Chrome is enterprise-managed
    ("Managed by your organization"), and such policies block
    --remote-debugging-port on the managed profile, so neither ChromeDriver's
    own launch path nor a manual launch-then-attach can ever connect (Chrome
    opens but never exposes a debug port). An anonymous browser session is
    sufficient anyway: Kompass listings are public, so no login is needed,
    and a fresh non-headless session passes DataDome's device check."""
    try:
        from vendor_intel.scraping import selenium_browser as sb
    except Exception:
        return None
    try:
        # DataDome fingerprints headless Chrome specifically and blocks it
        # even with anti-automation flags set — force a real, visible window
        # for this fallback (same as the ChatGPT scraper, which is
        # non-headless by default and works).
        previous_headless = sb.HEADLESS
        sb.HEADLESS = False
        sb.shutdown_chrome_driver()  # force a fresh driver with the new setting
        try:
            result = sb.fetch_page_html(url)
        finally:
            sb.HEADLESS = previous_headless
            sb.shutdown_chrome_driver()  # don't leave a visible window/session lingering
    except Exception as exc:  # noqa: BLE001
        _log(f"browser fallback fetch failed: {type(exc).__name__}: {exc}")
        return None
    html = result.html or ""
    is_challenge_page = "captcha-delivery.com" in html or "DataDome Device Check" in html
    alive = result.alive and len(html) > 500 and not is_challenge_page
    return KompassPage(
        url=url,
        final_url=result.final_url or url,
        alive=alive,
        html=html,
        status_code=200 if alive else 0,
        error="" if alive else (result.error or ("datadome_challenge" if is_challenge_page else "")),
    )


def fetch_kompass_page(url: str, *, timeout: float = 20.0) -> KompassPage:
    """Fetch a Kompass page: disk cache -> rate-limited httpx GET with retry/
    backoff. Returns alive=False (never raises) on any failure so callers can
    treat "page not found" as a normal, expected outcome (per spec: verify
    the URL, don't assume it exists)."""
    url = url.strip()
    if not url:
        return KompassPage(url=url, final_url=url, alive=False, error="empty_url")

    cached = _get_cached_page(url)
    if cached is not None:
        return cached

    _rate_limiter.configure(RATE_ENGINE, capacity=_DEFAULT_RATE[0], refill_per_sec=_DEFAULT_RATE[1])
    if not _rate_limiter.acquire(RATE_ENGINE, max_wait=30.0):
        return KompassPage(url=url, final_url=url, alive=False, error="rate_limited")

    delay = request_delay_seconds()
    retries = max_retries()
    last_error = ""
    for attempt in range(retries + 1):
        if attempt > 0:
            backoff = min(60.0, delay * (2 ** attempt))
            time.sleep(backoff)
        try:
            with httpx.Client(
                follow_redirects=True, timeout=timeout, headers=_DEFAULT_HEADERS
            ) as client:
                resp = client.get(url)
                if resp.status_code == 429 or resp.status_code >= 500:
                    last_error = f"HTTP {resp.status_code}"
                    _rate_limiter.note_outcome(RATE_ENGINE, blocked=(resp.status_code == 429))
                    continue
                if resp.status_code == 403:
                    # Kompass fronts pages with DataDome bot-detection, which
                    # blocks plain HTTP clients (no JS/cookies/TLS
                    # fingerprint) with a 403 challenge page regardless of
                    # retries — a real browser is required to get past it.
                    _rate_limiter.note_outcome(RATE_ENGINE, blocked=False)
                    browser_page = _fetch_via_real_browser(url)
                    if browser_page is not None:
                        if browser_page.alive:
                            _put_cached_page(url, browser_page)
                        return browser_page
                    return KompassPage(
                        url=url,
                        final_url=str(resp.url),
                        alive=False,
                        status_code=resp.status_code,
                        error="HTTP 403 (bot-blocked, browser fallback also failed)",
                    )
                if resp.status_code >= 400:
                    _rate_limiter.note_outcome(RATE_ENGINE, blocked=False)
                    return KompassPage(
                        url=url,
                        final_url=str(resp.url),
                        alive=False,
                        status_code=resp.status_code,
                        error=f"HTTP {resp.status_code}",
                    )
                _rate_limiter.note_outcome(RATE_ENGINE, blocked=False)
                html = resp.text or ""
                page = KompassPage(
                    url=url,
                    final_url=str(resp.url),
                    alive=len(html.strip()) > 500,
                    html=html,
                    status_code=resp.status_code,
                )
                if page.alive:
                    _put_cached_page(url, page)
                    if delay > 0:
                        time.sleep(delay)
                return page
        except httpx.TimeoutException:
            last_error = "timeout"
        except Exception as exc:
            last_error = f"{type(exc).__name__}: {exc}"[:200]

    return KompassPage(url=url, final_url=url, alive=False, error=last_error or "fetch_failed")
