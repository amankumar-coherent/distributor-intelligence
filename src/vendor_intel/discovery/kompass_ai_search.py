"""AI search-provider abstraction for Kompass market/product-page discovery.

Scope is intentionally narrow: the AI is ONLY used to (1) locate the correct
Kompass market/product page for a given market name and (2) assist
distributor-status verification when deterministic HTML rules are
insufficient (see kompass_distributors.py). It never invents company data —
every AI-proposed URL/ID is verified against an actual fetched Kompass page
before being trusted (see verify_product_page()).

Provider selection: AI_SEARCH_PROVIDER (gemini|openai|deepseek), with
AI_FALLBACK_PROVIDER tried automatically if the primary provider errors or
returns nothing usable. Both read from vendor_intel.placeholders.llm's
existing provider dispatch (ANTHROPIC/GEMINI/GROQ/OPENAI keys) so API keys
and provider plumbing are not duplicated — "openai" here means whatever
OPENAI_API_KEY/OPENAI_BASE_URL/OPENAI_MODEL currently point at, which is
DeepSeek-compatible by simply setting OPENAI_BASE_URL to DeepSeek's endpoint
(the existing convention in this repo — see chatgpt_expand.py's
_is_deepseek()). "deepseek" is accepted as an explicit alias for the same
OpenAI-compatible path pointed at DeepSeek's API.
"""
from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass

from vendor_intel.clients import kompass as kompass_client

_AI_CACHE_NAMESPACE = "kompass_page_discovery"
_UNDERSTANDING_CACHE_NAMESPACE = "kompass_market_understanding"


def _log(msg: str) -> None:
    print(f"  [kompass-ai] {msg}", flush=True)


def ai_search_provider() -> str:
    return (os.getenv("AI_SEARCH_PROVIDER") or "gemini").strip().lower()


def ai_fallback_provider() -> str | None:
    v = (os.getenv("AI_FALLBACK_PROVIDER") or "").strip().lower()
    return v or None


def _provider_chain() -> list[str]:
    chain = [ai_search_provider()]
    fb = ai_fallback_provider()
    if fb and fb not in chain:
        chain.append(fb)
    # Try Google AI Overview FIRST when it's available. It searches the live
    # web, so it can return a Kompass URL that genuinely exists right now;
    # an LLM answers from training data and tends to produce plausible URLs
    # that then fail verification. The LLM stays in the chain as fallback.
    try:
        from vendor_intel.clients.google_ai_scraper import google_ai_enabled

        if google_ai_enabled() and "google_ai" not in chain:
            chain.insert(0, "google_ai")
    except Exception:
        pass
    # DeepSeek is reached through the same OpenAI-compatible dispatch branch;
    # normalize here so callers can say AI_SEARCH_PROVIDER=deepseek directly.
    return [("openai" if p == "deepseek" else p) for p in chain]


def market_understanding_enabled() -> bool:
    return (os.getenv("KOMPASS_MARKET_UNDERSTANDING_ENABLED", "true") or "true").strip().lower() not in (
        "0",
        "false",
        "no",
        "off",
    )


def _define_market_via_chatgpt_scraper(market: str) -> str:
    """Step A: ask the free ChatGPT web-UI scraper what the market actually
    means/includes AND, specifically, what the closest matching real
    kompass.com category/classification name is — e.g. for "Wearable Medical
    Devices" ChatGPT knows the real Kompass taxonomy uses a broader category
    like "Medical equipment and instruments" / "Medical instruments,
    electronic", not the narrow literal market name. Asking only for a
    generic definition (the old prompt) throws this real-world taxonomy
    knowledge away and just paraphrases the market name back, which is no
    more useful to Kompass page-discovery than the bare name alone.
    Returns "" on any failure (disabled, not configured, network error) so
    the caller falls back to searching with the bare market name, same as
    before this feature existed."""
    from vendor_intel.scraping.google_ai_mode import ask_generative, generative_backend_enabled

    if not generative_backend_enabled():
        _log("market definition: skipped (no generative web backend enabled)")
        return ""
    try:
        reply = ask_generative(
            f"What does the {market} market mean? What products, materials, or "
            "services does it include?\n"
            "Separately: kompass.com (a global B2B company directory) organizes "
            "companies into broader product/classification categories, not exact "
            f"market names — what is the closest real Kompass category or "
            f"subcategory name that companies in the {market} market would "
            "actually be listed under? Name the specific category/subcategory, "
            "not just a paraphrase of the market name.\n"
            "Answer in 2-4 plain sentences total, no markdown, no preamble — just "
            "the definition and the Kompass category name."
        )
        return (reply or "").strip()
    except Exception as exc:
        _log(f"market definition: generative backend error — {type(exc).__name__}: {exc}")
        return ""


def _distill_search_terms_via_deepseek(market: str, definition: str) -> str:
    """Step B: send the raw definition to DeepSeek (OpenAI-compatible dispatch,
    same as the rest of this repo) to distill it into a short, precise phrase
    better suited to matching an actual Kompass product/classification page
    name than the raw market name or the full multi-sentence definition would
    be. Returns "" on any failure — caller falls back to the raw market name."""
    api_key = (os.getenv("OPENAI_API_KEY") or "").strip()
    if not api_key:
        _log("market understanding: DeepSeek skipped (OPENAI_API_KEY not set)")
        return ""
    try:
        from vendor_intel.config import Settings
        from vendor_intel.pipeline.chatgpt_expand import _chat_json, _client, _model

        settings = Settings()
        client = _client(settings)
        model = _model(settings)
        system = (
            "You distill a market definition into a short, precise search phrase "
            "for finding the matching product/classification category on a B2B "
            "company directory (kompass.com). Do not invent new products/materials "
            "not implied by the definition. If the definition names or implies a "
            "specific real category/subcategory on that directory (e.g. 'Medical "
            "equipment and instruments' or 'Medical instruments, electronic'), use "
            "THAT broader category name as the search phrase — directories organize "
            "by broad category, not by narrow market names, so a named category is "
            "always more useful here than a paraphrase of the market name itself."
        )
        user = (
            f'Market name: "{market}"\n'
            f'Definition (may name a specific directory category to use instead of '
            f'the bare market name): "{definition}"\n\n'
            "Reply with a single JSON object: "
            '{"search_terms": "<a short phrase, 3-8 words, naming the specific '
            'directory category if one was named, otherwise the core product/'
            'material/category this market covers>"}.'
        )
        data = _chat_json(client, model, system, user, label="kompass-market-understand", max_tokens=200, retries=2)
        terms = str((data or {}).get("search_terms") or "").strip()
        return terms
    except Exception as exc:
        _log(f"market understanding: DeepSeek error — {type(exc).__name__}: {exc}")
        return ""


def understand_market(market: str) -> str:
    """New step: ChatGPT scraper defines what the market means -> DeepSeek
    distills that into concise search terms -> returned text is used (in
    place of the bare market name) as the input to discover_product_page(),
    so the AI page-discovery step is working from an actual understanding of
    the market rather than pattern-matching the raw name alone. Falls back
    to the bare market name (returns "") whenever ChatGPT scraper or
    DeepSeek aren't available/configured or fail — this step is a quality
    improvement, never a hard requirement for the pipeline to run."""
    if not market_understanding_enabled():
        return ""

    from vendor_intel.clients.search_cache_disk import get_disk_cached_search, put_disk_cached_search

    cached = get_disk_cached_search(market, "", _UNDERSTANDING_CACHE_NAMESPACE)
    if cached:
        return str(cached[0].get("search_terms") or "")

    definition = _define_market_via_chatgpt_scraper(market)
    if not definition:
        return ""
    _log(f"market definition for {market!r}: {definition[:160]}{'…' if len(definition) > 160 else ''}")

    terms = _distill_search_terms_via_deepseek(market, definition)
    if not terms:
        return ""
    _log(f"market understanding for {market!r} -> search terms: {terms!r}")

    put_disk_cached_search(
        market, "", _UNDERSTANDING_CACHE_NAMESPACE, [{"definition": definition, "search_terms": terms}]
    )
    return terms


@dataclass(frozen=True)
class ProductPageCandidate:
    market: str
    kompass_product_url: str
    kompass_product_id: str
    confidence: float
    provider_used: str


_DISCOVERY_SYSTEM = """You help locate pages on kompass.com, a global B2B company directory.
Given a market/product name, identify the single most likely Kompass product/classification
page URL for that market — the page listing companies classified under that product category.
Respond with ONLY a JSON object, no markdown fences, no commentary:
{"kompass_product_url": "<full https url on kompass.com, or empty string if unsure>",
 "confidence": <0.0-1.0>}
Rules:
- The URL MUST be on the kompass.com domain.
- Prefer the international/global page structure when you know of one (paths containing "/z/ww/"),
  but any real kompass.com product/classification URL is acceptable.
- If you are not reasonably confident of a real URL, return an empty string for
  kompass_product_url and a low confidence — do NOT fabricate a plausible-looking URL."""


_KOMPASS_URL_RE = re.compile(r"https?://[a-z0-9.-]*kompass\.com/[^\s)\"'<>]+", re.I)


def _ask_google_ai_overview(market: str, *, search_terms: str = "") -> dict | None:
    """Find a Kompass category URL by searching, not by asking an LLM.

    Returns the same {"kompass_product_url", "confidence"} shape as the LLM
    path so callers don't care which produced it. The URL is still verified
    by a real fetch downstream — nothing here is trusted on its own.
    """
    import asyncio

    from vendor_intel.clients.google_ai_scraper import google_ai_ask, google_ai_enabled

    if not google_ai_enabled():
        return None

    subject = search_terms or market
    query = (
        f"site:kompass.com {subject} — what is the Kompass category page URL "
        f"listing companies for {market}?"
    )
    try:
        data = asyncio.run(google_ai_ask(query))
    except Exception as exc:  # noqa: BLE001
        _log(f"google_ai: {type(exc).__name__}: {exc}")
        return None

    text = (data.get("markdown") or "") + " " + " ".join(
        str(c.get("url") or "") for c in (data.get("citations") or []) if isinstance(c, dict)
    )
    for url in _KOMPASS_URL_RE.findall(text):
        # Only an activity/classification page is useful — a company profile
        # or the homepage is not what we're looking for.
        if kompass_client.extract_product_id(url):
            return {"kompass_product_url": url.rstrip(".,);"), "confidence": 0.8}
    return None


def _ask_provider(provider: str, market: str, *, search_terms: str = "") -> dict | None:
    """Call one provider via the existing LLM dispatch. Returns the parsed
    JSON dict, or None on any failure (network, non-JSON, wrong provider not
    configured, etc.) — callers fall through to the next provider in chain.

    search_terms, when given (from understand_market()), is included
    alongside the raw market name so the provider is working from an actual
    understanding of the market, not just pattern-matching the bare name.

    provider="google_ai" asks Google AI Overview instead of an LLM. Worth
    trying FIRST: it searches the live web, so it can surface a Kompass URL
    that actually exists today, whereas an LLM answers from training data
    and tends to invent plausible-looking URLs that fail verification."""
    from vendor_intel.placeholders import llm as ph

    user_msg = f"Market/product: {market}"
    if search_terms:
        user_msg += f"\nWhat this market covers: {search_terms}"

    if provider == "google_ai":
        return _ask_google_ai_overview(market, search_terms=search_terms)

    original_provider = ph.LLM_PROVIDER
    try:
        ph.LLM_PROVIDER = provider
        if not ph.is_configured():
            return None
        raw = ph.llm_complete(
            _DISCOVERY_SYSTEM,
            user_msg,
            max_tokens=300,
        )
        text = (raw or "").strip()
        # Strip markdown fences some models add despite instructions not to.
        if text.startswith("```"):
            text = text.strip("`").removeprefix("json").strip()
        start, end = text.find("{"), text.rfind("}")
        if start >= 0 and end > start:
            text = text[start : end + 1]
        return json.loads(text)
    except Exception as exc:
        _log(f"provider {provider} failed: {type(exc).__name__}: {exc}")
        return None
    finally:
        ph.LLM_PROVIDER = original_provider


def discover_product_page(market: str) -> ProductPageCandidate | None:
    """Ask the AI provider chain for a candidate Kompass product/market page,
    then verify it against a REAL fetched page before returning anything.
    Never trusts an AI-invented ID — the product id is always re-extracted
    from the verified URL itself, not taken from the AI's JSON.

    Before searching, runs understand_market() — ChatGPT scraper defines
    what the market means, DeepSeek distills that into short search terms —
    so the page-discovery prompt is working from an actual understanding of
    the market rather than the bare name alone. This step degrades silently
    to "" (bare market name only, exactly the prior behavior) whenever
    ChatGPT scraper / DeepSeek aren't available — it's a quality
    improvement, never a hard dependency."""
    from vendor_intel.clients.search_cache_disk import (
        get_disk_cached_search,
        put_disk_cached_search,
    )

    cached = get_disk_cached_search(market, "", _AI_CACHE_NAMESPACE)
    if cached:
        row = cached[0]
        return ProductPageCandidate(
            market=market,
            kompass_product_url=row["kompass_product_url"],
            kompass_product_id=row["kompass_product_id"],
            confidence=float(row.get("confidence", 0.0)),
            provider_used=row.get("provider_used", "cache"),
        )

    search_terms = understand_market(market)

    for provider in _provider_chain():
        data = _ask_provider(provider, market, search_terms=search_terms)
        if not data:
            continue
        url = str(data.get("kompass_product_url") or "").strip()
        confidence = float(data.get("confidence") or 0.0)
        if not url or not kompass_client.is_kompass_url(url):
            _log(f"{provider}: no usable URL (confidence={confidence:.2f}) — trying next provider")
            continue

        # MUST verify against the real page — never trust the AI's URL/ID blind.
        page = kompass_client.fetch_kompass_page(url)
        if not page.alive:
            _log(f"{provider}: proposed URL not reachable ({page.error}) — rejecting: {url}")
            continue
        product_id = kompass_client.extract_product_id(page.final_url) or kompass_client.extract_product_id(url)
        if not product_id:
            _log(f"{provider}: verified page has no extractable product id — rejecting: {url}")
            continue

        candidate = ProductPageCandidate(
            market=market,
            kompass_product_url=page.final_url,
            kompass_product_id=product_id,
            confidence=confidence,
            provider_used=provider,
        )
        put_disk_cached_search(
            market,
            "",
            _AI_CACHE_NAMESPACE,
            [
                {
                    "kompass_product_url": candidate.kompass_product_url,
                    "kompass_product_id": candidate.kompass_product_id,
                    "confidence": candidate.confidence,
                    "provider_used": candidate.provider_used,
                }
            ],
        )
        _log(
            f"resolved market {market!r} -> product_id={product_id} "
            f"via {provider} (confidence={confidence:.2f})"
        )
        return candidate

    _log(f"no provider found a verifiable Kompass product page for {market!r}")
    return None
