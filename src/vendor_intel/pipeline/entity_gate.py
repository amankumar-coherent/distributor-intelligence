"""
Pre-classification entity gate — reject non-participants (reports, media, marketplaces).

Keeps recall reasonable while blocking the ~40% junk the manager flagged.
Stricter lists live in discovery_fast._PIPELINE_JUNK_* (commented / retired there).
"""
from __future__ import annotations

import re
from typing import Any

from vendor_intel.pipeline.participant_domains import is_market_research_domain
from vendor_intel.validation.site_kind import classify_domain, is_non_product_site

# Hard domain blocks (platforms, gov, research portals)
_HARD_DOMAIN_EXACT = frozenset(
    {
        "researchgate.net",
        "researchgate.com",
        "pubmed.ncbi.nlm.nih.gov",
        "ncbi.nlm.nih.gov",
        "nasdaq.com",
        "polymerupdate.com",
        "exporthub.com",
        "export-hub.com",
        "iea.org",
        "fas.usda.gov",
        "usda.gov",
        "wikipedia.org",
        "facebook.com",
        "linkedin.com",
        "twitter.com",
        "x.com",
        "youtube.com",
        "instagram.com",
        "pinterest.com",
        "medium.com",
        "reddit.com",
        "github.com",
        "figma.com",
        "arduino-ide.org",
        "nature.com",
        "sourceforge.net",
        # User-requested blacklist (marketplaces, classifieds, data/news/portal sites)
        "tiktok.com",
        "globalsources.com",
        "craft.co",
        "signalhire.com",
        "slideserve.com",
        "brighthubengineering.com",
        "downtoearth.org.in",
        "carbike360.com",
        "environmental-expert.com",
        "marinelink.com",
        "profitableventure.com",
    }
)

_HARD_DOMAIN_PARTS = re.compile(
    r"(?:^|\.)("
    r"researchgate|pubmed|ncbi\.nlm|nasdaq|polymerupdate|exporthub|"
    r"indiamart|tradeindia|justdial|crunchbase|zauba|volza|seair|"
    r"alibaba|made-in-china|exportersindia|europages|dial4trade|everychina|ecer|globalsources|"
    r"marketresearch|coherentmarket|emergenresearch|expertmarket|"
    r"theinsightpartners|sphericalinsights|"
    r"chemdive|chemengonline|packaginginsights|farmprogress|"
    r"plasticstoday|worldatlas|handwiki|bloomberg|"
    r"wholesalesugar|braziliansugar|sugarexport|sugarsupplier"
    r")(?:\.|$)",
    re.I,
)

_REPORT_NAME = re.compile(
    r"\b(?:annual\s+report|biofuels?\s+annual|market\s+report|"
    r"industry\s+report|white\s+paper|press\s+release\s+archive)\b",
    re.I,
)

_GOV_EDU_TLD = re.compile(r"\.(?:gov|edu)(?:\.|$)", re.I)

_PHARMA_DOMAIN = re.compile(
    r"(?:^|\.)("
    r"firstwordpharma|emergobyul|mai-cdmo|mrchub|markspark|pharmtech|"
    r"pharmapproach|pharmchoices"
    r")(?:\.|$)",
    re.I,
)

_BAD_NAME = re.compile(
    r"^(?:home\s*[-–]|welcome|untitled|404|error|index|"
    r"top\s+\d+|best\s+\d+|list\s+of\b)",
    re.I,
)

# OS / platform brands mistaken for channel partners (e.g. "Android" → android.com).
_PLATFORM_BRAND_NAMES = frozenset(
    {
        "android",
        "ios",
        "ipados",
        "watchos",
        "wear os",
        "wearos",
        "chrome os",
        "chromeos",
        "macos",
        "windows phone",
    }
)

_PLATFORM_DOMAINS = frozenset(
    {
        "android.com",
    }
)

# Bare product/category labels — not operating companies.
_GENERIC_PRODUCT_NAMES = frozenset(
    {
        "smartwatch",
        "smartwatches",
        "smart watch",
        "smart watches",
        "firewall",
        "firewalls",
        "wearable",
        "wearables",
        "green chemical",
        "green chemicals",
    }
)

# Site nav / portal page labels scraped as company names (e.g. "Support").
_NAV_PAGE_NAMES = frozenset(
    {
        "support",
        "help",
        "contact",
        "contacts",
        "about",
        "about us",
        "home",
        "login",
        "sign in",
        "signin",
        "faq",
        "faqs",
        "blog",
        "news",
        "careers",
        "jobs",
        "docs",
        "documentation",
        "download",
        "downloads",
        "pricing",
        "services",
        "products",
        "welcome",
        "index",
        "customer support",
        "tech support",
        "technical support",
        "product support",
        "platinum support",
    }
)

# Vendor portal / helpdesk subdomains (support.paloaltonetworks.com).
_PORTAL_SUBDOMAIN = re.compile(
    r"^(?:support|help|docs|documentation|community|login|portal|kb|"
    r"knowledge|customer|customers|care|status|download|downloads|"
    r"partners?\.support|support\.partners?)\.",
    re.I,
)

# SEO / directory / listing titles scraped as company names.
# e.g. "Smartwatch Suppliers in Dubai", "Custom Wholesale Chain Smartwatch"
_SEO_COMPANY_NAME = re.compile(
    r"(?:"
    r"\b(?:suppliers?|distributors?|dealers?|wholesalers?|vendors?|"
    r"exporters?|importers?|manufacturers?|retailers?)\s+in\s+[a-z\u00c0-\u024f]"
    r"|"
    r"^(?:custom|wholesale|oem|odm|bulk|cheap)\b.+\b"
    r"(?:smartwatches?|smart\s*watches?|firewalls?|chemicals?|phones?|mobiles?|watches?)\b"
    r"|"
    r"\b(?:buy|shop|order)\s+.+\s+online\b"
    r"|"
    # Product/SKU support titles (not companies)
    r"^(?:software|hardware|network|next\s+generation|ngfw|endpoint)\b.+\bsupport\b"
    r"|"
    r"\b(?:platinum|premium|enterprise|priority)\s+support\s*$"
    r")",
    re.I,
)

_CORPORATE_SUFFIX_LIGHT = re.compile(
    r"\b(?:ltd\.?|limited|inc\.?|corp\.?|corporation|plc|gmbh|"
    r"pvt\.?\s*ltd|llp|llc|co\.?|company|group|holdings?|"
    r"industries|electronics|technologies|solutions|networks|"
    r"distributing|distribution|international)\b",
    re.I,
)


def reject_bad_company_name(name: str, domain: str = "") -> str | None:
    """
    Reject SEO page titles, platform brands, and bare product labels used as
    company names. Reasons here are never trust_llm-bypassed in quality_export.
    """
    raw = (name or "").strip()
    if not raw:
        return None
    low = re.sub(r"\s+", " ", raw.lower()).strip()
    domain = (domain or "").strip().lower().removeprefix("www.")

    if low in _PLATFORM_BRAND_NAMES or domain in _PLATFORM_DOMAINS:
        return "platform_brand_not_company"

    if low in _GENERIC_PRODUCT_NAMES:
        return "generic_product_name"

    if low in _NAV_PAGE_NAMES:
        return "nav_page_not_company"

    if domain and _PORTAL_SUBDOMAIN.search(domain):
        return "vendor_portal_subdomain"

    if _SEO_COMPANY_NAME.search(raw):
        return "seo_page_title"

    try:
        from vendor_intel.validation.site_kind import is_article_title_name

        if is_article_title_name(raw):
            return "seo_page_title"
    except Exception:
        pass

    # Long keyword-stuffed titles without a corporate form (page H1s).
    words = raw.split()
    if (
        len(words) >= 4
        and not _CORPORATE_SUFFIX_LIGHT.search(raw)
        and re.search(
            r"\b(?:smartwatch|firewall|supplier|distributor|dealer|wholesale|custom)\b",
            raw,
            re.I,
        )
    ):
        return "seo_page_title"

    return None


def reject_domain_only(domain: str, *, text: str = "", name: str = "") -> str | None:
    """Domain-only gate — safe when company name is not yet known (Phase 2 domain filter)."""
    domain = (domain or "").strip().lower().removeprefix("www.")
    if not domain:
        return "missing_domain"

    if domain in _HARD_DOMAIN_EXACT:
        return "blocked_platform"

    if is_market_research_domain(domain):
        return "market_research_site"

    if _HARD_DOMAIN_PARTS.search(domain):
        return "blocked_platform_pattern"

    if _GOV_EDU_TLD.search(domain):
        return "gov_or_edu"

    blob = f"{name} {domain} {text}".lower()
    if any(
        x in blob
        for x in (
            "pubmed central",
            "researchgate",
            "usda",
            "iea.org",
            "farmprogress",
            "the counter",
            "chemical engineering online",
            "plasticstoday",
        )
    ):
        return "non_company_entity"

    site_cls = classify_domain(domain)
    if site_cls in ("media", "education", "aggregator", "finance", "directory", "review"):
        return f"site_kind_{site_cls}"

    if is_non_product_site(domain, text=text, name=name or None):
        return "non_product_site"

    if _PHARMA_DOMAIN.search(domain):
        return "off_topic_pharma_cdmo"

    return None


def reject_reason(name: str, domain: str, *, text: str = "") -> str | None:
    """
    Return rejection reason string, or None if the entity may proceed.
    """
    name = (name or "").strip()
    domain = (domain or "").strip().lower().removeprefix("www.")

    # Name-first: SEO / platform junk must not proceed even on a "real" domain
    # (quality_export may trust_llm-bypass non_product_site).
    if name:
        bad = reject_bad_company_name(name, domain)
        if bad:
            return bad

    if not domain:
        return "missing_domain"
    dom_reason = reject_domain_only(domain, text=text, name=name)
    if dom_reason:
        return dom_reason
    if not name:
        return "missing_name"

    if _REPORT_NAME.search(name):
        return "report_not_company"

    if len(name) < 3 or len(name) > 120:
        return "bad_name_length"

    if name.lower() in ("home", "blog", "news", "contact", "about"):
        return "generic_name"

    if _BAD_NAME.search(name):
        return "article_or_page_title"

    try:
        from vendor_intel.discovery.entity_extract import is_plausible_company_name

        if not is_plausible_company_name(name, domain):
            return "not_plausible_company_name"
    except Exception:
        pass

    return None


def filter_companies(
    companies: list[dict[str, str]],
    *,
    keep_registry: bool = True,
) -> tuple[list[dict[str, str]], list[dict[str, Any]]]:
    """Return (kept, rejected_with_reason)."""
    kept: list[dict[str, str]] = []
    rejected: list[dict[str, Any]] = []
    for c in companies:
        name = str(c.get("name") or "")
        domain = str(c.get("domain") or "")
        reason = reject_reason(name, domain)
        if reason:
            rejected.append({**c, "reject_reason": reason})
            continue
        kept.append(c)
    return kept, rejected
