"""Post-enrich: verify each channel partner's role from its OWN website.

Channel Partner Intelligence hard rule: every exported company must carry
website evidence of a channel role ("Authorized Distributor", "Official Dealer",
"Wholesale Supplier", "Certified Integrator", "Service Provider", …). This pass
crawls the homepage + About/Partners/Distributors/Dealers/Services pages, finds
the strongest channel-role phrase, and records it as ``evidence`` + the source
``evidence_url``. Companies with no verifiable channel evidence are flagged
``_evidence_verified=False`` so the export gate can drop them.

Merge-only and resume-safe: each row is flagged ``_evidence_attempted`` and
skipped on re-run; ``persist_fn`` is called after every company.
"""
from __future__ import annotations

import re
from typing import Any
from urllib.parse import urlparse

from vendor_intel.integrations.page_fetch import fetch_page_text

# Channel-role evidence cues (strongest / most specific first). Each becomes a
# case-insensitive regex; the first hit on a page is captured as evidence.
_EVIDENCE_CUES: tuple[str, ...] = (
    r"authoris(?:z)?ed\s+distributor",
    r"official\s+distributor",
    r"master\s+distributor",
    r"exclusive\s+distributor",
    r"sole\s+distributor",
    r"authoris(?:z)?ed\s+dealer",
    r"official\s+dealer",
    r"authoris(?:z)?ed\s+reseller",
    r"value[-\s]added\s+reseller",
    r"certified\s+(?:integrator|partner|installer)",
    r"authoris(?:z)?ed\s+(?:integrator|service\s+cent(?:er|re))",
    r"system\s+integrator",
    r"security\s+integrator",
    r"managed\s+security\s+service",
    r"managed\s+service\s+provider",
    r"\bmssp\b",
    r"wholesale\s+(?:supplier|distributor)",
    r"\bwholesaler\b",
    r"retail\s+(?:partner|chain)",
    r"channel\s+partner",
    r"reseller\s+program",
    r"partner\s+program",
    r"installation\s+partner",
    r"repair\s+partner",
    r"maintenance\s+partner",
    r"consulting\s+partner",
    r"technology\s+partner",
    r"solution\s+provider",
    r"service\s+provider",
    r"distributor\s+of\b",
    r"we\s+(?:distribute|wholesale|resell|import|stock|supply)\b",
    r"suppliers?\s+of\b",
    r"\bwe\s+are\s+a\s+(?:leading\s+)?(?:distributor|reseller|dealer|wholesaler|supplier)",
)
_EVIDENCE_RE = re.compile("|".join(f"(?:{c})" for c in _EVIDENCE_CUES), re.I)

# Pages most likely to state a channel role; homepage first, then partner pages.
# Kept short for speed — the crawler stops on the first evidence match.
_EVIDENCE_PATHS: tuple[str, ...] = (
    "",
    "/about",
    "/about-us",
    "/partners",
    "/distributors",
    "/services",
)


async def _fetch(url: str) -> str:
    try:
        text, _ = await fetch_page_text(url, max_chars=16000)
        return text or ""
    except Exception:
        return ""


def _base_urls(row: dict[str, Any]) -> list[str]:
    raw = str(row.get("website") or row.get("domain") or "").strip()
    if not raw:
        return []
    if "://" not in raw:
        raw = "https://" + raw
    p = urlparse(raw)
    host = p.netloc or p.path
    if not host:
        return []
    base = f"https://{host.rstrip('/')}"
    return [base + path for path in _EVIDENCE_PATHS]


def _snippet(text: str, m: re.Match[str], *, max_len: int = 220) -> str:
    start = max(0, m.start() - 60)
    end = min(len(text), m.end() + 120)
    snip = re.sub(r"\s+", " ", text[start:end]).strip()
    return snip[:max_len]


async def channel_evidence(row: dict[str, Any]) -> dict[str, str]:
    """Crawl the company's site for a channel-role evidence phrase + source URL."""
    for url in _base_urls(row):
        text = await _fetch(url)
        if not text:
            continue
        m = _EVIDENCE_RE.search(text)
        if m:
            return {"evidence": _snippet(text, m), "evidence_url": url}
    return {}


def needs_channel_evidence(row: dict[str, Any]) -> bool:
    if not row.get("is_relevant", True):
        return False
    crawl = row.get("_crawl") or {}
    if crawl.get("evidence_attempted") or row.get("_evidence_attempted"):
        return False
    if row.get("_evidence_verified") is True:
        return False
    if not str(row.get("website") or row.get("domain") or "").strip():
        return False
    return True


async def verify_channel_evidence_for_market(
    rows: list[dict[str, Any]],
    *,
    persist_fn: Any = None,
) -> dict[str, Any]:
    """Verify channel-role evidence from each company's own website (merge-only)."""
    todo = [r for r in rows if needs_channel_evidence(r)]
    if not todo:
        return {"queued": 0, "verified": 0, "unverified": 0}

    print(
        f"  [channel_evidence] queue={len(todo)} (crawl own site for channel-role proof)",
        flush=True,
    )

    def _checkpoint() -> None:
        if persist_fn is None:
            return
        try:
            persist_fn()
        except Exception as exc:
            print(f"  [channel_evidence] checkpoint save failed: {exc}", flush=True)

    verified = unverified = errors = 0
    for i, row in enumerate(todo, 1):
        name = str(row.get("company") or row.get("brand") or "").strip()
        try:
            info = await channel_evidence(row)
        except Exception as exc:
            errors += 1
            info = {}
            print(f"  [channel_evidence] [{i}/{len(todo)}] {name[:40]} ERROR: {exc}", flush=True)

        if info.get("evidence") and info.get("evidence_url"):
            row["evidence"] = info["evidence"]
            row["evidence_url"] = info["evidence_url"]
            row["_evidence_verified"] = True
            verified += 1
            print(
                f"  [channel_evidence] [{i}/{len(todo)}] {name[:36]} VERIFIED "
                f"\"{info['evidence'][:60]}\"",
                flush=True,
            )
        else:
            row["_evidence_verified"] = False
            unverified += 1

        crawl = dict(row.get("_crawl") or {})
        crawl["evidence_attempted"] = True
        row["_crawl"] = crawl
        row["_evidence_attempted"] = True
        _checkpoint()

    return {
        "queued": len(todo),
        "verified": verified,
        "unverified": unverified,
        "errors": errors,
    }
