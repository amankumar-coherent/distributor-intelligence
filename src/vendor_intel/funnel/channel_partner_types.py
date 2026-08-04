"""Market-aware channel-partner taxonomy for discovery + export.

Channel Partner Intelligence mode discovers verified channel participants
(distributors, VARs, MSSPs, integrators, converters, …) — not pure manufacturers
or non-companies. Excel / Neon column layout is unchanged: Channel Type is stored
in the existing `role` / Business Type fields.
"""
from __future__ import annotations

import re
from typing import Any

# Canonical Channel Types (Title Case) — written into existing `role` column.
CHANNEL_TYPES: frozenset[str] = frozenset(
    {
        "Distributor",
        "Authorized Distributor",
        "Electronics Distributor",
        "Mobile Device Distributor",
        "Wearable Distributor",
        "Wholesaler",
        "Importer",
        "Dealer",
        "Authorized Dealer",
        "Reseller",
        "VAR",
        "Retailer",
        "Retail Chain",
        "System Integrator",
        "Security Integrator",
        "MSP",
        "MSSP",
        "Service Provider",
        "Service Partner",
        "Repair Partner",
        "Consulting Partner",
        "Technology Partner",
        "OEM Partner",
        "Installer",
        "Contractor",
        "Representative",
        "Agent",
        "Channel Partner",
        "Converter",
        "Fabricator",
    }
)

# Roles that must never export in channel-partner mode (unless dual-role remapped).
REJECT_ROLES: frozenset[str] = frozenset(
    {
        "Manufacturer",
        "Supplier",
        "Industry Body",
        "Project Developer",
        "EPC / Engineering",
        "Technology Provider",  # product/SaaS OEM without channel role
        "Research / Consulting",  # pure research — use Consulting Partner when channel
        "Other",
    }
)

_DEFAULT_CHANNEL_TYPES: list[str] = [
    "Distributor",
    "Wholesaler",
    "Reseller",
    "System Integrator",
    "Service Provider",
    "Channel Partner",
]

# (keyword predicates on market+query lower text) → preferred channel types
_MARKET_CHANNEL_RULES: list[tuple[tuple[str, ...], list[str]]] = [
    (
        ("smartwatch", "smart watch", "wearable", "wearables", "fitness band", "smart band"),
        [
            "Authorized Distributor",
            "Distributor",
            "Wholesaler",
            "Electronics Distributor",
            "Mobile Device Distributor",
            "Wearable Distributor",
            "Authorized Dealer",
            "Dealer",
            "Retailer",
            "Retail Chain",
            "VAR",
            "Channel Partner",
            "Service Partner",
            "Repair Partner",
            "Reseller",
        ],
    ),
    (
        ("physical security", "access control", "cctv", "video surveillance", "intrusion"),
        [
            "Security Integrator",
            "Installer",
            "Service Provider",
            "Distributor",
            "Dealer",
            "System Integrator",
        ],
    ),
    (
        ("cyber-physical", "cyber physical", "ot security", "ics security", "scada security"),
        [
            "System Integrator",
            "MSSP",
            "VAR",
            "Distributor",
            "Consulting Partner",
            "Reseller",
        ],
    ),
    (
        ("cyber", "infosec", "mssp", "endpoint security", "siem", "xdr"),
        [
            "MSSP",
            "MSP",
            "VAR",
            "System Integrator",
            "Consulting Partner",
            "Reseller",
            "Distributor",
        ],
    ),
    (
        ("industrial automation", "plc", "scada", "factory automation", "motion control"),
        ["System Integrator", "Distributor", "VAR", "OEM Partner", "Service Provider"],
    ),
    (
        ("medical device", "medtech", "diagnostic device", "surgical", "ivd"),
        ["Distributor", "Dealer", "Importer", "Reseller", "Channel Partner"],
    ),
    (
        ("food", "beverage", "fmcg", "snack", "dairy", "avocado", "grocery"),
        ["Distributor", "Importer", "Wholesaler", "Dealer", "Channel Partner"],
    ),
    (
        ("foam", "polyurethane foam", "technical foam", "eps foam", "xpe"),
        ["Distributor", "Converter", "Fabricator", "Wholesaler", "Reseller"],
    ),
    (
        ("packaging", "flexible packaging", "corrugated", "labels"),
        ["Distributor", "Reseller", "Wholesaler", "Converter", "Channel Partner"],
    ),
    (
        ("digital signage", "dooh", "video wall", "av integration"),
        ["System Integrator", "Reseller", "VAR", "Distributor", "Installer"],
    ),
]

_FN_TO_CHANNEL: list[tuple[tuple[str, ...], str]] = [
    (("mssp", "managed security"), "MSSP"),
    (("msp", "managed service"), "MSP"),
    (("security integrat",), "Security Integrator"),
    (("system integrat", "integrat"), "System Integrator"),
    (("value added", "value-added", " var", "var ", "vars"), "VAR"),
    (("fabricat",), "Fabricator"),
    (("convert",), "Converter"),
    (("install",), "Installer"),
    (("contract",), "Contractor"),
    (("consult",), "Consulting Partner"),
    (("technolog", "oem partner"), "Technology Partner"),
    (("repair partner", "repair center", "authorized repair", "service & repair"), "Repair Partner"),
    (("service partner",), "Service Partner"),
    (("retail chain", "department store", "hypermarket", "big-box"), "Retail Chain"),
    (("retailer", "retail partner", "retail store", "e-tailer", "ecommerce retailer"), "Retailer"),
    (("authorized dealer", "authorised dealer"), "Authorized Dealer"),
    (("authorized distributor", "authorised distributor", "authorized partner"), "Authorized Distributor"),
    (("wearable distributor", "wearables distributor"), "Wearable Distributor"),
    (("mobile device distributor", "mobile phone distributor", "handset distributor"), "Mobile Device Distributor"),
    (("electronics distributor", "electronic component distributor", "electronics wholesaler"), "Electronics Distributor"),
    (("reseller", "resell"), "Reseller"),
    (("dealer",), "Dealer"),
    (("wholesal",), "Wholesaler"),
    (("import",), "Importer"),
    (("agent",), "Agent"),
    (("represent",), "Representative"),
    (("service provider", "services"), "Service Provider"),
    (("distribut", "channel partner"), "Distributor"),
]


def channel_partner_mode_enabled(
    query_context: dict[str, Any] | None = None,
    *,
    settings: Any = None,
) -> bool:
    """Default ON for vendor-intel; opt out via discovery_mode=landscape or env false."""
    qc = query_context or {}
    mode = str(qc.get("discovery_mode") or "").strip().lower()
    if mode in ("landscape", "manufacturer", "manufacturers", "full_landscape"):
        return False
    if mode in ("channel_partners", "channel", "partners"):
        return True
    if settings is not None:
        return bool(getattr(settings, "pipeline_channel_partner_mode", True))
    return True


def infer_channel_partner_types(market: str, query: str = "") -> list[str]:
    """Dominant channel types for this market (never assumes distributors-only)."""
    low = f"{market} {query}".lower()
    for keys, types in _MARKET_CHANNEL_RULES:
        if any(k in low for k in keys):
            return list(types)
    return list(_DEFAULT_CHANNEL_TYPES)


def ecosystem_labels_for_channel_types(types: list[str]) -> list[str]:
    """Lowercase ecosystem_functions labels used by discovery prompt builders."""
    out: list[str] = []
    seen: set[str] = set()
    for t in types:
        label = str(t or "").strip()
        if not label:
            continue
        key = label.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(key)
    return out


def map_to_channel_type(
    role: str = "",
    company_function: str = "",
    *,
    text: str = "",
) -> str | None:
    """Map classifier/discovery labels → canonical Channel Type, or None."""
    blob = f"{role} {company_function} {text}".lower()
    role_s = (role or "").strip()
    if role_s in CHANNEL_TYPES:
        return role_s

    # Common landscape aliases → channel types
    alias = {
        "distributor": "Distributor",
        "authorized distributor": "Authorized Distributor",
        "electronics distributor": "Electronics Distributor",
        "mobile device distributor": "Mobile Device Distributor",
        "wearable distributor": "Wearable Distributor",
        "authorized dealer": "Authorized Dealer",
        "authorised dealer": "Authorized Dealer",
        "authorised distributor": "Authorized Distributor",
        "retailer": "Retailer",
        "retail chain": "Retail Chain",
        "service partner": "Service Partner",
        "repair partner": "Repair Partner",
        "integrator": "System Integrator",
        "research / consulting": "Consulting Partner",
        "technology provider": "Technology Partner",
    }
    if role_s.lower() in alias:
        return alias[role_s.lower()]

    for keys, channel in _FN_TO_CHANNEL:
        if any(k in blob for k in keys):
            return channel
    return None


def is_accepted_channel_type(role: str) -> bool:
    return (role or "").strip() in CHANNEL_TYPES


# company_function values (from discovery/offline compiler) that mark a company
# as a product MAKER / brand / source — not a channel intermediary. Strict
# channel mode rejects these even when the LLM assigned a channel role like
# "Distributor" or "Channel Partner". NOTE: matched as substrings against the
# lowercased company_function; "service provider" is intentionally NOT here.
_MAKER_FUNCTIONS: tuple[str, ...] = (
    "manufactur", "producer", "oem", "odm", "supplier",
    "technology provider", "technology_provider", "materials provider",
    "market_leader", "product brand", "product_brand", "brand owner",
)


def is_pure_manufacturer_reject(
    role: str,
    *,
    company_function: str = "",
    signals: dict[str, Any] | None = None,
    text: str = "",
) -> bool:
    """True when the company is a product maker/brand/source — not a channel partner.

    Strict channel-partner intent: the honest ``company_function`` (from
    discovery) wins over the LLM-assigned ``role``. A maker is kept only when a
    real (non-LLM) distributor signal is present. We deliberately do NOT trust
    channel wording in free text/summaries — the classifier writes "channel
    partner" into those fields, which would be circular.
    """
    sig = signals or {}
    fn_l = (company_function or "").lower().strip()
    role_l = (role or "").lower().strip()

    # Honest maker/brand/source *function* (from discovery) is the reliable
    # signal and wins over everything — incl. the broadly-set is_distributor
    # flag and any LLM channel role. This is what removes Cisco/Samsung/Solvay
    # (role="Channel Partner"/"Distributor" but company_function="manufacturer").
    if any(k in fn_l for k in _MAKER_FUNCTIONS):
        return True

    # Role-based manufacturer/OEM: reject unless there is a genuine distributor
    # signal or an accepted channel type (a real hybrid maker + distributor).
    if "manufactur" in role_l or "oem" in role_l:
        if sig.get("is_distributor") is True:
            return False
        channel = map_to_channel_type(role, company_function)
        if channel and is_accepted_channel_type(channel):
            return False
        return True

    # Non-maker rows: a real distributor signal or accepted channel type keeps it.
    if sig.get("is_distributor") is True:
        return False
    channel = map_to_channel_type(role, company_function)
    if channel and is_accepted_channel_type(channel):
        return False
    if role.strip() in REJECT_ROLES and not channel:
        return True
    return False


# ── Confidence policy (0.0–1.0 scale; 0–100 tolerated) ─────────────────────
#   >= 0.90        → "auto"   : Auto Export
#   0.70 – 0.89    → "high"   : Export (High Confidence)
#   0.50 – 0.69    → "review" : Export as "Needs Review"
#   < 0.50         → "reject" : dropped entirely
# All non-reject tiers are exported (the review tier is flagged, not withheld).
CONF_AUTO_MIN = 0.90
CONF_HIGH_MIN = 0.70
CONF_REVIEW_MIN = 0.50

_TIER_LABELS = {
    "auto": "Auto Export",
    "high": "High Confidence",
    "review": "Needs Review",
    "reject": "Rejected",
}


def confidence_tier(confidence: Any) -> str:
    """Return 'auto', 'high', 'review', or 'reject' for a confidence value."""
    try:
        c = float(confidence)
    except (TypeError, ValueError):
        c = 0.0
    if c > 1.5:  # tolerate a 0–100 scale
        c = c / 100.0
    if c >= CONF_AUTO_MIN:
        return "auto"
    if c >= CONF_HIGH_MIN:
        return "high"
    if c >= CONF_REVIEW_MIN:
        return "review"
    return "reject"


def confidence_tier_label(confidence: Any) -> str:
    """Human-readable action label for a confidence value."""
    return _TIER_LABELS.get(confidence_tier(confidence), "Rejected")


def apply_confidence_tier_to_rows(rows: list[dict[str, Any]]) -> dict[str, int]:
    """Tag rows by confidence tier and drop only < 0.50.

    Mutates in place: sets ``_confidence_tier`` + ``review_status``. The review
    tier (0.50–0.69) is exported but flagged ``_review_only=True``. Reject tier
    (< 0.50) gets ``is_relevant=False``. Idempotent.
    """
    counts = {"auto": 0, "high": 0, "review": 0, "reject": 0}
    for r in rows:
        # Analyst seeds are guaranteed — tag their tier but never reject them.
        # (Auto-enumerated market players are NOT exempt — they are gated normally.)
        if r.get("is_analyst_seed"):
            tier = confidence_tier(r.get("confidence"))
            r["_confidence_tier"] = tier
            r["review_status"] = _TIER_LABELS[tier] if tier != "reject" else "Needs Review"
            r["is_relevant"] = True
            r.pop("_channel_gate", None) if r.get("_channel_gate") == "low_confidence" else None
            counts[tier if tier != "reject" else "review"] += 1
            continue
        if not r.get("is_relevant", True) and r.get("_channel_gate") != "low_confidence":
            continue
        tier = confidence_tier(r.get("confidence"))
        r["_confidence_tier"] = tier
        r["review_status"] = _TIER_LABELS[tier]
        counts[tier] += 1
        if tier == "reject":
            r["_was_relevant"] = bool(r.get("is_relevant", True))
            r["is_relevant"] = False
            r["_channel_gate"] = "low_confidence"
            r.pop("_review_only", None)
        elif tier == "review":
            r["is_relevant"] = True
            r["_review_only"] = True
            r.pop("_channel_gate", None) if r.get("_channel_gate") == "low_confidence" else None
        else:
            r["is_relevant"] = True
            r.pop("_review_only", None)
            r.pop("_channel_gate", None) if r.get("_channel_gate") == "low_confidence" else None
    return counts


def has_channel_evidence(row: dict[str, Any]) -> bool:
    """True when channel evidence was verified by crawling the official site."""
    return row.get("_evidence_verified") is True


def is_production_exportable(row: dict[str, Any]) -> bool:
    """Channel-partner production gate: confidence >= 0.50 (non-reject tier).

    Analyst seeds are always exportable. Evidence is captured best-effort (see
    channel_evidence_verify) but no longer blocks export. Founded/HQ/market-audit
    are enforced at the call site.
    """
    if row.get("is_analyst_seed"):
        return True
    return confidence_tier(row.get("confidence")) != "reject"


def apply_channel_gate_to_rows(
    rows: list[dict[str, Any]],
    *,
    reason: str = "pure_manufacturer",
) -> int:
    """Strict channel-only filter: mark pure makers/brands as not relevant.

    Mutates each row in place: pure manufacturers/brands get ``is_relevant=False``
    plus ``_channel_gate`` + ``_was_relevant`` (so the change is reversible/auditable).
    Returns the number of rows newly rejected. Safe to call repeatedly.
    """
    rejected = 0
    for r in rows:
        if not r.get("is_relevant", True):
            continue
        if r.get("is_analyst_seed"):  # analyst-guaranteed seed — never gate out
            continue
        if is_pure_manufacturer_reject(
            str(r.get("role") or ""),
            company_function=str(r.get("company_function") or ""),
            signals=r.get("signals") or {},
        ):
            r["_was_relevant"] = bool(r.get("is_relevant", True))
            r["is_relevant"] = False
            r["_channel_gate"] = reason
            rejected += 1
    return rejected


def channel_evidence_snippet(text: str, *, max_len: int = 220) -> str:
    """Pull a short evidence phrase for export_data (does not add Excel columns)."""
    if not text:
        return ""
    m = re.search(
        r".{0,40}\b(?:authorized\s+distributor|value[- ]added\s+reseller|"
        r"managed\s+security|system\s+integrat\w*|channel\s+partner|"
        r"wholesaler|dealer|installer|fabricat\w*|convert\w*).{0,80}",
        text,
        re.I | re.S,
    )
    if not m:
        return ""
    snip = re.sub(r"\s+", " ", m.group(0)).strip()
    return snip[:max_len]


def apply_channel_types_to_scope(
    scope: dict[str, Any],
    market: str,
    query: str = "",
    *,
    query_context: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Ensure scope carries channel_partner_types + channel-focused ecosystem_functions."""
    out = dict(scope or {})
    qc = query_context or {}
    existing = out.get("channel_partner_types") or qc.get("channel_partner_types")
    if isinstance(existing, list) and existing:
        types = [str(t).strip() for t in existing if str(t).strip()]
    else:
        types = infer_channel_partner_types(market, query)
    out["channel_partner_types"] = types
    out["discovery_mode"] = "channel_partners"
    # Replace manufacturer-heavy ecosystem lists with channel types
    out["ecosystem_functions"] = ecosystem_labels_for_channel_types(types)
    return out
