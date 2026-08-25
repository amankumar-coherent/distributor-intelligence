"""Channel-partner purity filter (shared by run_batch + OpenAI expand).

Design:
  - Discovery / recall stay WIDE (market coverage).
  - Before Excel export, drop OEMs, pure manufacturers, utilities/IPPs, and
    junk that are not distributors / VADs / resellers / dealers / wholesalers /
    system integrators / channel partners.

Env:
  CHANNEL_PURITY_FILTER=true|false   (default true)
  CHANNEL_PURITY_MODE=balanced|strict (default balanced)
  CHANNEL_PURITY_LLM=true|false      (default true — DeepSeek/OpenAI via OPENAI_*)
  CHANNEL_PURITY_LLM_CHUNK=20        (companies per LLM call)
  CHANNEL_PURITY_WEB=true|false      (default false; optional SearXNG confirm)
"""
from __future__ import annotations

import json
import os
import re
from concurrent.futures import ThreadPoolExecutor
from dataclasses import asdict, dataclass
from typing import Any


def channel_purity_enabled() -> bool:
    raw = (os.getenv("CHANNEL_PURITY_FILTER") or "true").strip().lower()
    return raw in ("1", "true", "yes", "on")


def channel_purity_mode() -> str:
    raw = (os.getenv("CHANNEL_PURITY_MODE") or "balanced").strip().lower()
    return raw if raw in ("balanced", "strict") else "balanced"


def channel_purity_llm_enabled() -> bool:
    raw = (os.getenv("CHANNEL_PURITY_LLM") or "true").strip().lower()
    return raw in ("1", "true", "yes", "on")


def channel_purity_web_enabled() -> bool:
    raw = (os.getenv("CHANNEL_PURITY_WEB") or "false").strip().lower()
    return raw in ("1", "true", "yes", "on")


def _llm_chunk_size() -> int:
    try:
        return max(5, min(40, int(os.getenv("CHANNEL_PURITY_LLM_CHUNK") or "20")))
    except ValueError:
        return 20


CHANNEL_RE = re.compile(
    r"\b(distributor|distribution|wholesaler|wholesale|reseller|dealer|"
    r"channel partner|value[- ]added distributor|\bvad\b|authorized dealer|"
    r"stockist|importer|system integrator|var\b|mssp|epc contractor|"
    r"solar installer|installation partner|merchant wholesaler|"
    r"electrical distributor|pv distributor|authorized partner)\b",
    re.I,
)
OEM_RE = re.compile(
    r"\b(manufacturer|manufactures|oem\b|module maker|inverter manufacturer|"
    r"turbine manufacturer|produces (?:solar|wind|pv|inverters|modules|devices)|"
    r"leading manufacturer|global manufacturer|tier[- ]?1 manufacturer|"
    r"original equipment)\b",
    re.I,
)
UTILITY_RE = re.compile(
    r"\b(utility|independent power producer|\bipp\b|power producer|asset owner|"
    r"project developer|owns and operates|energy retailer|electricity supplier)\b",
    re.I,
)
NAME_CHANNEL_RE = re.compile(
    r"\b(distributor|distributors|wholesale|wholesaler|dealer|dealers|reseller|"
    r"channel|vad|var)\b",
    re.I,
)

CHANNEL_ROLE_HINT = (
    "distributors, wholesalers, channel partners, VADs, resellers, dealers, "
    "importers, system integrators relevant to the market (NOT pure "
    "OEMs/manufacturers of the core product, NOT utilities/IPPs/project "
    "developers as primary business)"
)

# Templated LLM invents: "Electrical Wholesale (Cameroon)" + electricalwholesale.cm
_TEMPLATED_GENERIC_NAME = re.compile(
    r"^(?:electrical|electro)\s+(?:wholesale|distributors?)\b",
    re.I,
)
_FAKE_GENERIC_HOST = re.compile(
    r"(?:^|\.)(?:electricalwholesale|electrowholesale)"
    r"\.(?:cm|ma|zw|tz|et|ng|gh|co\.zw|co\.tz|co\.nz|com\.ng|com\.et|com\.au)$"
    r"|(?:^|\.)electricaldistributors?(?:et|gh|ng|zw|tz|cm|ma)?\."
    r"(?:com|cm|ma|zw|tz|et|ng|gh|co\.zw|co\.tz|co\.nz|com\.ng|com\.et)$",
    re.I,
)
_AS_DISTRIBUTOR_RE = re.compile(
    r"\((?:as\s+distributor|as\s+separate\s+entity)\)|\bas\s+distributor\b|\bas\s+separate\s+entity\b",
    re.I,
)


def _norm(s: object) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(s or "").lower()).strip()


def _scrub_boilerplate(text: str) -> str:
    t = text or ""
    t = re.sub(r"(?i)channel participant[^.]*\.?", " ", t)
    t = re.sub(r"(?i)verify:\s*", " ", t)
    t = re.sub(r"(?i)\bnot publicly disclosed\b", " ", t)
    return t


def _website_host(website: str) -> str:
    w = str(website or "").strip().lower()
    if not w or w in {"not publicly disclosed", "n/a", "na", "npd", "-"}:
        return ""
    w = re.sub(r"^https?://", "", w)
    w = w.removeprefix("www.").split("/")[0].strip()
    return w


def _is_fake_templated_channel(name: str, website: str = "") -> bool:
    """Drop LLM-invented 'Electrical Wholesale (Cameroon)' + fake TLD domains."""
    n = str(name or "").strip()
    if not _TEMPLATED_GENERIC_NAME.match(n):
        return False
    host = _website_host(website)
    # Country-parenthetical + generated-looking host → fake
    has_country_paren = bool(re.search(r"\([^)]+\)\s*$", n))
    if has_country_paren and host and _FAKE_GENERIC_HOST.search(host):
        return True
    if has_country_paren and not host:
        return True
    # Non-paren generic with African / invented host
    if host and _FAKE_GENERIC_HOST.search(host):
        # Allow the known US firm electricalwholesale.com (no country paren)
        if host in {"electricalwholesale.com"} and not has_country_paren:
            return False
        if has_country_paren or host != "electricalwholesale.com":
            # country clones always fake; bare name with .cm/.ma etc fake
            if has_country_paren or re.search(
                r"\.(cm|ma|zw|tz|et|ng|gh|co\.zw|co\.tz|com\.ng)$", host
            ):
                return True
    return False


_CONNECTOR_WORDS = frozenset(
    {
        "de", "del", "des", "of", "and", "y", "e", "la", "le", "el", "los", "las",
        "da", "do", "dos", "das", "the", "for", "und", "et",
    }
)


def _strip_accents(text: str) -> str:
    import unicodedata

    return "".join(c for c in unicodedata.normalize("NFKD", text) if not unicodedata.combining(c))


def _looks_like_generic_placeholder(name: str, website: str) -> bool:
    """Structural, language-agnostic signal that a name may be an LLM-
    fabricated placeholder rather than a real company — its significant
    words mostly reappear verbatim in its own website domain, with no
    distinctive brand word setting it apart from a plain category label
    (e.g. "Semiconductores de Colombia" living at
    semiconductorescolombia.com, or "Kenya Electronics" at
    kenyaelectronics.co.ke). Deliberately has NO hardcoded list of
    countries/industries in any language — it only checks whether the name
    IS its own domain, which a real company's (usually shorter or more
    distinctively branded) domain rarely is.

    This is a SIGNAL for a real-evidence check, not an automatic drop — a
    genuine small reseller can coincidentally have a descriptive domain
    too, so nothing gets dropped on this alone.
    """
    host = _website_host(website)
    if not host:
        return False
    domain_stem = re.sub(r"\.(com|net|org|co|io|biz|info)(\.\w{2,3})?$", "", host, flags=re.I)
    domain_letters = re.sub(r"[^a-z]", "", domain_stem.lower())
    if len(domain_letters) < 8:
        return False
    words = re.findall(r"[a-z]+", _strip_accents(name).lower())
    significant = [w for w in words if w not in _CONNECTOR_WORDS and len(w) >= 3]
    if len(significant) < 2:
        return False
    matched = sum(1 for w in significant if w in domain_letters)
    return matched >= max(2, round(len(significant) * 0.7))


_NEGATION_RE = re.compile(r"\b(?:not|never|no longer|isn'?t|wasn'?t)\b", re.I)
_ROLE_DETERMINER_RE = re.compile(r"\b(?:a|an|one of the)\b", re.I)
# manufactur\w* covers manufacturer/manufacturing/manufactures/manufactured —
# real evidence often self-describes as an "industrial manufacturing group"
# rather than literally "a manufacturer".
_SELF_MANUFACTURER_RE = re.compile(
    r"\b(?:is|was)\b([\s\S]{0,75}?)"
    r"\b(?:manufactur\w*|producer|oem|original equipment manufacturer)\b",
    re.I,
)
_SELF_DISTRIBUTOR_RE = re.compile(
    r"\b(?:is|was)\b([\s\S]{0,75}?)"
    r"\b(?:distributor|wholesaler|reseller|dealer|stockist|importer|channel partner)\b",
    re.I,
)


def _name_tokens(company: str) -> list[str]:
    return [t for t in re.findall(r"[A-Za-z]+", company or "") if len(t) >= 3]


def _self_described(rx: re.Pattern, text: str, name_tokens: list[str]) -> bool:
    """True if ANY "is/was ... <role noun>" occurrence in the text both (a)
    contains a determiner ("a"/"an"/"one of the") in between, confirming
    this is really an "X is a <role>" self-description and not just two
    unrelated words within range of each other, (b) contains no negation
    anywhere in that span — "is a distributor, not a manufacturer" must NOT
    count as a manufacturer self-description just because "manufacturer"
    appears within the scanned window of "is" — and (c), when a company name
    was given, has that name appear shortly BEFORE the match: evidence text
    (especially Google AI Overview responses) often blends in citation
    snippets about OTHER companies, and an unrelated "X is a manufacturer"
    sentence elsewhere in the blob must not be attributed to the company
    actually being classified. Checks every match, not just the first — a
    text can open with a negated clause ("is not manufactured by...") and
    only state the real self-description later ("...itself is a major
    manufacturing group")."""
    for m in rx.finditer(text):
        mid = m.group(1)
        if not _ROLE_DETERMINER_RE.search(mid) or _NEGATION_RE.search(mid):
            continue
        if not name_tokens:
            return True
        context = text[max(0, m.start() - 120) : m.end()]
        if any(re.search(rf"\b{re.escape(tok)}\b", context, re.I) for tok in name_tokens):
            return True
    return False


def classify_manufacturer_signal(evidence_text: str, company: str = "") -> tuple[bool | None, str]:
    """Pure regex classify of free evidence/summary text (no network, no LLM).

    Returns (keep, reason). keep=None means inconclusive — the caller decides
    the default in that case. Used as the DeepSeek-independent floor for
    channel-fit decisions, since LLM verification (DeepSeek) can silently be
    unavailable (rate limits, 0 balance) and must never be the only thing
    standing between a manufacturer/OEM and the export file.

    A raw word-count of "distributor"/"manufacturer" anywhere in the text is
    unreliable on its own: a genuine OEM's own evidence is often FULL of
    "distributor"/"channel partner" language because it's describing ITS OWN
    downstream network ("Schneider Electric works with a global network of
    distributors"), not itself. So a subject-anchored self-description
    ("X is a leading manufacturer of...") is checked FIRST and takes priority
    over the word count — matching the same "OEM-primary always drops, even
    with a channel arm" policy used in the LLM prompt elsewhere in this file.
    """
    text = evidence_text or ""
    name_tokens = _name_tokens(company)
    self_oem = _self_described(_SELF_MANUFACTURER_RE, text, name_tokens)
    self_dist = _self_described(_SELF_DISTRIBUTOR_RE, text, name_tokens)
    if self_oem:
        return False, "self_described_manufacturer"
    if self_dist:
        return True, "self_described_distributor"

    oem = len(OEM_RE.findall(text))
    ch = len(CHANNEL_RE.findall(text))
    if ch > 0 and ch >= oem:
        return True, f"evidence_channel_signals={ch}"
    if oem > 0 and oem > ch:
        return False, f"evidence_oem_signals={oem}"
    return None, "evidence_inconclusive"


@dataclass
class PurityVerdict:
    keep: bool
    reason: str
    label: str

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def classify_channel_purity(
    name: str,
    *,
    family: str = "general",
    specialty: str = "",
    core_categories: str = "",
    summary: str = "",
    distribution_type: str = "",
    website: str = "",
    extra_text: str = "",
) -> PurityVerdict:
    """Fast heuristic classify (no network / no LLM)."""
    family = (family or "general").strip().lower() or "general"
    name = (name or "").strip()
    if not name:
        return PurityVerdict(False, "empty_name", "drop_unclear")

    if _is_fake_templated_channel(name, website):
        return PurityVerdict(False, "fake_templated_channel", "drop_unrelated")

    if NAME_CHANNEL_RE.search(name):
        return PurityVerdict(True, "name_channel_marker", "keep")

    dist = (distribution_type or "").strip()
    dist_evidence = ""
    if dist and dist.lower() not in {"distributor", "not publicly disclosed", "n/a", "na", ""}:
        dist_evidence = dist

    blob = _scrub_boilerplate(
        "\n".join(
            [
                specialty or "",
                core_categories or "",
                summary or "",
                dist_evidence,
                extra_text or "",
                website or "",
            ]
        )
    )
    # Subject-anchored self-description check FIRST — a raw word count of
    # "distributor"/"manufacturer" anywhere in the blob is unreliable on its
    # own (a real OEM's own text is often full of "distributor"/"channel
    # partner" language describing ITS OWN downstream network, not itself;
    # confirmed by direct testing against ABB/Siemens/Schneider/Eaton/Legrand
    # evidence). Only fall through to the word-count heuristic below when
    # this comes back inconclusive — it doesn't cover utilities, which the
    # word-count path still needs to catch.
    self_signal, self_reason = classify_manufacturer_signal(blob, name)
    if self_signal is False:
        return PurityVerdict(False, self_reason, "drop_oem")
    if self_signal is True:
        return PurityVerdict(True, self_reason, "keep")

    ch = len(CHANNEL_RE.findall(blob))
    oem = len(OEM_RE.findall(blob))
    util = len(UTILITY_RE.findall(blob))
    mode = channel_purity_mode()

    if oem >= 2 and ch == 0:
        return PurityVerdict(False, f"oem_signals={oem}", "drop_oem")
    if util >= 2 and ch == 0:
        return PurityVerdict(False, f"utility_signals={util}", "drop_utility")
    if oem >= 1 and ch == 0 and mode == "strict":
        return PurityVerdict(False, f"oem_only={oem}", "drop_oem")
    if ch >= 1 and ch >= oem:
        return PurityVerdict(True, f"channel_signals={ch} oem={oem}", "keep")
    if CHANNEL_RE.search(specialty or "") and oem <= 1:
        return PurityVerdict(True, "specialty_channel", "keep")
    if mode == "strict":
        if oem >= 1:
            return PurityVerdict(False, f"strict_oem={oem}", "drop_oem")
        if util >= 1:
            return PurityVerdict(False, f"strict_utility={util}", "drop_utility")
        return PurityVerdict(False, "strict_unclear", "drop_unclear")
    if oem >= 1 and ch == 0:
        return PurityVerdict(False, f"oem_leaning={oem}", "drop_oem")
    if util >= 1 and ch == 0:
        return PurityVerdict(False, f"utility_leaning={util}", "drop_utility")
    return PurityVerdict(True, "balanced_keep_coverage", "keep")


def _row_fields(row: dict[str, Any]) -> dict[str, str]:
    name = str(
        row.get("Company")
        or row.get("company")
        or row.get("brand")
        or row.get("name")
        or ""
    ).strip()
    crawl = row.get("_crawl") if isinstance(row.get("_crawl"), dict) else {}
    return {
        "name": name,
        "specialty": str(
            row.get("Specialty Focus")
            or crawl.get("specialty_focus")
            or row.get("specialty_focus")
            or ""
        ),
        "core": str(
            row.get("Core Categories")
            or crawl.get("core_categories")
            or row.get("key_products")
            or ""
        ),
        "summary": str(
            row.get("Summary")
            or row.get("company_summary")
            or row.get("polished_summary")
            or row.get("snippet")
            or row.get("role_description")
            or ""
        ),
        "distribution": str(
            row.get("Distribution Type")
            or row.get("Business Type")
            or row.get("role")
            or ""
        ),
        "website": str(row.get("Website") or row.get("website") or row.get("domain") or ""),
    }


def _resolve_family(family: str, query: str) -> str:
    family = (family or "general").strip().lower() or "general"
    q = (query or "").lower()
    if family == "general" and any(
        t in q for t in ("renewable", "solar", "wind energy", "photovoltaic", "energy storage")
    ):
        return "renewable"
    if family == "general" and any(
        t in q
        for t in (
            "electrical distribution",
            "electrical wholesale",
            "electrical distributor",
            "switchgear",
            "wiring device",
            "lv switchgear",
            "mcb market",
        )
    ):
        return "electrical"
    if family == "general" and "electrical" in q and "distribut" in q:
        return "electrical"
    return family


def purity_role_hint(family: str) -> str:
    return CHANNEL_ROLE_HINT


def _openai_client():
    """OpenAI-compatible client (DeepSeek when OPENAI_BASE_URL points there)."""
    from openai import OpenAI

    try:
        from vendor_intel.config import Settings

        settings = Settings.load()
        key = (settings.openai_api_key or os.getenv("OPENAI_API_KEY") or "").strip()
        base = (settings.openai_base_url or os.getenv("OPENAI_BASE_URL") or "").strip().rstrip("/")
        model = (settings.openai_model or os.getenv("OPENAI_MODEL") or "").strip()
    except Exception:
        key = (os.getenv("OPENAI_API_KEY") or "").strip()
        base = (os.getenv("OPENAI_BASE_URL") or "").strip().rstrip("/")
        model = (os.getenv("OPENAI_MODEL") or "").strip()

    if not key:
        return None, None
    if not base:
        base = "https://api.openai.com/v1"
    if not model:
        model = "deepseek-chat" if "deepseek" in base.lower() else "gpt-4o-mini"
    return OpenAI(api_key=key, base_url=base), model


def _parse_llm_json(text: str) -> Any:
    raw = (text or "").strip()
    m = re.search(r"```(?:json)?\s*([\s\S]*?)```", raw, re.I)
    if m:
        raw = m.group(1).strip()
    start = raw.find("{")
    start_a = raw.find("[")
    if start_a != -1 and (start == -1 or start_a < start):
        raw = raw[start_a:]
    elif start != -1:
        raw = raw[start:]
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return {}


def llm_classify_channel_batch(
    items: list[dict[str, str]],
    *,
    family: str,
    query: str,
) -> dict[str, PurityVerdict]:
    """Classify a batch via DeepSeek/OpenAI. Keyed by normalized company name."""
    if not items:
        return {}
    client, model = _openai_client()
    if client is None or not model:
        print("  [channel-purity] LLM skipped: OPENAI_API_KEY not set", flush=True)
        return {}

    roles = purity_role_hint(family)
    payload = [
        {
            "Company": it.get("name") or "",
            "Website": (it.get("website") or "")[:80],
            "Specialty Focus": (it.get("specialty") or "")[:120],
            "Core Categories": (it.get("core") or "")[:120],
            "Summary": _scrub_boilerplate(it.get("summary") or "")[:180],
        }
        for it in items
    ]
    system = (
        "You are a strict channel-partner analyst. "
        f"Market: {query or family}. Valid KEEP roles: {roles}. "
        "KEEP only if the company's PRIMARY business is distribution / wholesale / "
        "dealer / reseller / VAD / importer / system integrator / authorized channel partner. "
        "DROP (keep=false) if PRIMARY business is OEM/manufacturer of the core product, "
        "utility/IPP/project developer, pure brand owner, media, or unrelated. "
        "If a manufacturer also has a small channel arm, still DROP (OEM-primary). "
        "A name containing Distributor/Wholesale/Dealer is NOT enough — judge PRIMARY business. "
        "Use the Website domain as a hint (catalog/storefront/partner language = channel; "
        "we manufacture/we design/our products = OEM). When unclear, DROP. "
        "Authorized dealers and wholesalers KEEP even if they install. "
        "Return ONLY JSON: "
        '{"results":[{"Company":"...","keep":true|false,'
        '"label":"keep|drop_oem|drop_utility|drop_unclear|drop_unrelated","reason":"short"}]}'
    )
    user = (
        f"Classify each company for channel-partner purity.\n"
        f"Companies:\n{json.dumps(payload, ensure_ascii=False)}"
    )
    content = ""
    try:
        resp = client.chat.completions.create(
            model=model,
            temperature=0,
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
            response_format={"type": "json_object"},
            max_tokens=3500,
        )
        content = (resp.choices[0].message.content or "").strip()
    except Exception:
        try:
            resp = client.chat.completions.create(
                model=model,
                temperature=0,
                messages=[
                    {"role": "system", "content": system},
                    {"role": "user", "content": user + "\nReturn JSON only."},
                ],
                max_tokens=3500,
            )
            content = (resp.choices[0].message.content or "").strip()
        except Exception as exc2:  # noqa: BLE001
            print(
                f"  [channel-purity] LLM error: {type(exc2).__name__}: {exc2}",
                flush=True,
            )
            return {}

    data = _parse_llm_json(content)
    results = data.get("results") if isinstance(data, dict) else data
    out: dict[str, PurityVerdict] = {}
    if not isinstance(results, list):
        return out
    for r in results:
        if not isinstance(r, dict):
            continue
        name = str(r.get("Company") or r.get("name") or "").strip()
        if not name:
            continue
        keep = r.get("keep")
        if keep is None:
            continue
        keep_b = bool(keep) if not isinstance(keep, str) else keep.strip().lower() in (
            "true",
            "1",
            "yes",
            "keep",
        )
        label = str(r.get("label") or ("keep" if keep_b else "drop_unclear")).strip()
        if keep_b:
            label = "keep"
        elif label == "keep":
            label = "drop_unclear"
        reason = str(r.get("reason") or "llm").strip() or "llm"
        out[_norm(name)] = PurityVerdict(keep_b, f"llm:{reason}", label)
    return out


def _is_hard_heuristic_drop(v: PurityVerdict) -> bool:
    if v.keep:
        return False
    if v.reason in {
        "known_oem_blocklist",
        "fake_templated_channel",
        "oem_as_distributor_label",
    }:
        return True
    prefixes = (
        "oem_signals",
        "utility_signals",
        "oem_only",
        "oem_leaning",
        "utility_leaning",
        "strict_oem",
        "strict_utility",
        "strict_unclear",
    )
    return any(v.reason.startswith(p) for p in prefixes)


def _llm_filter_rows(
    rows: list[dict[str, Any]],
    *,
    family: str,
    query: str,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Heuristic pre-drop known OEMs, then DeepSeek classify the rest."""
    family = _resolve_family(family, query)
    hard_keep: list[dict[str, Any]] = []
    hard_drop: list[dict[str, Any]] = []
    to_llm: list[tuple[dict[str, Any], dict[str, str]]] = []

    for row in rows:
        fields = _row_fields(row)
        v = classify_channel_purity(
            fields["name"],
            family=family,
            specialty=fields["specialty"],
            core_categories=fields["core"],
            summary=fields["summary"],
            distribution_type=fields["distribution"],
            website=fields["website"],
        )
        if _is_hard_heuristic_drop(v):
            hard_drop.append(
                {"company": fields["name"], "verdict": v.label, "reason": v.reason}
            )
            continue
        # Name markers (e.g. "XYZ Distribution") still go to LLM — manufacturers
        # sometimes use channel-like subsidiary names.
        to_llm.append((row, fields))

    if not channel_purity_llm_enabled() or not to_llm:
        for row, fields in to_llm:
            v = classify_channel_purity(
                fields["name"],
                family=family,
                specialty=fields["specialty"],
                core_categories=fields["core"],
                summary=fields["summary"],
                distribution_type=fields["distribution"],
                website=fields["website"],
            )
            if v.keep:
                hard_keep.append(row)
            else:
                hard_drop.append(
                    {"company": fields["name"], "verdict": v.label, "reason": v.reason}
                )
        return hard_keep, hard_drop

    client, model = _openai_client()
    provider = "DeepSeek" if model and "deepseek" in model.lower() else "OpenAI"
    if client is None:
        print("  [channel-purity] LLM off (no API key) — heuristic only", flush=True)
        for row, fields in to_llm:
            v = classify_channel_purity(
                fields["name"],
                family=family,
                specialty=fields["specialty"],
                core_categories=fields["core"],
                summary=fields["summary"],
                distribution_type=fields["distribution"],
                website=fields["website"],
            )
            if v.keep:
                hard_keep.append(row)
            else:
                hard_drop.append(
                    {"company": fields["name"], "verdict": v.label, "reason": v.reason}
                )
        return hard_keep, hard_drop

    print(
        f"  [channel-purity] {provider} ({model}) classifying {len(to_llm)} companies "
        f"(chunk={_llm_chunk_size()})…",
        flush=True,
    )
    chunk = _llm_chunk_size()
    llm_map: dict[str, PurityVerdict] = {}
    total_batches = (len(to_llm) + chunk - 1) // chunk
    for i in range(0, len(to_llm), chunk):
        batch = to_llm[i : i + chunk]
        items = [fields for _, fields in batch]
        print(
            f"  [channel-purity] LLM batch {i // chunk + 1}/{total_batches} "
            f"({len(items)} companies)",
            flush=True,
        )
        part = llm_classify_channel_batch(items, family=family, query=query)
        llm_map.update(part)

    kept = list(hard_keep)
    dropped = list(hard_drop)
    for row, fields in to_llm:
        key = _norm(fields["name"])
        v = llm_map.get(key)
        if v is None:
            v = classify_channel_purity(
                fields["name"],
                family=family,
                specialty=fields["specialty"],
                core_categories=fields["core"],
                summary=fields["summary"],
                distribution_type=fields["distribution"],
                website=fields["website"],
            )
            v = PurityVerdict(v.keep, f"heuristic_fallback:{v.reason}", v.label)
        # Final hard veto: never export OEM blocklist / fake templates even if LLM kept
        hard = classify_channel_purity(
            fields["name"],
            family=family,
            specialty=fields["specialty"],
            core_categories=fields["core"],
            summary=fields["summary"],
            distribution_type=fields["distribution"],
            website=fields["website"],
        )
        if _is_hard_heuristic_drop(hard):
            dropped.append(
                {
                    "company": fields["name"],
                    "verdict": hard.label,
                    "reason": f"post_llm_hard:{hard.reason}",
                }
            )
            continue
        if v.keep:
            kept.append(row)
        else:
            dropped.append(
                {"company": fields["name"], "verdict": v.label, "reason": v.reason}
            )
    return kept, dropped


_OEM_COUNT_IN_REASON_RE = re.compile(r"oem=(\d+)")


async def _gai_confirm_ambiguous(
    rows: list[dict[str, Any]], *, family: str, query: str
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Second opinion for rows the heuristic layer found truly ambiguous
    (no channel or OEM signal in the local text at all — the "balanced_keep_
    coverage" default). Asks the Google AI scraper directly, which does not
    depend on DeepSeek balance/availability. Rows the scraper also can't
    confirm stay kept — this only ever removes a row on positive OEM
    evidence, never on absence of evidence.
    """
    import asyncio

    ambiguous: list[tuple[dict[str, Any], dict[str, str]]] = []
    other: list[dict[str, Any]] = []
    for row in rows:
        fields = _row_fields(row)
        v = classify_channel_purity(
            fields["name"],
            family=family,
            specialty=fields["specialty"],
            core_categories=fields["core"],
            summary=fields["summary"],
            distribution_type=fields["distribution"],
            website=fields["website"],
        )
        # Re-confirm via a direct question not just on a pure tie, but also
        # any "keep" that only won via the raw word-count fallback rather
        # than a confident self-description — that fallback path is the
        # same one that false-kept real OEMs before the self-description
        # check existed, and it's still reachable whenever the evidence
        # never uses an "X is a manufacturer/distributor" construction at
        # all (two possible reason shapes reach this: classify_manufacturer_
        # signal's own "evidence_channel_signals=N", and this function's
        # legacy "channel_signals=N oem=M" block below it).
        oem_in_reason = _OEM_COUNT_IN_REASON_RE.search(v.reason)
        needs_confirm = (
            v.reason == "balanced_keep_coverage"
            or v.reason.startswith("evidence_channel_signals")
            or (v.keep and bool(oem_in_reason) and int(oem_in_reason.group(1)) >= 1)
        )
        if needs_confirm:
            ambiguous.append((row, fields))
        else:
            other.append(row)

    if not ambiguous:
        return rows, []

    try:
        from vendor_intel.pipeline.field_verify_fill import GAI_BASE_URL, gai_ask_safe
        from vendor_intel.clients.google_ai_scraper import google_ai_ping
    except Exception:
        return rows, []

    try:
        ok = await google_ai_ping(base_url=GAI_BASE_URL)
    except Exception:
        ok = False
    if not ok:
        return rows, []

    market = query or family
    sem = asyncio.Semaphore(6)

    async def _check(row: dict[str, Any], fields: dict[str, str]):
        async with sem:
            host = _website_host(fields["website"])
            q = (
                f"{fields['name']} {host} is it primarily a manufacturer/OEM that makes "
                f"its own {market} products, or a distributor/wholesaler/reseller/channel "
                f"partner that sells other companies' {market} products"
            )
            md = await gai_ask_safe(q)
            keep, reason = classify_manufacturer_signal(md, fields["name"])
            return row, fields, keep, reason

    checked = await asyncio.gather(*(_check(r, f) for r, f in ambiguous))
    kept = list(other)
    dropped: list[dict[str, Any]] = []
    for row, fields, keep, reason in checked:
        if keep is False:
            dropped.append({"company": fields["name"], "verdict": "drop_oem", "reason": f"gai_confirm:{reason}"})
        else:
            kept.append(row)
    return kept, dropped


async def _gai_confirm_existence(
    rows: list[dict[str, Any]], *, family: str, query: str
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Second opinion for rows _looks_like_generic_placeholder flagged as a
    structural fabrication risk (name mirrors its own domain, no distinctive
    brand word). That signal alone is too weak to drop on — plenty of real
    companies (Richardson RFPD, WPG Holdings) also have a domain matching
    their name — so this asks for independent, third-party evidence of the
    company's existence (news, LinkedIn, an industry directory — NOT just
    the company's own site) and drops ONLY when that comes back empty.
    Anything inconclusive stays kept, same policy as everywhere else here.
    """
    suspicious: list[tuple[dict[str, Any], dict[str, str]]] = []
    other: list[dict[str, Any]] = []
    for row in rows:
        fields = _row_fields(row)
        if _looks_like_generic_placeholder(fields["name"], fields["website"]):
            suspicious.append((row, fields))
        else:
            other.append(row)

    if not suspicious:
        return rows, []

    import asyncio

    try:
        from vendor_intel.pipeline.field_verify_fill import GAI_BASE_URL, gai_ask_safe
        from vendor_intel.clients.google_ai_scraper import google_ai_ping
    except Exception:
        return rows, []

    try:
        ok = await google_ai_ping(base_url=GAI_BASE_URL)
    except Exception:
        ok = False
    if not ok:
        return rows, []

    _ = family, query  # kept for signature symmetry with the other confirm pass
    sem = asyncio.Semaphore(6)
    no_evidence_markers = (
        "no results", "could not find", "couldn't find", "no information",
        "does not appear to exist", "unable to find", "no evidence",
        "no verifiable", "not a recognized", "no record of",
    )

    async def _check(row: dict[str, Any], fields: dict[str, str]):
        async with sem:
            host = _website_host(fields["website"])
            q = (
                f'"{fields["name"]}" company — independent evidence it is a real business '
                f"(news coverage, LinkedIn company page, industry directory listing, "
                f"business registry), not just its own website {host}"
            )
            md = await gai_ask_safe(q)
            return row, fields, md

    checked = await asyncio.gather(*(_check(r, f) for r, f in suspicious))
    kept = list(other)
    dropped: list[dict[str, Any]] = []
    for row, fields, md in checked:
        text = (md or "").strip().lower()
        if not text or len(text) < 40 or any(m in text for m in no_evidence_markers):
            dropped.append(
                {
                    "company": fields["name"],
                    "verdict": "drop_unrelated",
                    "reason": "no_independent_existence_evidence",
                }
            )
        else:
            kept.append(row)
    return kept, dropped


def filter_channel_pure_rows(
    rows: list[dict[str, Any]],
    *,
    family: str = "general",
    query: str = "",
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Filter landscape/pipeline rows (heuristic + DeepSeek when enabled)."""
    if not channel_purity_enabled():
        return list(rows), []
    return _llm_filter_rows(rows, family=family, query=query)


async def filter_channel_pure_rows_async(
    rows: list[dict[str, Any]],
    *,
    family: str = "general",
    query: str = "",
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Async wrapper: runs LLM purity in a worker thread, then optional web pass."""
    import asyncio

    if not channel_purity_enabled():
        return list(rows), []

    loop = asyncio.get_running_loop()
    with ThreadPoolExecutor(max_workers=1) as pool:
        kept, dropped = await loop.run_in_executor(
            pool,
            lambda: _llm_filter_rows(rows, family=family, query=query),
        )

    # Truly ambiguous rows (no local text signal either way) get one more
    # check via live Google AI evidence — independent of DeepSeek, so a
    # manufacturer never sneaks through just because the LLM pass was down.
    resolved_family = _resolve_family(family, query)
    kept, gai_dropped = await _gai_confirm_ambiguous(kept, family=resolved_family, query=query)
    if gai_dropped:
        dropped = list(dropped) + gai_dropped

    # Names structurally suspicious as fabricated placeholders (mirrors its
    # own domain, no distinctive brand word) get a real existence check
    # before anything is dropped — catches the "Semiconductores de Colombia"
    # / "Kenya Electronics" pattern that has no hardcoded-name-list defense.
    kept, existence_dropped = await _gai_confirm_existence(kept, family=resolved_family, query=query)
    if existence_dropped:
        dropped = list(dropped) + existence_dropped

    if not channel_purity_web_enabled() or not kept:
        return kept, dropped

    try:
        from vendor_intel.clients.searxng import searxng_search_any, searxng_urls
    except Exception:
        return kept, dropped

    urls = searxng_urls((os.getenv("SEARXNG_BASE_URL") or "http://127.0.0.1:8080").strip())
    if not urls:
        return kept, dropped

    sem = asyncio.Semaphore(8)
    final_kept: list[dict[str, Any]] = []
    extra_drop: list[dict[str, Any]] = []

    async def check(row: dict[str, Any]) -> tuple[dict[str, Any], PurityVerdict | None]:
        fields = _row_fields(row)
        async with sem:
            q = (
                f'"{fields["name"]}" (manufacturer OR OEM OR distributor OR wholesaler '
                f"OR dealer OR utility)"
            )
            try:
                hits, _ = await searxng_search_any(q, urls, max_results=8)
            except Exception:
                return row, None
            blob = "\n".join(f"{h.title} {h.snippet}" for h in (hits or []))
            oem = len(OEM_RE.findall(blob))
            ch = len(CHANNEL_RE.findall(blob))
            util = len(UTILITY_RE.findall(blob))
            if oem >= 2 and ch == 0:
                return row, PurityVerdict(False, f"web_oem={oem}", "drop_oem")
            if util >= 2 and ch == 0:
                return row, PurityVerdict(False, f"web_utility={util}", "drop_utility")
            return row, None

    results = await asyncio.gather(*[check(r) for r in kept])
    for row, v in results:
        if v is not None and not v.keep:
            fields = _row_fields(row)
            extra_drop.append(
                {"company": fields["name"], "verdict": v.label, "reason": v.reason}
            )
        else:
            final_kept.append(row)
    return final_kept, dropped + extra_drop
