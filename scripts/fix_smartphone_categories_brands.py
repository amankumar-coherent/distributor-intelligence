#!/usr/bin/env python3
"""Verify/fix Core Categories, Specialty Focus, Price Segment, Key Brands on smartphone FINAL.

Uses curated web-verified overrides + DuckDuckGo evidence + DeepSeek for residual rows.
Never invent personal emails/phones. Brands must be OEM names (or honest multi-brand NPD).
"""
from __future__ import annotations

import asyncio
import json
import os
import re
import shutil
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from dotenv import dotenv_values, load_dotenv
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side

load_dotenv(ROOT / ".env")
_cmi = Path(r"C:\Users\aman.kumar\Desktop\Cmi\agents\vendor-intel\.env")
if _cmi.exists():
    for k, v in (dotenv_values(_cmi) or {}).items():
        if v and k.startswith(("OPENAI_", "LLM_", "COMPILER_", "CLASSIFIER_", "MARKET_MAP_")):
            os.environ[k] = v

FINAL = (
    ROOT
    / "output/web_expand/global_smartphone_market_global/global_smartphone_market_global_FINAL.xlsx"
)
SHEET = "Companies"
MODEL = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip() or "deepseek-chat"
BASE = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
CONCURRENT = int(os.getenv("CAT_BRAND_CONCURRENT") or "5")
LIMIT = int(os.getenv("CAT_BRAND_LIMIT") or "0")
SKIP_LLM = (os.getenv("CAT_BRAND_SKIP_LLM") or "").strip() in ("1", "true", "yes")

DEFAULT_CORE = "Smartphones; Mobile Phones; Tablets; Accessories"
PRICE_OK = {
    "full range (entry–premium)",
    "full range (entry-premium)",
    "mass / premium mix",
    "premium",
    "mid-range / mass market",
    "multi-tier",
    "value / refurbished",
    "carrier / subsidized mix",
}
BRAND_ALLOW = [
    "Apple",
    "Samsung",
    "Xiaomi",
    "Huawei",
    "Honor",
    "OPPO",
    "Vivo",
    "Realme",
    "OnePlus",
    "Google",
    "Motorola",
    "Nokia",
    "Sony",
    "LG",
    "ASUS",
    "Nothing",
    "Fairphone",
    "ZTE",
    "Lenovo",
    "TCL",
    "Alcatel",
    "Tecno",
    "Infinix",
    "Itel",
    "HTC",
    "Sharp",
    "Microsoft",
    "Blackberry",
    "Cat",
    "Kyocera",
    "Panasonic",
    "Meizu",
    "Nubia",
    "iQOO",
    "POCO",
    "Redmi",
    "Galaxy",  # normalize away
]

# Web-verified / well-known channel facts (agent + public sources)
# tuple: core, specialty, price, brands
KNOWN: dict[str, tuple[str, str, str, str]] = {
    "amazon": (
        "Smartphones; Consumer Electronics; E-commerce",
        "Global online marketplace for smartphones & CE",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Google; Motorola; Xiaomi; OnePlus; OPPO; TCL",
    ),
    "best buy": (
        "Smartphones; Consumer Electronics; Accessories",
        "US big-box CE retail including carrier activations",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Google; Motorola; OnePlus",
    ),
    "currys": (
        "Smartphones; Consumer Electronics; Accessories",
        "UK/IE CE retail & online smartphone sales",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Google; Motorola; Nokia; Honor",
    ),
    "cdw": (
        "Smartphones; Mobile Devices; Enterprise Mobility",
        "B2B reseller of smartphones & mobility solutions",
        "Mass / Premium mix",
        "Apple; Samsung; Google; Motorola",
    ),
    "also holding": (
        "Smartphones; IT; Consumer Electronics",
        "European broadline IT/CE wholesale including mobiles",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Xiaomi; Google; Motorola; Honor",
    ),
    "ab sa": (
        "Smartphones; IT; Consumer Electronics",
        "Poland IT/CE distributor with mobile devices",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Xiaomi; Motorola; Huawei",
    ),
    "ingram micro": (
        "Smartphones; Mobility; IT Distribution",
        "Global IT & mobility distribution (Xvantage / Mobility)",
        "Full range (Entry–Premium)",
        "Samsung; Apple; Google; Motorola; OtterBox accessories channel",
    ),
    "td synnex": (
        "Smartphones; IT; Endpoint Devices",
        "Global broadline IT distributor including mobile endpoints",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Google; Motorola",
    ),
    "d&h distributing": (
        "Smartphones; IT; Consumer Electronics",
        "US/Canada IT & CE wholesale including smartphones",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Google; Motorola",
    ),
    "esprinet": (
        "Smartphones; IT; Consumer Electronics",
        "Southern European IT/CE wholesale (IT/ES)",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Xiaomi; Honor; Motorola",
    ),
    "computacenter": (
        "Smartphones; Enterprise Mobility; IT Services",
        "Enterprise IT reseller / integrator including mobiles",
        "Mass / Premium mix",
        "Apple; Samsung; Google",
    ),
    "bechtle": (
        "Smartphones; IT; Enterprise Mobility",
        "European IT reseller/distributor including mobiles",
        "Mass / Premium mix",
        "Apple; Samsung; Google",
    ),
    "dicker data": (
        "Smartphones; IT; Consumer Electronics",
        "ANZ IT distributor including mobile devices",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Google; Motorola",
    ),
    "elko group": (
        "Smartphones; IT; Consumer Electronics",
        "Baltic/Nordic IT/CE distributor",
        "Full range (Entry–Premium)",
        "Samsung; Xiaomi; Apple; Huawei; Honor",
    ),
    "brightstar": (
        "Smartphones; Mobile Devices; Device Lifecycle",
        "Global wireless device distribution & lifecycle (now Likewize heritage)",
        "Full range (Entry–Premium)",
        "Apple; Samsung; LG; Sony; Motorola",
    ),
    "exertis": (
        "Smartphones; Consumer Electronics; Accessories",
        "UK/EU CE & tech wholesale (Currys group)",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Google; Motorola; Nokia",
    ),
    "exertis ireland": (
        "Smartphones; Consumer Electronics; Accessories",
        "Ireland CE wholesale & distribution",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Google; Motorola",
    ),
    "shi international": (
        "Smartphones; Enterprise Mobility; IT",
        "Enterprise IT reseller including smartphones",
        "Mass / Premium mix",
        "Apple; Samsung; Google",
    ),
    "insight enterprises": (
        "Smartphones; Enterprise Mobility; IT",
        "Global enterprise IT solutions reseller",
        "Mass / Premium mix",
        "Apple; Samsung; Google",
    ),
    "softcat": (
        "Smartphones; IT; Enterprise Devices",
        "UK IT reseller including endpoint mobiles",
        "Mass / Premium mix",
        "Apple; Samsung; Google",
    ),
    "softchoice": (
        "Smartphones; IT; Enterprise Devices",
        "NA IT solutions reseller",
        "Mass / Premium mix",
        "Apple; Samsung; Google",
    ),
    "zones": (
        "Smartphones; IT; Endpoint Devices",
        "US IT solutions provider / reseller",
        "Mass / Premium mix",
        "Apple; Samsung; Google",
    ),
    "erajaya": (
        "Smartphones; Consumer Electronics; Retail",
        "Indonesia authorized smartphone retail & distribution",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Xiaomi; OPPO; Vivo; Huawei",
    ),
    "reliance jio": (
        "Smartphones; Mobile Services; Carrier Retail",
        "India carrier retail & JioPhone / partner handsets",
        "Carrier / Subsidized mix",
        "Samsung; Apple; Xiaomi; OPPO; Vivo; Realme; JioPhone",
    ),
    "fido solutions inc.": (
        "Smartphones; Carrier Retail; Mobile Plans",
        "Canadian mobile operator device retail (Rogers group)",
        "Carrier / Subsidized mix",
        "Apple; Samsung; Google; Motorola",
    ),
    "croma": (
        "Smartphones; Consumer Electronics; Retail",
        "India CE retail chain (Tata) including smartphones",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Xiaomi; OnePlus; OPPO; Vivo; Realme",
    ),
    "redington": (
        "Smartphones; IT; Consumer Electronics",
        "ME/Africa/India technology distribution",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Xiaomi; Huawei; OPPO",
    ),
    "rashi peripherals": (
        "Smartphones; IT; Consumer Electronics",
        "India IT/CE distributor (RPTech)",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Xiaomi; OPPO; Vivo",
    ),
    "asbis": (
        "Smartphones; IT; Consumer Electronics",
        "CEE/CIS IT & CE distribution",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Xiaomi; Huawei; Honor",
    ),
    "westcoast": (
        "Smartphones; IT; Consumer Electronics",
        "UK technology wholesale distributor",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Google; Motorola",
    ),
    "ban leong technologies": (
        "Smartphones; IT; Consumer Electronics",
        "Singapore/ASEAN IT & CE distributor",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Xiaomi; ASUS; Lenovo",
    ),
    "3g mobile": (
        "Smartphones; Mobile Phones; Wholesale",
        "US B2B wholesale of new/refurbished mobiles",
        "Value / Refurbished",
        "Apple; Samsung; Google; Motorola; LG",
    ),
    "arrow electronics": (
        "Smartphones; Components; Technology Distribution",
        "Global tech distributor; limited finished handset channel",
        "Multi-tier",
        "Multi-brand OEM / partner portfolio (varies by region)",
    ),
    "lazada": (
        "Smartphones; E-commerce; Consumer Electronics",
        "SEA e-commerce marketplace for smartphones",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Xiaomi; OPPO; Vivo; Realme; Huawei",
    ),
    "flipkart": (
        "Smartphones; E-commerce; Consumer Electronics",
        "India e-commerce marketplace for smartphones",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Xiaomi; OnePlus; OPPO; Vivo; Realme; Motorola",
    ),
    "jd.com": (
        "Smartphones; E-commerce; Consumer Electronics",
        "China online retail of smartphones & CE",
        "Full range (Entry–Premium)",
        "Apple; Huawei; Xiaomi; OPPO; Vivo; Honor; Samsung",
    ),
    "mediamarktsaturn": (
        "Smartphones; Consumer Electronics; Retail",
        "European CE retail group (MediaMarkt / Saturn)",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Google; Xiaomi; Honor; Motorola",
    ),
    "elkjøp": (
        "Smartphones; Consumer Electronics; Retail",
        "Nordic CE retail (Currys / Elkjøp Nordic)",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Google; Xiaomi; Motorola",
    ),
    "carrefour": (
        "Smartphones; Consumer Electronics; Hypermarket Retail",
        "Global hypermarket CE / smartphone retail",
        "Mass / Premium mix",
        "Samsung; Apple; Xiaomi; Motorola; TCL",
    ),
    "newegg": (
        "Smartphones; Consumer Electronics; E-commerce",
        "Online CE retail including smartphones",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Google; Motorola; OnePlus",
    ),
    "kogan": (
        "Smartphones; Consumer Electronics; E-commerce",
        "Australia online CE retail / private label + brands",
        "Mass / Premium mix",
        "Apple; Samsung; Google; Kogan; Motorola",
    ),
    "intcomex": (
        "Smartphones; IT; Consumer Electronics",
        "Latin America IT/CE wholesale distribution",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Xiaomi; Motorola; Huawei",
    ),
    "vst ecs": (
        "Smartphones; IT; Consumer Electronics",
        "Asia IT distributor including mobiles",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Xiaomi; Huawei; OPPO",
    ),
    "sis distribution": (
        "Smartphones; IT; Consumer Electronics",
        "Asia IT/CE distribution",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Xiaomi; ASUS",
    ),
    "logicom": (
        "Smartphones; IT; Telecom Distribution",
        "Eastern Med / ME technology distribution",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Xiaomi; Huawei",
    ),
    "axiom telecom": (
        "Smartphones; Telecom Retail; Distribution",
        "UAE/KSA telecom retail & handset distribution",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Huawei; Nokia; Lenovo",
    ),
    "emitac": (
        "Smartphones; IT; Consumer Electronics",
        "UAE technology distribution & services",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Huawei; Google",
    ),
    "t-gaia": (
        "Smartphones; Carrier Distribution; Telecom",
        "Japan mobile handset distribution to carriers/retail",
        "Carrier / Subsidized mix",
        "Apple; Samsung; Sony; Google; Sharp; Fujitsu",
    ),
    "komsa": (
        "Smartphones; Telecom Distribution; Accessories",
        "Germany/Europe mobile device wholesale",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Google; Xiaomi; Fairphone",
    ),
    "siewert & kau": (
        "Smartphones; IT; Consumer Electronics",
        "Germany IT/CE distributor",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Google; Xiaomi",
    ),
    "computer gross": (
        "Smartphones; IT; Consumer Electronics",
        "Italy IT wholesale distributor",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Xiaomi; Honor",
    ),
    "vinzeo": (
        "Smartphones; IT; Consumer Electronics",
        "Spain IT/CE wholesale (TD SYNNEX group)",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Xiaomi; Honor",
    ),
    "ed system": (
        "Smartphones; IT; Consumer Electronics",
        "CEE IT/CE distributor",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Xiaomi; Honor",
    ),
    "digiworld": (
        "Smartphones; IT; Consumer Electronics",
        "Vietnam ICT distributor including mobiles",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Xiaomi; OPPO; Vivo",
    ),
    "vision distribution pvt. ltd.": (
        "Smartphones; Mobile Distribution; India Channel",
        "India smartphone distribution (Vivo group channel)",
        "Mass / Premium mix",
        "Vivo; iQOO",
    ),
    "jumia": (
        "Smartphones; E-commerce; Consumer Electronics",
        "Africa e-commerce marketplace",
        "Mass / Premium mix",
        "Samsung; Tecno; Infinix; Itel; Xiaomi; Apple",
    ),
    "freedom mobile": (
        "Smartphones; Carrier Retail; Mobile Plans",
        "Canadian mobile operator device retail",
        "Carrier / Subsidized mix",
        "Apple; Samsung; Google; Motorola",
    ),
    "koodo mobile": (
        "Smartphones; Carrier Retail; Mobile Plans",
        "Canadian flanker-brand carrier retail (Telus)",
        "Carrier / Subsidized mix",
        "Apple; Samsung; Google; Motorola",
    ),
    "the source": (
        "Smartphones; Consumer Electronics; Retail",
        "Canadian CE retail (BCE) including smartphones",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Google; Motorola",
    ),
    "london drugs": (
        "Smartphones; Consumer Electronics; Retail",
        "Canadian pharmacy/CE retail including mobiles",
        "Mass / Premium mix",
        "Apple; Samsung; Google; Motorola",
    ),
    "staples canada": (
        "Smartphones; Office; Consumer Electronics",
        "Canadian office/CE retail channel",
        "Mass / Premium mix",
        "Apple; Samsung; Google",
    ),
    "visions electronics": (
        "Smartphones; Consumer Electronics; Retail",
        "Canadian CE specialty retail",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Google; Motorola",
    ),
    "mobile klinik": (
        "Smartphones; Repair; Accessories",
        "Canadian repair retail with device sales",
        "Value / Refurbished",
        "Apple; Samsung; Google",
    ),
    "tbooth wireless": (
        "Smartphones; Carrier Dealer; Retail",
        "Canadian wireless dealer / multi-carrier retail",
        "Carrier / Subsidized mix",
        "Apple; Samsung; Google; Motorola",
    ),
    "powerhouse alliance": (
        "Smartphones; Independent Retail Buying Group",
        "US independent wireless dealer alliance",
        "Carrier / Subsidized mix",
        "Apple; Samsung; Google; Motorola",
    ),
    "mr. checkout": (
        "Smartphones; Convenience Wholesale; Accessories",
        "US convenience/wholesale channel including phones accessories",
        "Mass / Premium mix",
        "Multi-brand smartphones & accessories (assortment varies)",
    ),
    "fry's electronics": (
        "Smartphones; Consumer Electronics; Retail",
        "Former US CE retail (legacy channel listing)",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Google; Motorola",
    ),
    "acquvia global": (
        "Smartphones; Premium Consumer Electronics; Wholesale",
        "Premium Apple-focused wholesale / grey-channel supply",
        "Premium",
        "Apple",
    ),
    "amir group fzco": (
        "Smartphones; Mobile Accessories; Wholesale",
        "Dubai mobile phones & gadget accessories wholesale",
        "Mass / Premium mix",
        "Samsung; Huawei",
    ),
    "abc mobile ltd": (
        "Smartphones; Mobile Phones; Accessories",
        "UK premium mobile phones & accessories distributor",
        "Premium",
        "Apple; Samsung; Google",
    ),
    "amgm universal center sl": (
        "Smartphones; Wholesale; Consumer Electronics",
        "Spain wholesale smartphones (Apple/Samsung/Xiaomi/Google)",
        "Full range (Entry–Premium)",
        "Apple; Samsung; Xiaomi; Google",
    ),
}


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


def _load_rows(path: Path) -> tuple[list[str], list[dict[str, str]], int]:
    wb = load_workbook(path, data_only=True)
    ws = wb[SHEET]
    rows = list(ws.iter_rows(values_only=True))
    header_idx = None
    headers: list[str] = []
    for i, r in enumerate(rows[:20]):
        vals = [str(c or "").strip() for c in r]
        if "Company" in vals:
            header_idx = i
            headers = vals
            break
    if header_idx is None:
        raise RuntimeError("No Company header")
    data: list[dict[str, str]] = []
    for r in rows[header_idx + 1 :]:
        d = {
            headers[j]: ("" if j >= len(r) or r[j] is None else str(r[j]).strip())
            for j in range(len(headers))
            if headers[j]
        }
        if d.get("Company"):
            data.append(d)
    return [h for h in headers if h], data, header_idx


def _is_firewall_core(v: str) -> bool:
    s = _norm(v)
    return any(x in s for x in ("firewall", "ngfw", "sase", "sd-wan", "cybersecurity"))


def _is_bad_brands(v: str) -> bool:
    s = (v or "").strip()
    sl = s.lower()
    if not s:
        return True
    if len(s) > 90:
        return True
    if "distributor" in sl or "wholesaler" in sl or "channel partner" in sl:
        return True
    if sl.startswith("provides ") or sl.startswith("specializes ") or sl.startswith("global "):
        return True
    if "established" in sl and "distributor" in sl:
        return True
    # no brand-like tokens
    if not any(re.search(rf"\b{re.escape(b)}\b", s, re.I) for b in BRAND_ALLOW if b != "Galaxy"):
        if "multi-brand" not in sl and "portfolio varies" not in sl and "not publicly" not in sl:
            if ";" not in s and "," not in s and len(s.split()) > 4:
                return True
            if len(s.split()) > 6:
                return True
    return False


def _is_bad_specialty(v: str) -> bool:
    s = (v or "").strip()
    if not s:
        return True
    if len(s) > 140:
        return True
    sl = s.lower()
    if sl in ("distributors of smartphones", "channel partners of smartphones"):
        return True
    if s.lower().startswith("confirmed exhibitor"):
        return True
    return False


def _is_bad_price(v: str) -> bool:
    s = _norm(v)
    if not s:
        return True
    if "enterprise / commercial" in s or s == "enterprise":
        return True
    return False


def _needs(row: dict[str, str]) -> bool:
    return (
        _is_firewall_core(row.get("Core Categories") or "")
        or _is_bad_brands(row.get("Key Brands Represented") or "")
        or _is_bad_specialty(row.get("Specialty Focus") or "")
        or _is_bad_price(row.get("Price Segment") or "")
    )


def _extract_brands(text: str) -> list[str]:
    found: list[str] = []
    t = text or ""
    aliases = {
        "iphone": "Apple",
        "galaxy": "Samsung",
        "pixel": "Google",
        "redmi": "Xiaomi",
        "poco": "Xiaomi",
        "oppo": "OPPO",
        "asus": "ASUS",
    }
    for b in BRAND_ALLOW:
        if b == "Galaxy":
            continue
        if re.search(rf"\b{re.escape(b)}\b", t, re.I):
            name = aliases.get(b.lower(), b)
            if name == "OPPO":
                name = "OPPO"
            elif name == "ASUS":
                name = "ASUS"
            if name not in found:
                found.append(name)
    for raw, canon in aliases.items():
        if re.search(rf"\b{re.escape(raw)}\b", t, re.I) and canon not in found:
            found.append(canon)
    # stable order preference
    prefer = [
        "Apple",
        "Samsung",
        "Google",
        "Xiaomi",
        "Huawei",
        "Honor",
        "OPPO",
        "Vivo",
        "Realme",
        "OnePlus",
        "Motorola",
        "Nokia",
        "Sony",
        "Tecno",
        "Infinix",
        "Itel",
    ]
    ordered = [b for b in prefer if b in found]
    ordered += [b for b in found if b not in ordered]
    return ordered[:8]


def _normalize_brands(v: str) -> str:
    s = (v or "").strip()
    if not s:
        return "Multi-brand smartphones (assortment varies / not fully public)"
    if _is_bad_brands(s):
        extracted = _extract_brands(s)
        if extracted:
            return "; ".join(extracted)
        return "Multi-brand smartphones (assortment varies / not fully public)"
    # clean separators
    parts = re.split(r"[;,/|]+", s)
    cleaned: list[str] = []
    for p in parts:
        p = p.strip()
        if not p or len(p) > 40:
            continue
        extracted = _extract_brands(p) or ([p] if len(p.split()) <= 3 else [])
        for e in extracted:
            if e not in cleaned:
                cleaned.append(e)
    if cleaned:
        return "; ".join(cleaned[:8])
    extracted = _extract_brands(s)
    if extracted:
        return "; ".join(extracted)
    if "multi-brand" in s.lower() or "portfolio varies" in s.lower() or "not publicly" in s.lower():
        return s[:120]
    return "Multi-brand smartphones (assortment varies / not fully public)"


def _normalize_price(v: str) -> str:
    s = _norm(v)
    mapping = [
        (r"carrier|subsid", "Carrier / Subsidized mix"),
        (r"refurb|used|value", "Value / Refurbished"),
        (r"premium only|^premium$", "Premium"),
        (r"mid.?range|mass market|entry.?level", "Mid-range / Mass market"),
        (r"mass\s*/\s*premium|mass.?premium", "Mass / Premium mix"),
        (r"multi.?tier|multi.?brand tier", "Multi-tier"),
        (r"full range|entry.?premium|value.?premium", "Full range (Entry–Premium)"),
        (r"enterprise", "Full range (Entry–Premium)"),
    ]
    for pat, label in mapping:
        if re.search(pat, s):
            return label
    if s in PRICE_OK:
        # restore canonical
        for cand in (
            "Full range (Entry–Premium)",
            "Mass / Premium mix",
            "Premium",
            "Mid-range / Mass market",
            "Multi-tier",
            "Value / Refurbished",
            "Carrier / Subsidized mix",
        ):
            if _norm(cand) == s.replace("-", "–"):
                return cand
    return "Full range (Entry–Premium)"


def _normalize_core(v: str, specialty: str = "", brands: str = "") -> str:
    if not _is_firewall_core(v) and v and len(v) < 120:
        # keep if already smartphone-ish
        sl = _norm(v)
        if any(x in sl for x in ("smartphone", "mobile phone", "handset", "consumer electronic", "tablet")):
            # standardize separators
            parts = [p.strip() for p in re.split(r"[;,]+", v) if p.strip()]
            return "; ".join(parts[:6]) if parts else DEFAULT_CORE
    blob = f"{specialty} {brands}".lower()
    parts = ["Smartphones", "Mobile Phones"]
    if "tablet" in blob:
        parts.append("Tablets")
    if "accessor" in blob or "repair" in blob:
        parts.append("Accessories")
    if "carrier" in blob or "plan" in blob or "wireless service" in blob:
        parts.append("Mobile Services")
    if "e-commerce" in blob or "marketplace" in blob:
        parts.append("E-commerce")
    if "consumer electronic" in blob or "ce " in blob:
        parts.append("Consumer Electronics")
    if len(parts) < 3:
        parts.extend(["Tablets", "Accessories"])
    # unique preserve order
    out: list[str] = []
    for p in parts:
        if p not in out:
            out.append(p)
    return "; ".join(out[:6])


def _normalize_specialty(v: str, row: dict[str, str]) -> str:
    s = (v or "").strip()
    if not _is_bad_specialty(s):
        # trim trailing fluff
        s = re.sub(r"\s+", " ", s)
        if len(s) > 120:
            s = s[:117].rsplit(" ", 1)[0] + "…"
        return s
    bt = _norm(row.get("Business Type") or "")
    hq = (row.get("Headquarters") or "").strip()
    geo = hq.split(",")[-1].strip() if hq else "regional"
    if "carrier" in bt:
        return f"{geo} carrier device retail & handset channel"
    if "e-commerce" in bt or "retailer / e-commerce" in bt:
        return f"Online smartphone / CE retail ({geo})"
    if "retail" in bt:
        return f"Consumer electronics retail including smartphones ({geo})"
    if "reseller" in bt:
        return f"B2B reseller of smartphones & mobility ({geo})"
    if "wholesale" in bt or "distributor" in bt or "importer" in bt:
        return f"Smartphone wholesale / B2B handset distribution ({geo})"
    return f"Smartphone channel partner ({geo})"


async def _web_evidence(name: str, website: str) -> str:
    host = ""
    if website:
        host = re.sub(r"^https?://", "", website.strip(), flags=re.I).split("/")[0]
        host = host.replace("www.", "")
    queries = [
        f'"{name}" smartphone OR mobile OR handset distributor OR retailer brands OR Apple OR Samsung',
        f'"{name}" authorized distributor OR wholesale phones OR "mobile phones" brands',
    ]
    if host:
        queries.insert(0, f'site:{host} brands OR partners OR Apple OR Samsung OR Xiaomi')
        queries.insert(1, f'"{name}" {host} smartphones brands portfolio')

    parts: list[str] = []
    try:
        from ddgs import DDGS

        def _one(q: str) -> list[str]:
            lines: list[str] = []
            try:
                hits = list(DDGS().text(q, max_results=5)) or []
            except Exception:
                hits = []
            for h in hits:
                lines.append(
                    f"- {h.get('title', '')}: {str(h.get('body', ''))[:240]} ({h.get('href', '')})"
                )
            return lines

        for q in queries[:3]:
            try:
                lines = await asyncio.wait_for(asyncio.to_thread(_one, q), timeout=22)
            except Exception:
                lines = []
            if lines:
                parts.append(f"Query: {q}\n" + "\n".join(lines))
    except Exception as err:
        parts.append(f"(ddgs: {err})")

    if not any(p.startswith("Query:") for p in parts):
        try:
            from vendor_intel.clients.search_router import FreeSearchRouter
            from vendor_intel.config import Settings
            from vendor_intel.enrichment.gap_fill.sources import _search_snippets

            router = FreeSearchRouter(Settings.load())
            for q in queries[:2]:
                try:
                    snip = await asyncio.wait_for(
                        _search_snippets(router, q, max_hits=5), timeout=25
                    )
                except Exception:
                    snip = ""
                if snip:
                    parts.append(f"Query: {q}\n{snip}")
        except Exception as err:
            parts.append(f"(router: {err})")
    return "\n\n".join(parts)[:7000] or "(no web hits)"


def _client():
    from openai import OpenAI

    key = (os.getenv("OPENAI_API_KEY") or "").strip()
    if not key:
        raise RuntimeError("OPENAI_API_KEY missing")
    return OpenAI(api_key=key, base_url=BASE)


def _llm(client: Any, row: dict[str, str], evidence: str) -> dict[str, Any]:
    name = row.get("Company") or ""
    system = (
        "You classify smartphone CHANNEL companies (distributors, wholesalers, retailers, "
        "carriers, resellers). Return JSON only. Use WEB EVIDENCE; do not invent brand "
        "authorizations. If brands unclear, say multi-brand assortment varies.\n"
        "Core Categories: short semicolon list like "
        "'Smartphones; Mobile Phones; Tablets; Accessories'. NEVER firewall/NGFW/SASE.\n"
        "Specialty Focus: one short phrase (<=110 chars) describing channel niche.\n"
        "Price Segment: one of "
        "[Full range (Entry–Premium) | Mass / Premium mix | Premium | "
        "Mid-range / Mass market | Multi-tier | Value / Refurbished | "
        "Carrier / Subsidized mix]. NEVER 'Enterprise / Commercial'.\n"
        "Key Brands Represented: semicolon-separated OEM brand names only "
        "(Apple; Samsung; …). No sentences. If unknown: "
        "'Multi-brand smartphones (assortment varies / not fully public)'."
    )
    user = f"""Company: {name}
Website: {row.get('Website') or ''}
HQ: {row.get('Headquarters') or ''}
Business Type: {row.get('Business Type') or ''}
Current Core: {row.get('Core Categories') or ''}
Current Specialty: {(row.get('Specialty Focus') or '')[:200]}
Current Brands: {(row.get('Key Brands Represented') or '')[:200]}
Summary: {(row.get('Summary') or '')[:280]}

WEB EVIDENCE:
{evidence}

Return JSON:
{{
  "core_categories": "Smartphones; ...",
  "specialty_focus": "short phrase",
  "price_segment": "one allowed value",
  "key_brands": "Apple; Samsung; ...",
  "confidence": "high|medium|low",
  "notes": "short"
}}
"""
    resp = client.chat.completions.create(
        model=MODEL,
        temperature=0.1,
        max_tokens=450,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
    )
    raw = resp.choices[0].message.content or "{}"
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        m = re.search(r"\{[\s\S]*\}", raw)
        data = json.loads(m.group(0)) if m else {}
    return data if isinstance(data, dict) else {}


def _heuristic_from_row(row: dict[str, str], evidence: str = "") -> dict[str, str]:
    blob = " ".join(
        [
            row.get("Specialty Focus") or "",
            row.get("Key Brands Represented") or "",
            row.get("Summary") or "",
            evidence,
        ]
    )
    brands = _extract_brands(blob)
    brand_str = (
        "; ".join(brands)
        if brands
        else "Multi-brand smartphones (assortment varies / not fully public)"
    )
    specialty = _normalize_specialty(row.get("Specialty Focus") or "", row)
    # refine specialty from evidence keywords
    el = evidence.lower()
    if "refurb" in el or "used phone" in el:
        specialty = specialty if "refurb" in specialty.lower() else (
            specialty.rstrip(".") + " — new & refurbished"
        )[:120]
    price = "Full range (Entry–Premium)"
    bt = _norm(row.get("Business Type") or "")
    if "carrier" in bt:
        price = "Carrier / Subsidized mix"
    elif "refurb" in el or "used" in specialty.lower():
        price = "Value / Refurbished"
    elif brands and set(brands) <= {"Apple"}:
        price = "Premium"
    elif "retail" in bt and "e-commerce" not in bt:
        price = "Mass / Premium mix"
    core = _normalize_core("", specialty, brand_str)
    return {
        "Core Categories": core,
        "Specialty Focus": specialty,
        "Price Segment": price,
        "Key Brands Represented": brand_str,
    }


def _apply_known_or_llm(
    row: dict[str, str], data: dict[str, Any] | None, evidence: str
) -> dict[str, str]:
    name = _norm(row.get("Company") or "")
    out = dict(row)
    if name in KNOWN:
        core, spec, price, brands = KNOWN[name]
        out["Core Categories"] = core
        out["Specialty Focus"] = spec
        out["Price Segment"] = price
        out["Key Brands Represented"] = brands
        src = out.get("Data Sources") or ""
        if "web_cat_brands_known" not in src:
            out["Data Sources"] = (src + "; web_cat_brands_known").strip("; ")
        return out

    heur = _heuristic_from_row(row, evidence)
    if data:
        core = str(data.get("core_categories") or heur["Core Categories"])
        spec = str(data.get("specialty_focus") or heur["Specialty Focus"])
        price = str(data.get("price_segment") or heur["Price Segment"])
        brands = str(data.get("key_brands") or heur["Key Brands Represented"])
    else:
        core, spec, price, brands = (
            heur["Core Categories"],
            heur["Specialty Focus"],
            heur["Price Segment"],
            heur["Key Brands Represented"],
        )

    out["Core Categories"] = _normalize_core(core, spec, brands)
    out["Specialty Focus"] = _normalize_specialty(spec, row)
    out["Price Segment"] = _normalize_price(price)
    out["Key Brands Represented"] = _normalize_brands(brands)
    # merge extracted brands from evidence if still generic
    if "assortment varies" in out["Key Brands Represented"].lower():
        extra = _extract_brands(evidence + " " + (row.get("Summary") or ""))
        if extra:
            out["Key Brands Represented"] = "; ".join(extra)

    src = out.get("Data Sources") or ""
    tag = "web_cat_brands_llm" if data else "web_cat_brands_heur"
    if tag not in src:
        out["Data Sources"] = (src + f"; {tag}").strip("; ")
    return out


async def _process_one(sem: asyncio.Semaphore, client: Any | None, row: dict[str, str]) -> dict[str, str]:
    name = _norm(row.get("Company") or "")
    if name in KNOWN:
        return _apply_known_or_llm(row, None, "")

    async with sem:
        evidence = await _web_evidence(row.get("Company") or "", row.get("Website") or "")
        data = None
        if client and not SKIP_LLM:
            try:
                data = await asyncio.to_thread(_llm, client, row, evidence)
            except Exception as err:
                data = {"notes": f"llm_err:{err}"}
        return _apply_known_or_llm(row, data, evidence)


def _write_excel(headers: list[str], data: list[dict[str, str]], header_idx: int) -> Path:
    path = FINAL
    ts = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    bak = path.with_suffix(path.suffix + f".bak_cat_brands_{ts}")
    shutil.copy2(path, bak)

    wb = load_workbook(path)
    ws = wb[SHEET]
    # map header -> col
    hr = header_idx + 1
    col = {ws.cell(hr, c).value: c for c in range(1, ws.max_column + 1) if ws.cell(hr, c).value}
    fields = [
        "Core Categories",
        "Specialty Focus",
        "Price Segment",
        "Key Brands Represented",
        "Data Sources",
    ]
    wrap = Alignment(wrap_text=True, vertical="top")
    thin = Border(
        left=Side(style="thin", color="DDDDDD"),
        right=Side(style="thin", color="DDDDDD"),
        top=Side(style="thin", color="DDDDDD"),
        bottom=Side(style="thin", color="DDDDDD"),
    )
    # build name->row map from sheet
    name_to_r: dict[str, int] = {}
    for r in range(hr + 1, ws.max_row + 1):
        n = str(ws.cell(r, col["Company"]).value or "").strip()
        if n:
            name_to_r[_norm(n)] = r

    updated = 0
    for d in data:
        r = name_to_r.get(_norm(d.get("Company") or ""))
        if not r:
            continue
        for f in fields:
            if f not in col:
                continue
            ws.cell(r, col[f]).value = d.get(f) or ""
            ws.cell(r, col[f]).alignment = wrap
            ws.cell(r, col[f]).border = thin
        updated += 1

    wb.save(path)
    print(f"updated_rows={updated} backup={bak.name}")
    return bak


async def main() -> None:
    headers, data, header_idx = _load_rows(FINAL)
    print(f"loaded={len(data)} needs_fix={sum(1 for r in data if _needs(r))} known={len(KNOWN)}")

    work = data
    if LIMIT > 0:
        work = data[:LIMIT]
        print(f"LIMIT={LIMIT}")

    client = None if SKIP_LLM else _client()
    sem = asyncio.Semaphore(CONCURRENT)
    results: list[dict[str, str]] = []
    # process all (KNOWN fast-path inside)
    tasks = [_process_one(sem, client, row) for row in work]
    done = 0
    for coro in asyncio.as_completed(tasks):
        row = await coro
        results.append(row)
        done += 1
        if done % 25 == 0 or done == len(tasks):
            print(f"progress {done}/{len(tasks)}")

    # merge back by company name for full list if LIMIT
    by_name = {_norm(r["Company"]): r for r in results}
    merged = []
    for r in data:
        n = _norm(r["Company"])
        merged.append(by_name.get(n, r) if n in by_name else r)

    # force-fix any remaining firewall/enterprise leftovers deterministically
    fixed_extra = 0
    for r in merged:
        before = (
            r.get("Core Categories"),
            r.get("Specialty Focus"),
            r.get("Price Segment"),
            r.get("Key Brands Represented"),
        )
        if _is_firewall_core(r.get("Core Categories") or ""):
            r["Core Categories"] = _normalize_core("", r.get("Specialty Focus") or "", r.get("Key Brands Represented") or "")
        if _is_bad_price(r.get("Price Segment") or ""):
            r["Price Segment"] = _normalize_price(r.get("Price Segment") or "")
        if _is_bad_brands(r.get("Key Brands Represented") or ""):
            r["Key Brands Represented"] = _normalize_brands(r.get("Key Brands Represented") or "")
        if _is_bad_specialty(r.get("Specialty Focus") or ""):
            r["Specialty Focus"] = _normalize_specialty(r.get("Specialty Focus") or "", r)
        after = (
            r.get("Core Categories"),
            r.get("Specialty Focus"),
            r.get("Price Segment"),
            r.get("Key Brands Represented"),
        )
        if before != after:
            fixed_extra += 1

    _write_excel(headers, merged, header_idx)

    # audit summary
    from collections import Counter

    cc = Counter(r.get("Core Categories") for r in merged)
    ps = Counter(r.get("Price Segment") for r in merged)
    bad_cc = sum(1 for r in merged if _is_firewall_core(r.get("Core Categories") or ""))
    bad_ps = sum(1 for r in merged if _is_bad_price(r.get("Price Segment") or ""))
    bad_kb = sum(1 for r in merged if _is_bad_brands(r.get("Key Brands Represented") or ""))
    audit = {
        "ts": datetime.now(timezone.utc).isoformat(),
        "total": len(merged),
        "fixed_extra_pass": fixed_extra,
        "remaining_firewall_core": bad_cc,
        "remaining_bad_price": bad_ps,
        "remaining_bad_brands": bad_kb,
        "top_core": cc.most_common(8),
        "price_dist": ps.most_common(),
        "sample": [
            {
                "Company": r["Company"],
                "Core Categories": r.get("Core Categories"),
                "Specialty Focus": r.get("Specialty Focus"),
                "Price Segment": r.get("Price Segment"),
                "Key Brands Represented": r.get("Key Brands Represented"),
            }
            for r in merged[:12]
        ],
    }
    out = FINAL.parent / "category_brands_audit.json"
    out.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps({k: audit[k] for k in audit if k != "sample"}, indent=2, ensure_ascii=False))
    print("audit", out)


if __name__ == "__main__":
    asyncio.run(main())
