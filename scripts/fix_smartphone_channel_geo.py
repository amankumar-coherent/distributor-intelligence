#!/usr/bin/env python3
"""Verify/fix Retail Chains, Distribution Type, Regions Served, Regional Extensions,
Warehouse / Logistics, Delivery / Storage on smartphone FINAL.

Uses curated web-verified overrides + HQ/cities heuristics + optional web+DeepSeek
for seed_file / ambiguous region rows.
"""
from __future__ import annotations

import asyncio
import json
import os
import re
import shutil
import sys
from collections import Counter
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
CONCURRENT = int(os.getenv("CHANNEL_GEO_CONCURRENT") or "4")
LIMIT = int(os.getenv("CHANNEL_GEO_LIMIT") or "0")
SKIP_LLM = (os.getenv("CHANNEL_GEO_SKIP_LLM") or "").strip().lower() in ("1", "true", "yes")

FIELDS = [
    "Retail Chains",
    "Distribution Type",
    "Regions Served",
    "Regional Extensions",
    "Warehouse / Logistics",
    "Delivery / Storage",
]

# Web-verified / well-known: retail, dist, regions, extensions, warehouse, delivery
KNOWN: dict[str, dict[str, str]] = {
    "amazon": {
        "Retail Chains": "No",
        "Distribution Type": "Online retail / marketplace",
        "Regions Served": "Global (marketplace; primary US + international stores)",
        "Regional Extensions": "United States; Europe; Japan; India; Australia; UAE; Brazil",
        "Warehouse / Logistics": "Yes — global fulfillment network (FBA / Amazon logistics)",
        "Delivery / Storage": "Standard e-commerce fulfillment + last-mile delivery",
    },
    "best buy": {
        "Retail Chains": "Yes",
        "Distribution Type": "Omnichannel",
        "Regions Served": "United States; Canada",
        "Regional Extensions": "US nationwide stores; Best Buy Canada provinces",
        "Warehouse / Logistics": "Yes — national DC + store inventory network",
        "Delivery / Storage": "Store pickup + home delivery / Geek Squad fulfillment",
    },
    "currys": {
        "Retail Chains": "Yes",
        "Distribution Type": "Omnichannel",
        "Regions Served": "United Kingdom; Ireland",
        "Regional Extensions": "UK & Ireland Currys / PC World stores + online",
        "Warehouse / Logistics": "Yes — UK/IE retail logistics (Exertis group adjacency)",
        "Delivery / Storage": "Store + home delivery / installation",
    },
    "mediamarktsaturn": {
        "Retail Chains": "Yes",
        "Distribution Type": "Omnichannel",
        "Regions Served": "Europe (Germany; Austria; Belgium; Hungary; Italy; Luxembourg; Netherlands; Poland; Spain; Switzerland; Turkey)",
        "Regional Extensions": "MediaMarkt / Saturn / MediaWorld store network (~11 countries)",
        "Warehouse / Logistics": "Yes — European retail DC & store replenishment",
        "Delivery / Storage": "Store + online omnichannel fulfillment",
    },
    "elkjøp": {
        "Retail Chains": "Yes",
        "Distribution Type": "Omnichannel",
        "Regions Served": "Nordics (Norway; Sweden; Denmark; Finland)",
        "Regional Extensions": "Elkjøp Nordic retail footprint",
        "Warehouse / Logistics": "Yes — Nordic retail logistics",
        "Delivery / Storage": "Store + home delivery",
    },
    "carrefour": {
        "Retail Chains": "Yes",
        "Distribution Type": "Omnichannel",
        "Regions Served": "Europe; Latin America; Middle East; Africa; Asia (selected markets)",
        "Regional Extensions": "Hypermarket / supermarket CE aisles by country",
        "Warehouse / Logistics": "Yes — retail grocery/CE logistics",
        "Delivery / Storage": "Store + e-commerce fulfillment",
    },
    "ingram micro": {
        "Retail Chains": "No",
        "Distribution Type": "Broadline Distributor",
        "Regions Served": "Global (NA; EMEA; Asia Pacific; Latin America)",
        "Regional Extensions": "~57 countries ops; 160+ countries served via logistics centers",
        "Warehouse / Logistics": "Yes — global logistics & service centers (~189 sites)",
        "Delivery / Storage": "B2B wholesale fulfillment / Xvantage channel delivery",
    },
    "td synnex": {
        "Retail Chains": "No",
        "Distribution Type": "Broadline Distributor",
        "Regions Served": "Global (Americas; Europe; Asia Pacific)",
        "Regional Extensions": "US; Canada; Europe; Asia Pacific hubs",
        "Warehouse / Logistics": "Yes — global IT distribution warehouses",
        "Delivery / Storage": "B2B wholesale / drop-ship fulfillment",
    },
    "d&h distributing": {
        "Retail Chains": "No",
        "Distribution Type": "Broadline Distributor",
        "Regions Served": "United States; Canada",
        "Regional Extensions": "US & Canada distribution hubs",
        "Warehouse / Logistics": "Yes — NA distribution centers",
        "Delivery / Storage": "B2B wholesale fulfillment",
    },
    "also holding": {
        "Retail Chains": "No",
        "Distribution Type": "Broadline Distributor",
        "Regions Served": "Europe",
        "Regional Extensions": "Switzerland HQ; Western & Central Europe markets",
        "Warehouse / Logistics": "Yes — European wholesale logistics",
        "Delivery / Storage": "B2B channel fulfillment",
    },
    "ab sa": {
        "Retail Chains": "No",
        "Distribution Type": "Broadline Distributor",
        "Regions Served": "Poland; Central & Eastern Europe",
        "Regional Extensions": "Poland HQ; CEE export markets",
        "Warehouse / Logistics": "Yes — Poland / CEE DCs",
        "Delivery / Storage": "B2B wholesale fulfillment",
    },
    "esprinet": {
        "Retail Chains": "No",
        "Distribution Type": "Broadline Distributor",
        "Regions Served": "Italy; Spain; Portugal",
        "Regional Extensions": "Southern Europe IT/CE wholesale",
        "Warehouse / Logistics": "Yes — IT/ES distribution logistics",
        "Delivery / Storage": "B2B wholesale fulfillment",
    },
    "asbis": {
        "Retail Chains": "No",
        "Distribution Type": "Broadline Distributor",
        "Regions Served": "CEE; CIS; Middle East (selected)",
        "Regional Extensions": "Cyprus / CEE hubs; CIS markets",
        "Warehouse / Logistics": "Yes — regional DCs",
        "Delivery / Storage": "B2B wholesale fulfillment",
    },
    "elko group": {
        "Retail Chains": "No",
        "Distribution Type": "Broadline Distributor",
        "Regions Served": "Nordics; Baltics; CEE",
        "Regional Extensions": "Latvia HQ; Nordic/Baltic coverage",
        "Warehouse / Logistics": "Yes — Baltic/Nordic warehouses",
        "Delivery / Storage": "B2B wholesale fulfillment",
    },
    "dicker data": {
        "Retail Chains": "No",
        "Distribution Type": "Broadline Distributor",
        "Regions Served": "Australia; New Zealand",
        "Regional Extensions": "ANZ distribution hubs",
        "Warehouse / Logistics": "Yes — ANZ warehouses",
        "Delivery / Storage": "B2B wholesale fulfillment",
    },
    "redington": {
        "Retail Chains": "No",
        "Distribution Type": "Broadline Distributor",
        "Regions Served": "India; Middle East; Africa",
        "Regional Extensions": "India; GCC; African markets",
        "Warehouse / Logistics": "Yes — regional distribution logistics",
        "Delivery / Storage": "B2B wholesale fulfillment",
    },
    "rashi peripherals": {
        "Retail Chains": "No",
        "Distribution Type": "Broadline Distributor",
        "Regions Served": "India",
        "Regional Extensions": "India pan-national channel",
        "Warehouse / Logistics": "Yes — India DCs",
        "Delivery / Storage": "B2B wholesale fulfillment",
    },
    "vst ecs": {
        "Retail Chains": "No",
        "Distribution Type": "Broadline Distributor",
        "Regions Served": "Asia Pacific",
        "Regional Extensions": "Singapore / SEA / Asia hubs",
        "Warehouse / Logistics": "Yes — Asia distribution logistics",
        "Delivery / Storage": "B2B wholesale fulfillment",
    },
    "ban leong technologies": {
        "Retail Chains": "No",
        "Distribution Type": "Broadline Distributor",
        "Regions Served": "Singapore; Southeast Asia",
        "Regional Extensions": "Singapore HQ; ASEAN channel",
        "Warehouse / Logistics": "Yes — Singapore / SEA logistics",
        "Delivery / Storage": "B2B wholesale fulfillment",
    },
    "westcoast": {
        "Retail Chains": "No",
        "Distribution Type": "Broadline Distributor",
        "Regions Served": "United Kingdom",
        "Regional Extensions": "UK wholesale technology channel",
        "Warehouse / Logistics": "Yes — UK warehouses",
        "Delivery / Storage": "B2B wholesale fulfillment",
    },
    "exertis": {
        "Retail Chains": "No",
        "Distribution Type": "Specialty Distributor",
        "Regions Served": "United Kingdom; Europe (selected)",
        "Regional Extensions": "UK CE/tech wholesale (Currys group)",
        "Warehouse / Logistics": "Yes — UK/EU wholesale logistics",
        "Delivery / Storage": "B2B wholesale fulfillment",
    },
    "exertis ireland": {
        "Retail Chains": "No",
        "Distribution Type": "Specialty Distributor",
        "Regions Served": "Ireland",
        "Regional Extensions": "Ireland CE wholesale",
        "Warehouse / Logistics": "Yes — Ireland logistics",
        "Delivery / Storage": "B2B wholesale fulfillment",
    },
    "brightstar": {
        "Retail Chains": "No",
        "Distribution Type": "Specialty Distributor",
        "Regions Served": "Global (mobility / device lifecycle markets)",
        "Regional Extensions": "Americas; EMEA; APAC device supply-chain hubs",
        "Warehouse / Logistics": "Yes — device lifecycle / reverse logistics",
        "Delivery / Storage": "Carrier & retail device fulfillment / repair flow",
    },
    "komsa": {
        "Retail Chains": "No",
        "Distribution Type": "Value-Added Distributor",
        "Regions Served": "Germany; Central & Eastern Europe",
        "Regional Extensions": "Germany HQ; CEE partner channel",
        "Warehouse / Logistics": "Yes — telecom/IT logistics & repair",
        "Delivery / Storage": "B2B + DaaS device fulfillment",
    },
    "intcomex": {
        "Retail Chains": "No",
        "Distribution Type": "Broadline Distributor",
        "Regions Served": "Latin America",
        "Regional Extensions": "Miami / LatAm country ops",
        "Warehouse / Logistics": "Yes — LatAm regional DCs",
        "Delivery / Storage": "B2B wholesale fulfillment",
    },
    "computer gross": {
        "Retail Chains": "No",
        "Distribution Type": "Broadline Distributor",
        "Regions Served": "Italy",
        "Regional Extensions": "Italy IT wholesale",
        "Warehouse / Logistics": "Yes — Italy DC network",
        "Delivery / Storage": "B2B wholesale fulfillment",
    },
    "vinzeo": {
        "Retail Chains": "No",
        "Distribution Type": "Broadline Distributor",
        "Regions Served": "Spain",
        "Regional Extensions": "Spain (TD SYNNEX group)",
        "Warehouse / Logistics": "Yes — Spain wholesale logistics",
        "Delivery / Storage": "B2B wholesale fulfillment",
    },
    "ed system": {
        "Retail Chains": "No",
        "Distribution Type": "Broadline Distributor",
        "Regions Served": "Central & Eastern Europe",
        "Regional Extensions": "CEE IT/CE wholesale",
        "Warehouse / Logistics": "Yes — regional DCs",
        "Delivery / Storage": "B2B wholesale fulfillment",
    },
    "siewert & kau": {
        "Retail Chains": "No",
        "Distribution Type": "Broadline Distributor",
        "Regions Served": "Germany",
        "Regional Extensions": "Germany IT/CE wholesale",
        "Warehouse / Logistics": "Yes — Germany warehouses",
        "Delivery / Storage": "B2B wholesale fulfillment",
    },
    "logicom": {
        "Retail Chains": "No",
        "Distribution Type": "Broadline Distributor",
        "Regions Served": "Eastern Mediterranean; Middle East",
        "Regional Extensions": "Cyprus HQ; ME/EM markets",
        "Warehouse / Logistics": "Yes — regional logistics",
        "Delivery / Storage": "B2B wholesale fulfillment",
    },
    "sis distribution": {
        "Retail Chains": "No",
        "Distribution Type": "Broadline Distributor",
        "Regions Served": "Asia Pacific",
        "Regional Extensions": "Hong Kong / Asia channel",
        "Warehouse / Logistics": "Yes — Asia distribution logistics",
        "Delivery / Storage": "B2B wholesale fulfillment",
    },
    "digiworld": {
        "Retail Chains": "No",
        "Distribution Type": "Broadline Distributor",
        "Regions Served": "Vietnam; Southeast Asia (selected)",
        "Regional Extensions": "Vietnam ICT distribution",
        "Warehouse / Logistics": "Yes — Vietnam DCs",
        "Delivery / Storage": "B2B wholesale fulfillment",
    },
    "cdw": {
        "Retail Chains": "No",
        "Distribution Type": "Reseller / Distributor",
        "Regions Served": "United States; United Kingdom; Canada",
        "Regional Extensions": "US HQ; UK / Canada ops",
        "Warehouse / Logistics": "Yes — configuration & distribution centers",
        "Delivery / Storage": "B2B configured fulfillment",
    },
    "computacenter": {
        "Retail Chains": "No",
        "Distribution Type": "Reseller / Systems Integrator",
        "Regions Served": "United Kingdom; Europe; North America (selected)",
        "Regional Extensions": "UK HQ; EU / NA enterprise delivery",
        "Warehouse / Logistics": "Yes — integration & logistics centers",
        "Delivery / Storage": "Enterprise staged deployment / fulfillment",
    },
    "bechtle": {
        "Retail Chains": "No",
        "Distribution Type": "Reseller / Systems Integrator",
        "Regions Served": "Germany; Europe",
        "Regional Extensions": "Germany HQ; European subsidiaries",
        "Warehouse / Logistics": "Yes — European logistics",
        "Delivery / Storage": "Enterprise fulfillment",
    },
    "insight enterprises": {
        "Retail Chains": "No",
        "Distribution Type": "Reseller / Distributor",
        "Regions Served": "Global (Americas; EMEA; APAC)",
        "Regional Extensions": "US HQ; international ops",
        "Warehouse / Logistics": "Yes — global fulfillment",
        "Delivery / Storage": "B2B solutions fulfillment",
    },
    "shi international": {
        "Retail Chains": "No",
        "Distribution Type": "Reseller / Distributor",
        "Regions Served": "United States; international (selected)",
        "Regional Extensions": "US enterprise channel",
        "Warehouse / Logistics": "Yes — US logistics",
        "Delivery / Storage": "B2B fulfillment",
    },
    "softcat": {
        "Retail Chains": "No",
        "Distribution Type": "Reseller / Distributor",
        "Regions Served": "United Kingdom",
        "Regional Extensions": "UK enterprise IT channel",
        "Warehouse / Logistics": "Partial — partner / vendor logistics heavy",
        "Delivery / Storage": "B2B project fulfillment",
    },
    "softchoice": {
        "Retail Chains": "No",
        "Distribution Type": "Reseller / Distributor",
        "Regions Served": "Canada; United States",
        "Regional Extensions": "Canada / US enterprise channel",
        "Warehouse / Logistics": "Partial — partner logistics",
        "Delivery / Storage": "B2B fulfillment",
    },
    "zones": {
        "Retail Chains": "No",
        "Distribution Type": "Reseller / Distributor",
        "Regions Served": "United States",
        "Regional Extensions": "US enterprise / SMB channel",
        "Warehouse / Logistics": "Partial — partner logistics",
        "Delivery / Storage": "B2B fulfillment",
    },
    "arrow electronics": {
        "Retail Chains": "No",
        "Distribution Type": "Broadline Distributor",
        "Regions Served": "Global",
        "Regional Extensions": "Americas; EMEA; APAC",
        "Warehouse / Logistics": "Yes — global component/tech logistics",
        "Delivery / Storage": "B2B wholesale fulfillment",
    },
    "erajaya": {
        "Retail Chains": "Yes",
        "Distribution Type": "Retail / Wholesale",
        "Regions Served": "Indonesia; Singapore; Malaysia (selected)",
        "Regional Extensions": "Indonesia retail brands (Erafone; iBox) + distribution centers",
        "Warehouse / Logistics": "Yes — Indonesia DCs + retail replenishment",
        "Delivery / Storage": "Store + wholesale channel fulfillment",
    },
    "croma": {
        "Retail Chains": "Yes",
        "Distribution Type": "Omnichannel",
        "Regions Served": "India",
        "Regional Extensions": "India Croma store network (Tata)",
        "Warehouse / Logistics": "Yes — India retail logistics",
        "Delivery / Storage": "Store + home delivery",
    },
    "reliance jio": {
        "Retail Chains": "Yes",
        "Distribution Type": "Carrier retail",
        "Regions Served": "India",
        "Regional Extensions": "India Jio / Reliance Retail device channel",
        "Warehouse / Logistics": "Yes — India telecom retail logistics",
        "Delivery / Storage": "Store + online device fulfillment",
    },
    "fido solutions inc.": {
        "Retail Chains": "Yes",
        "Distribution Type": "Carrier retail",
        "Regions Served": "Canada",
        "Regional Extensions": "Canada Rogers/Fido dealer & corporate stores",
        "Warehouse / Logistics": "Yes — carrier device logistics",
        "Delivery / Storage": "Store activation + shipping",
    },
    "freedom mobile": {
        "Retail Chains": "Yes",
        "Distribution Type": "Carrier retail",
        "Regions Served": "Canada",
        "Regional Extensions": "Canada Freedom Mobile stores",
        "Warehouse / Logistics": "Yes — carrier device logistics",
        "Delivery / Storage": "Store activation + shipping",
    },
    "koodo mobile": {
        "Retail Chains": "Yes",
        "Distribution Type": "Carrier retail",
        "Regions Served": "Canada",
        "Regional Extensions": "Canada Koodo / Telus channel",
        "Warehouse / Logistics": "Yes — carrier device logistics",
        "Delivery / Storage": "Store activation + shipping",
    },
    "the source": {
        "Retail Chains": "Yes",
        "Distribution Type": "Omnichannel",
        "Regions Served": "Canada",
        "Regional Extensions": "Canada The Source / Best Buy Express footprint",
        "Warehouse / Logistics": "Yes — retail logistics",
        "Delivery / Storage": "Store + online fulfillment",
    },
    "london drugs": {
        "Retail Chains": "Yes",
        "Distribution Type": "Omnichannel",
        "Regions Served": "Canada (Western focus)",
        "Regional Extensions": "British Columbia; Alberta; Saskatchewan; Manitoba stores",
        "Warehouse / Logistics": "Yes — retail pharmacy/CE logistics",
        "Delivery / Storage": "Store + delivery",
    },
    "staples canada": {
        "Retail Chains": "Yes",
        "Distribution Type": "Omnichannel",
        "Regions Served": "Canada",
        "Regional Extensions": "Canada Staples retail / B2B",
        "Warehouse / Logistics": "Yes — retail & B2B logistics",
        "Delivery / Storage": "Store + B2B delivery",
    },
    "visions electronics": {
        "Retail Chains": "Yes",
        "Distribution Type": "Omnichannel",
        "Regions Served": "Canada",
        "Regional Extensions": "Western Canada CE specialty retail",
        "Warehouse / Logistics": "Yes — retail logistics",
        "Delivery / Storage": "Store + delivery",
    },
    "mobile klinik": {
        "Retail Chains": "Yes",
        "Distribution Type": "Specialty retail",
        "Regions Served": "Canada",
        "Regional Extensions": "Canada repair/retail locations",
        "Warehouse / Logistics": "Partial — store-level inventory",
        "Delivery / Storage": "In-store repair / device sales",
    },
    "tbooth wireless": {
        "Retail Chains": "Yes",
        "Distribution Type": "Carrier dealer retail",
        "Regions Served": "Canada",
        "Regional Extensions": "Canada multi-carrier dealer stores",
        "Warehouse / Logistics": "Partial — dealer inventory",
        "Delivery / Storage": "In-store activation",
    },
    "powerhouse alliance": {
        "Retail Chains": "Partner network",
        "Distribution Type": "Dealer buying group",
        "Regions Served": "United States",
        "Regional Extensions": "US independent wireless dealer network",
        "Warehouse / Logistics": "Partial — member dealer logistics",
        "Delivery / Storage": "Dealer store fulfillment",
    },
    "axiom telecom": {
        "Retail Chains": "Yes",
        "Distribution Type": "Retail / Wholesale",
        "Regions Served": "United Arab Emirates; Saudi Arabia; Bahrain (selected)",
        "Regional Extensions": "UAE/KSA retail + distribution ops",
        "Warehouse / Logistics": "Yes — GCC distribution & retail logistics",
        "Delivery / Storage": "Store + wholesale fulfillment",
    },
    "emitac": {
        "Retail Chains": "No",
        "Distribution Type": "Value-Added Distributor",
        "Regions Served": "United Arab Emirates; Middle East",
        "Regional Extensions": "UAE technology distribution",
        "Warehouse / Logistics": "Yes — UAE logistics",
        "Delivery / Storage": "B2B fulfillment",
    },
    "celltel middle east l.l.c.": {
        "Retail Chains": "No",
        "Distribution Type": "Wholesale / Export",
        "Regions Served": "United Arab Emirates; Middle East; export markets",
        "Regional Extensions": "Dubai Deira + DAFZ freezone warehouses",
        "Warehouse / Logistics": "Yes — Dubai bonded / freezone storage",
        "Delivery / Storage": "Wholesale & export packing / logistics",
    },
    "jumia": {
        "Retail Chains": "No",
        "Distribution Type": "Online retail / marketplace",
        "Regions Served": "Africa (selected markets)",
        "Regional Extensions": "Nigeria; Egypt; Kenya; other Jumia markets",
        "Warehouse / Logistics": "Yes — marketplace fulfillment",
        "Delivery / Storage": "E-commerce last-mile",
    },
    "lazada": {
        "Retail Chains": "No",
        "Distribution Type": "Online retail / marketplace",
        "Regions Served": "Southeast Asia",
        "Regional Extensions": "Singapore; Indonesia; Malaysia; Thailand; Vietnam; Philippines",
        "Warehouse / Logistics": "Yes — SEA fulfillment network",
        "Delivery / Storage": "E-commerce fulfillment",
    },
    "flipkart": {
        "Retail Chains": "No",
        "Distribution Type": "Online retail / marketplace",
        "Regions Served": "India",
        "Regional Extensions": "India nationwide e-commerce",
        "Warehouse / Logistics": "Yes — India fulfillment centers",
        "Delivery / Storage": "E-commerce last-mile",
    },
    "jd.com": {
        "Retail Chains": "No",
        "Distribution Type": "Online retail / marketplace",
        "Regions Served": "China; selective cross-border",
        "Regional Extensions": "China nationwide logistics",
        "Warehouse / Logistics": "Yes — JD Logistics network",
        "Delivery / Storage": "E-commerce fulfillment",
    },
    "newegg": {
        "Retail Chains": "No",
        "Distribution Type": "Online retail / marketplace",
        "Regions Served": "United States; Canada; selective international",
        "Regional Extensions": "US / Canada online CE",
        "Warehouse / Logistics": "Yes — online fulfillment",
        "Delivery / Storage": "E-commerce shipping",
    },
    "kogan": {
        "Retail Chains": "No",
        "Distribution Type": "Online retail / marketplace",
        "Regions Served": "Australia; New Zealand",
        "Regional Extensions": "ANZ online CE",
        "Warehouse / Logistics": "Yes — ANZ fulfillment",
        "Delivery / Storage": "E-commerce shipping",
    },
    "handtec": {
        "Retail Chains": "No",
        "Distribution Type": "Online retail / refurbished",
        "Regions Served": "United Kingdom",
        "Regional Extensions": "UK online refurbished / used phones",
        "Warehouse / Logistics": "Partial — UK grading & shipping ops",
        "Delivery / Storage": "E-commerce shipping with warranty returns",
    },
    "3g mobile": {
        "Retail Chains": "No",
        "Distribution Type": "Wholesale",
        "Regions Served": "United States; export markets",
        "Regional Extensions": "US B2B wholesale / export",
        "Warehouse / Logistics": "Yes — wholesale inventory",
        "Delivery / Storage": "B2B bulk shipping",
    },
    "t-gaia": {
        "Retail Chains": "Partner network",
        "Distribution Type": "Carrier distribution",
        "Regions Served": "Japan",
        "Regional Extensions": "Japan carrier / retailer handset channel",
        "Warehouse / Logistics": "Yes — Japan device distribution logistics",
        "Delivery / Storage": "Carrier/retail device fulfillment",
    },
    "vision distribution pvt. ltd.": {
        "Retail Chains": "Partner network",
        "Distribution Type": "Specialty Distributor",
        "Regions Served": "India",
        "Regional Extensions": "India Vivo/iQOO channel",
        "Warehouse / Logistics": "Yes — India brand distribution logistics",
        "Delivery / Storage": "Channel wholesale fulfillment",
    },
    "abc mobile ltd": {
        "Retail Chains": "No",
        "Distribution Type": "Wholesale",
        "Regions Served": "United Kingdom; Europe",
        "Regional Extensions": "UK wholesale to EU retailers",
        "Warehouse / Logistics": "Partial — wholesale stock",
        "Delivery / Storage": "B2B shipping",
    },
    "acquvia global": {
        "Retail Chains": "No",
        "Distribution Type": "Wholesale",
        "Regions Served": "Global (premium CE wholesale)",
        "Regional Extensions": "Export-oriented Apple-focused wholesale",
        "Warehouse / Logistics": "Partial — wholesale inventory",
        "Delivery / Storage": "B2B shipping",
    },
    "fry's electronics": {
        "Retail Chains": "Yes",
        "Distribution Type": "Retail (legacy)",
        "Regions Served": "United States (legacy)",
        "Regional Extensions": "Former US CE big-box footprint",
        "Warehouse / Logistics": "Not publicly disclosed (stores closed)",
        "Delivery / Storage": "Not publicly disclosed (legacy)",
    },
    "mr. checkout": {
        "Retail Chains": "Partner network",
        "Distribution Type": "Convenience wholesale",
        "Regions Served": "United States",
        "Regional Extensions": "US convenience / independent retail channel",
        "Warehouse / Logistics": "Yes — convenience wholesale logistics",
        "Delivery / Storage": "Route / wholesale delivery",
    },
}

COUNTRY_MAP = {
    "usa": "United States",
    "us": "United States",
    "u.s.": "United States",
    "u.s.a.": "United States",
    "united states of america": "United States",
    "uk": "United Kingdom",
    "u.k.": "United Kingdom",
    "great britain": "United Kingdom",
    "england": "United Kingdom",
    "uae": "United Arab Emirates",
    "u.a.e.": "United Arab Emirates",
    "dubai": "United Arab Emirates",
    "brasil": "Brazil",
    "brazil": "Brazil",
    "hong kong sar": "Hong Kong",
    "hk": "Hong Kong",
    "korea": "South Korea",
    "republic of korea": "South Korea",
    "russian federation": "Russia",
    "czechia": "Czech Republic",
    "holland": "Netherlands",
    "ksa": "Saudi Arabia",
    "saudi": "Saudi Arabia",
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


def _country_from_text(text: str) -> str:
    t = _norm(text)
    if not t or t in ("seed_file", "n/a", "na", "unknown", "global", "worldwide"):
        return ""
    # direct map
    if t in COUNTRY_MAP:
        return COUNTRY_MAP[t]
    # known full names
    known = [
        "United States",
        "United Kingdom",
        "United Arab Emirates",
        "Canada",
        "Germany",
        "France",
        "India",
        "China",
        "Japan",
        "Australia",
        "Brazil",
        "Spain",
        "Italy",
        "Poland",
        "Netherlands",
        "Belgium",
        "Switzerland",
        "Austria",
        "Sweden",
        "Norway",
        "Denmark",
        "Finland",
        "Ireland",
        "Portugal",
        "Turkey",
        "Saudi Arabia",
        "South Africa",
        "Nigeria",
        "Kenya",
        "Egypt",
        "Ghana",
        "Tanzania",
        "Singapore",
        "Malaysia",
        "Indonesia",
        "Thailand",
        "Vietnam",
        "Philippines",
        "Hong Kong",
        "Taiwan",
        "South Korea",
        "Mexico",
        "Argentina",
        "Chile",
        "Colombia",
        "Peru",
        "Uruguay",
        "Russia",
        "Ukraine",
        "Czech Republic",
        "Hungary",
        "Romania",
        "Greece",
        "Israel",
        "New Zealand",
        "Pakistan",
        "Bangladesh",
        "Sri Lanka",
        "Cyprus",
        "Latvia",
        "Lithuania",
        "Estonia",
        "Bahrain",
        "Qatar",
        "Oman",
        "Kuwait",
        "Jordan",
        "Lebanon",
        "Morocco",
        "Tunisia",
    ]
    for c in known:
        if _norm(c) == t or _norm(c) in t:
            return c
    # last comma token often country
    if "," in text:
        tail = text.split(",")[-1].strip()
        if tail and _norm(tail) not in ("seed_file",):
            mapped = _country_from_text(tail)
            if mapped:
                return mapped
            if len(tail) > 2 and not re.search(r"\d", tail):
                return COUNTRY_MAP.get(_norm(tail), tail.title() if len(tail) < 40 else "")
    return ""


def _infer_country(row: dict[str, str]) -> str:
    for field in ("Regions Served", "Headquarters", "Cities / Regions", "Regional Extensions"):
        v = row.get(field) or ""
        if _norm(v) in ("", "seed_file", "yes", "no"):
            continue
        c = _country_from_text(v)
        if c:
            return c
    # website TLD hints
    web = (row.get("Website") or "").lower()
    tld_map = {
        ".uk": "United Kingdom",
        ".de": "Germany",
        ".fr": "France",
        ".pl": "Poland",
        ".in": "India",
        ".br": "Brazil",
        ".au": "Australia",
        ".ca": "Canada",
        ".ae": "United Arab Emirates",
        ".jp": "Japan",
        ".nl": "Netherlands",
        ".es": "Spain",
        ".it": "Italy",
        ".sg": "Singapore",
        ".hk": "Hong Kong",
        ".za": "South Africa",
        ".id": "Indonesia",
        ".mx": "Mexico",
        ".se": "Sweden",
        ".no": "Norway",
        ".ie": "Ireland",
        ".cz": "Czech Republic",
        ".hu": "Hungary",
    }
    for tld, country in tld_map.items():
        if tld in web:
            return country
    return ""


def _extensions_from_row(row: dict[str, str], country: str) -> str:
    cities = (row.get("Cities / Regions") or "").strip()
    hq = (row.get("Headquarters") or "").strip()
    if cities and _norm(cities) not in ("seed_file", "n/a", country.lower()):
        # keep concise
        c = re.sub(r"\s+", " ", cities)
        if len(c) > 120:
            c = c[:117].rsplit(" ", 1)[0] + "…"
        return c
    if hq and _norm(hq) not in ("seed_file", "n/a"):
        return hq
    return country or "Not publicly disclosed"


def _dist_from_bt(row: dict[str, str]) -> str:
    bt = _norm(row.get("Business Type") or "")
    ecom = _norm(row.get("E-commerce") or "")
    name = _norm(row.get("Company") or "")
    spec = _norm(row.get("Specialty Focus") or "")
    if "carrier" in bt:
        return "Carrier retail"
    if "e-commerce" in bt or (ecom == "yes" and "retail" in bt and "distributor" not in bt):
        if "refurb" in spec:
            return "Online retail / refurbished"
        return "Online retail / marketplace" if "marketplace" in spec or name in ("amazon", "flipkart", "jd.com", "lazada", "newegg", "kogan") else "Online retail"
    if "systems integrator" in bt:
        return "Reseller / Systems Integrator"
    if "reseller" in bt:
        return "Reseller / Distributor"
    if "retailer" in bt and "distributor" in bt:
        return "Retail / Wholesale"
    if "retail" in bt:
        return "Omnichannel" if ecom == "yes" else "Retail"
    if "wholesaler" in bt or "wholesale" in bt:
        return "Wholesale"
    if "importer" in bt:
        return "Importer / Wholesale"
    if "vad" in bt or "value-added" in spec:
        return "Value-Added Distributor"
    if "distributor" in bt:
        # size heuristic via turnover/employees optional — default specialty for small names
        return "Broadline Distributor" if any(
            x in name for x in ("ingram", "synnex", "also", "esprinet", "ab sa", "elko", "d&h", "arrow", "redington", "westcoast", "asbis", "intcomex", "vst")
        ) else "Specialty Distributor"
    return "Channel Partner"


def _retail_from_bt(row: dict[str, str], dist: str) -> str:
    bt = _norm(row.get("Business Type") or "")
    name = _norm(row.get("Company") or "")
    if any(x in name for x in ("alliance", "buying group", "dealer network")):
        return "Partner network"
    if "carrier" in bt or "retail" in bt:
        return "Yes"
    if dist in ("Omnichannel", "Retail", "Retail / Wholesale", "Carrier retail", "Carrier dealer retail", "Specialty retail"):
        return "Yes"
    if dist in ("Online retail / marketplace", "Online retail", "Online retail / refurbished"):
        return "No"
    if "partner network" in dist.lower() or "dealer" in dist.lower():
        return "Partner network"
    return "No"


def _warehouse_delivery(row: dict[str, str], dist: str) -> tuple[str, str]:
    bt = _norm(row.get("Business Type") or "")
    name = _norm(row.get("Company") or "")
    # majors already in KNOWN
    if dist.startswith("Broadline") or dist.startswith("Specialty") or dist.startswith("Value-Added"):
        return ("Yes — regional distribution logistics", "B2B wholesale fulfillment")
    if "Carrier" in dist:
        return ("Yes — carrier device logistics", "Store activation + shipping")
    if "Omnichannel" in dist or dist == "Retail" or dist == "Retail / Wholesale":
        return ("Yes — retail / DC logistics", "Store + online fulfillment")
    if "Online" in dist or "marketplace" in dist.lower():
        return ("Yes — e-commerce fulfillment", "E-commerce shipping")
    if "Reseller" in dist or "Systems Integrator" in dist:
        return ("Partial — partner / config logistics", "B2B project fulfillment")
    if dist == "Wholesale" or "Importer" in dist:
        return ("Partial — wholesale inventory", "B2B bulk shipping")
    if "refurb" in _norm(row.get("Specialty Focus") or ""):
        return ("Partial — grading & shipping ops", "E-commerce / wholesale shipping")
    if "distributor" in bt:
        return ("Partial — regional stock", "B2B shipping")
    return ("Partial", "Standard channel fulfillment")


def _heuristic(row: dict[str, str]) -> dict[str, str]:
    name = _norm(row.get("Company") or "")
    if name in KNOWN:
        return dict(KNOWN[name])

    country = _infer_country(row)
    if not country:
        country = "Not publicly disclosed"

    # expand regions from existing non-seed value if multi
    rs_raw = row.get("Regions Served") or ""
    if _norm(rs_raw) not in ("", "seed_file", "yes", "no") and "seed_file" not in _norm(rs_raw):
        # normalize USA/Brasil etc.
        parts = [p.strip() for p in re.split(r"[;,/|]+", rs_raw) if p.strip()]
        normed = []
        for p in parts:
            c = _country_from_text(p) or COUNTRY_MAP.get(_norm(p), p)
            if c and c not in normed:
                normed.append(c)
        regions = "; ".join(normed) if normed else country
    else:
        regions = country

    # Global marker
    blob = f"{row.get('Specialty Focus') or ''} {row.get('Summary') or ''} {row.get('Regions Served') or ''}".lower()
    if any(x in blob for x in ("worldwide", "global distributor", "across the globe", "international markets")) and country not in (
        "Not publicly disclosed",
    ):
        if regions == country and country in ("United States", "United Arab Emirates", "Hong Kong", "Singapore"):
            regions = f"{country}; export / international markets"

    extensions = _extensions_from_row(row, country if country != "Not publicly disclosed" else regions)
    # ensure RS != RE when RE is just same single country and cities exist
    if _norm(extensions) == _norm(regions):
        hq = (row.get("Headquarters") or "").strip()
        if hq and _norm(hq) not in ("seed_file", _norm(regions)):
            extensions = hq
        elif country and country != "Not publicly disclosed":
            extensions = f"{country} primary market footprint"

    dist = _dist_from_bt(row)
    retail = _retail_from_bt(row, dist)
    wh, deliv = _warehouse_delivery(row, dist)

    return {
        "Retail Chains": retail,
        "Distribution Type": dist,
        "Regions Served": regions,
        "Regional Extensions": extensions,
        "Warehouse / Logistics": wh,
        "Delivery / Storage": deliv,
    }


async def _web_evidence(name: str, website: str) -> str:
    host = ""
    if website:
        host = re.sub(r"^https?://", "", website.strip(), flags=re.I).split("/")[0].replace("www.", "")
    queries = [
        f'"{name}" distributor OR retailer OR wholesale regions OR countries OR stores OR warehouse',
        f'"{name}" "headquarters" OR "distribution center" OR "serves" OR logistics',
    ]
    if host:
        queries.insert(0, f'site:{host} about OR locations OR warehouses OR stores OR partners')
    parts: list[str] = []
    try:
        from ddgs import DDGS

        def _one(q: str) -> list[str]:
            lines = []
            try:
                hits = list(DDGS().text(q, max_results=4)) or []
            except Exception:
                hits = []
            for h in hits:
                lines.append(f"- {h.get('title','')}: {str(h.get('body',''))[:220]}")
            return lines

        for q in queries[:2]:
            try:
                lines = await asyncio.wait_for(asyncio.to_thread(_one, q), timeout=18)
            except Exception:
                lines = []
            if lines:
                parts.append(f"Query: {q}\n" + "\n".join(lines))
    except Exception as err:
        parts.append(f"(ddgs: {err})")
    return "\n\n".join(parts)[:5500] or "(no web hits)"


def _client():
    from openai import OpenAI

    key = (os.getenv("OPENAI_API_KEY") or "").strip()
    if not key:
        raise RuntimeError("OPENAI_API_KEY missing")
    return OpenAI(api_key=key, base_url=BASE)


def _llm(client: Any, row: dict[str, str], evidence: str, draft: dict[str, str]) -> dict[str, Any]:
    system = (
        "Classify smartphone CHANNEL logistics/geo fields. JSON only. Prefer WEB EVIDENCE.\n"
        "Retail Chains: Yes | No | Partner network\n"
        "Distribution Type: Broadline Distributor | Specialty Distributor | Value-Added Distributor | "
        "Wholesale | Retail | Omnichannel | Retail / Wholesale | Online retail / marketplace | "
        "Online retail | Online retail / refurbished | Carrier retail | Carrier dealer retail | "
        "Reseller / Distributor | Reseller / Systems Integrator | Channel Partner | Importer / Wholesale | "
        "Dealer buying group | Convenience wholesale | Specialty retail\n"
        "Regions Served: countries/regions served (normalize USA->United States, UAE->United Arab Emirates, Brasil->Brazil).\n"
        "Regional Extensions: more specific footprint (cities, store network, hubs) — not identical copy of Regions Served when cities known.\n"
        "Warehouse / Logistics: Yes — … | Partial — … | Not publicly disclosed\n"
        "Delivery / Storage: short fulfillment description\n"
        "Do not invent warehouse addresses. If unknown use Partial / Not publicly disclosed."
    )
    user = f"""Company: {row.get('Company')}
Website: {row.get('Website')}
HQ: {row.get('Headquarters')}
Cities: {row.get('Cities / Regions')}
Business Type: {row.get('Business Type')}
E-commerce: {row.get('E-commerce')}
Specialty: {(row.get('Specialty Focus') or '')[:160]}
Draft: {json.dumps(draft, ensure_ascii=False)}

WEB EVIDENCE:
{evidence}

Return JSON with keys: retail_chains, distribution_type, regions_served, regional_extensions,
warehouse_logistics, delivery_storage, confidence, notes
"""
    resp = client.chat.completions.create(
        model=MODEL,
        temperature=0.1,
        max_tokens=500,
        response_format={"type": "json_object"},
        messages=[{"role": "system", "content": system}, {"role": "user", "content": user}],
    )
    raw = resp.choices[0].message.content or "{}"
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        m = re.search(r"\{[\s\S]*\}", raw)
        data = json.loads(m.group(0)) if m else {}
    return data if isinstance(data, dict) else {}


def _merge_llm(draft: dict[str, str], data: dict[str, Any]) -> dict[str, str]:
    out = dict(draft)
    mapping = {
        "retail_chains": "Retail Chains",
        "distribution_type": "Distribution Type",
        "regions_served": "Regions Served",
        "regional_extensions": "Regional Extensions",
        "warehouse_logistics": "Warehouse / Logistics",
        "delivery_storage": "Delivery / Storage",
    }
    for k, col in mapping.items():
        v = str(data.get(k) or "").strip()
        if not v:
            continue
        if col == "Retail Chains":
            vl = v.lower()
            if "partner" in vl:
                out[col] = "Partner network"
            elif vl.startswith("y"):
                out[col] = "Yes"
            elif vl.startswith("n"):
                out[col] = "No"
            else:
                out[col] = v[:40]
        elif col == "Regions Served":
            parts = [p.strip() for p in re.split(r"[;,]+", v) if p.strip()]
            normed = []
            for p in parts:
                c = _country_from_text(p) or COUNTRY_MAP.get(_norm(p), p)
                if c and c not in normed:
                    normed.append(c)
            if normed:
                out[col] = "; ".join(normed[:8])
        else:
            out[col] = v[:160]
    return out


def _needs_llm(row: dict[str, str]) -> bool:
    rs = _norm(row.get("Regions Served") or "")
    rext = _norm(row.get("Regional Extensions") or "")
    if rs in ("seed_file", "", "unknown", "n/a") or rext in ("seed_file",):
        return True
    if not _infer_country(row):
        return True
    return False


async def _process_one(sem: asyncio.Semaphore, client: Any | None, row: dict[str, str]) -> dict[str, str]:
    draft = _heuristic(row)
    name = _norm(row.get("Company") or "")
    if name in KNOWN:
        out = dict(row)
        out.update(draft)
        src = out.get("Data Sources") or ""
        if "web_channel_geo_known" not in src:
            out["Data Sources"] = (src + "; web_channel_geo_known").strip("; ")
        return out

    if SKIP_LLM or client is None or not _needs_llm(row):
        out = dict(row)
        out.update(draft)
        src = out.get("Data Sources") or ""
        if "web_channel_geo_heur" not in src:
            out["Data Sources"] = (src + "; web_channel_geo_heur").strip("; ")
        return out

    async with sem:
        evidence = await _web_evidence(row.get("Company") or "", row.get("Website") or "")
        try:
            data = await asyncio.to_thread(_llm, client, row, evidence, draft)
        except Exception as err:
            data = {"notes": f"llm_err:{err}"}
        merged = _merge_llm(draft, data)
        out = dict(row)
        out.update(merged)
        src = out.get("Data Sources") or ""
        if "web_channel_geo_llm" not in src:
            out["Data Sources"] = (src + "; web_channel_geo_llm").strip("; ")
        return out


def _write_excel(data: list[dict[str, str]], header_idx: int) -> Path:
    path = FINAL
    ts = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    bak = path.with_suffix(path.suffix + f".bak_channel_geo_{ts}")
    shutil.copy2(path, bak)
    wb = load_workbook(path)
    ws = wb[SHEET]
    hr = header_idx + 1
    col = {ws.cell(hr, c).value: c for c in range(1, ws.max_column + 1) if ws.cell(hr, c).value}
    wrap = Alignment(wrap_text=True, vertical="top")
    thin = Border(
        left=Side(style="thin", color="DDDDDD"),
        right=Side(style="thin", color="DDDDDD"),
        top=Side(style="thin", color="DDDDDD"),
        bottom=Side(style="thin", color="DDDDDD"),
    )
    name_to_r = {}
    for r in range(hr + 1, ws.max_row + 1):
        n = str(ws.cell(r, col["Company"]).value or "").strip()
        if n:
            name_to_r[_norm(n)] = r
    updated = 0
    for d in data:
        r = name_to_r.get(_norm(d.get("Company") or ""))
        if not r:
            continue
        for f in FIELDS + ["Data Sources"]:
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
    seed_n = sum(1 for r in data if _norm(r.get("Regions Served") or "") == "seed_file")
    print(f"loaded={len(data)} seed_file_regions={seed_n} known={len(KNOWN)}")
    work = data[:LIMIT] if LIMIT > 0 else data
    if LIMIT:
        print(f"LIMIT={LIMIT}")

    client = None if SKIP_LLM else _client()
    sem = asyncio.Semaphore(CONCURRENT)
    results: list[dict[str, str]] = []
    tasks = [_process_one(sem, client, row) for row in work]
    done = 0
    for coro in asyncio.as_completed(tasks):
        results.append(await coro)
        done += 1
        if done % 40 == 0 or done == len(tasks):
            print(f"progress {done}/{len(tasks)}")

    by_name = {_norm(r["Company"]): r for r in results}
    merged = [by_name.get(_norm(r["Company"]), r) if _norm(r["Company"]) in by_name else r for r in data]

    # final sanitize
    for r in merged:
        if _norm(r.get("Regions Served") or "") == "seed_file":
            c = _infer_country(r) or "Not publicly disclosed"
            r["Regions Served"] = c
            if _norm(r.get("Regional Extensions") or "") == "seed_file":
                r["Regional Extensions"] = _extensions_from_row(r, c)
        # normalize country aliases in Regions Served
        parts = [p.strip() for p in re.split(r"[;,]+", r.get("Regions Served") or "") if p.strip()]
        normed = []
        for p in parts:
            c = _country_from_text(p) or COUNTRY_MAP.get(_norm(p), p)
            if c and c not in normed:
                normed.append(c)
        if normed:
            r["Regions Served"] = "; ".join(normed)
        if _norm(r.get("Retail Chains") or "") not in ("yes", "no", "partner network"):
            # leave descriptive Partner network / Yes / No only
            vl = _norm(r.get("Retail Chains") or "")
            if "partner" in vl:
                r["Retail Chains"] = "Partner network"
            elif vl.startswith("y"):
                r["Retail Chains"] = "Yes"
            else:
                r["Retail Chains"] = "No"

    _write_excel(merged, header_idx)

    retail_c = Counter(r.get("Retail Chains") for r in merged)
    dist_c = Counter(r.get("Distribution Type") for r in merged)
    wh_c = Counter((r.get("Warehouse / Logistics") or "")[:40] for r in merged)
    seed_left = sum(1 for r in merged if "seed_file" in _norm(r.get("Regions Served") or ""))
    same = sum(1 for r in merged if _norm(r.get("Regions Served")) == _norm(r.get("Regional Extensions")))
    audit = {
        "ts": datetime.now(timezone.utc).isoformat(),
        "total": len(merged),
        "seed_file_left": seed_left,
        "regions_eq_extensions": same,
        "retail_dist": retail_c.most_common(),
        "distribution_type_dist": dist_c.most_common(15),
        "warehouse_top": wh_c.most_common(10),
        "sample": [
            {k: r.get(k) for k in ["Company"] + FIELDS}
            for r in merged
            if r["Company"]
            in (
                "Best Buy",
                "Amazon",
                "Ingram Micro",
                "MediaMarktSaturn",
                "Currys",
                "Erajaya",
                "Handtec",
                "Celltel Middle East L.L.C.",
                "TD SYNNEX",
                "Reliance Jio",
            )
        ],
    }
    out = FINAL.parent / "channel_geo_audit.json"
    out.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps({k: audit[k] for k in audit if k != "sample"}, indent=2, ensure_ascii=False))
    print("audit", out)


if __name__ == "__main__":
    asyncio.run(main())
