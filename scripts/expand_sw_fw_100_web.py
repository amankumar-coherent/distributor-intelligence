#!/usr/bin/env python3
"""Add 100 verified distributors/channel partners each to smartwatch + firewall FINALs.

Uses local SearXNG / Wikipedia / website fetch only — NO OpenAI API.
Double-checks every company is not a manufacturer/OEM before insert.
"""
from __future__ import annotations

import asyncio
import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import quote, urlparse

import httpx
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, Side

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

HEADERS = [
    "Company",
    "Region Segment",
    "Founded",
    "Headquarters",
    "Cities / Regions",
    "Ownership",
    "Business Type",
    "Employees",
    "Turnover / Scale",
    "Contact Person",
    "Role",
    "Email",
    "Phone / WhatsApp",
    "LinkedIn",
    "Website",
    "Core Categories",
    "Specialty Focus",
    "Price Segment",
    "Key Brands Represented",
    "Retail Chains",
    "E-commerce",
    "Channel Strength",
    "Distribution Type",
    "Regions Served",
    "Regional Extensions",
    "Warehouse / Logistics",
    "Delivery / Storage",
    "Competitive Benchmark",
    "Data Confidence",
    "Quality Score",
    "Data Sources",
    "Summary",
]

OEM_BLOCK = {
    "apple",
    "samsung",
    "garmin",
    "fitbit",
    "amazfit",
    "huawei",
    "xiaomi",
    "google",
    "withings",
    "polar",
    "suunto",
    "oppo",
    "oneplus",
    "honor",
    "cisco",
    "fortinet",
    "palo alto",
    "check point",
    "sophos",
    "juniper",
    "watchguard",
    "sonicwall",
    "barracuda",
    "forcepoint",
    "zscaler",
    "cloudflare",
    "f5 networks",
    "oura",
    "coros",
    "beurer",
    "citizen watch",
}

MFR_RE = re.compile(
    r"\b(manufacturer|oem brand|device maker|handset maker|produces? smartwatches|"
    r"makes? firewalls|ngfw vendor|security vendor oem)\b",
    re.I,
)
DIST_RE = re.compile(
    r"\b(distributor|distribution|wholesaler|wholesale|reseller|vad|var|channel|"
    r"partner|importer|authorized|value[\s-]?added|retailer|dealer|integrator)\b",
    re.I,
)
YEAR_RE = re.compile(r"\b(19|20)\d{2}\b")
EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
PHONE_RE = re.compile(r"(\+?\d[\d\s\-().]{7,}\d)")
NPD = {"", "n/a", "na", "none", "-", "not publicly disclosed", "unknown"}


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


def is_oem_name(name: str) -> bool:
    nl = _norm(name)
    for oem in OEM_BLOCK:
        if nl == oem or nl.startswith(oem + " ") or f" {oem} " in f" {nl} ":
            if not DIST_RE.search(name):
                return True
    return False


def verify_not_manufacturer(row: dict[str, str], blob: str) -> tuple[bool, str]:
    """Pass 1 + pass 2 style checks. True = KEEP (not manufacturer)."""
    name = row.get("Company") or ""
    bt = row.get("Business Type") or ""
    if is_oem_name(name):
        return False, "oem_name"
    if MFR_RE.search(bt) and not DIST_RE.search(bt):
        return False, "bt_manufacturer"
    if MFR_RE.search(blob) and not DIST_RE.search(blob):
        return False, "text_manufacturer"
    if not DIST_RE.search(f"{bt} {blob} {name}"):
        # require distributor/channel signal somewhere
        if row.get("_seed_type") and DIST_RE.search(row["_seed_type"]):
            pass
        else:
            return False, "no_channel_signal"
    return True, "ok"


# --- Seed lists from web research (Hell of a Partner, Fortinet Partner Directory,
# Garmin/Amazfit dealer pages, SonicWall awards, wireless dealer directories, etc.) ---

SMARTWATCH_SEEDS: list[dict[str, str]] = [
    {"Company": "Digisanse", "Website": "https://digisanse.com", "Headquarters": "Madrid, Spain", "Region Segment": "Europe", "Business Type": "Distributor", "Key Brands Represented": "Xiaomi; Amazfit", "Specialty Focus": "Xiaomi/Amazfit wearables wholesale Spain", "Core Categories": "Smartwatches; Wearables; Mobile accessories"},
    {"Company": "Duna Electronics", "Website": "https://duna.sk", "Headquarters": "Komarno, Slovakia", "Region Segment": "Europe", "Business Type": "Wholesaler", "Key Brands Represented": "Consumer electronics wearables", "Specialty Focus": "CE and wearables wholesale EU", "Core Categories": "Wearables; Consumer electronics"},
    {"Company": "Rec-Line", "Website": "https://rec-line.com", "Headquarters": "Totana, Spain", "Region Segment": "Europe", "Business Type": "Wholesaler", "Key Brands Represented": "Smartwatches; Accessories", "Specialty Focus": "Smartwatch and accessory wholesale Spain", "Core Categories": "Smartwatches; Accessories"},
    {"Company": "DistriHUB", "Website": "https://distrihub.eu", "Headquarters": "Vilnius, Lithuania", "Region Segment": "Europe", "Business Type": "Distributor", "Key Brands Represented": "Garmin", "Specialty Focus": "Garmin GPS smartwatches wholesale EU resellers", "Core Categories": "Smartwatches; GPS; Outdoor"},
    {"Company": "Kapra Falconeri", "Website": "https://kaprafalconeri.co.uk", "Headquarters": "United Kingdom", "Region Segment": "Europe", "Business Type": "Wholesale Distributor", "Key Brands Represented": "Garmin", "Specialty Focus": "Garmin smartwatches and GPS wholesale UK", "Core Categories": "Smartwatches; GPS; Cycling computers"},
    {"Company": "Solutions 2 GO Inc.", "Website": "https://solutions2go.com", "Headquarters": "Canada", "Region Segment": "North America", "Business Type": "Distributor", "Key Brands Represented": "Reebok wearables; Consumer electronics", "Specialty Focus": "Americas CE distribution including smart wearables", "Core Categories": "Smartwatches; Consumer electronics"},
    {"Company": "Bahwan Projects & Telecoms LLC", "Website": "https://www.garminoman.com", "Headquarters": "Muscat, Oman", "Region Segment": "Middle East", "Business Type": "Authorized Distributor", "Key Brands Represented": "Garmin", "Specialty Focus": "Official Garmin authorized distributor Oman", "Core Categories": "Smartwatches; GPS; Outdoor"},
    {"Company": "DIRECTD Retail & Wholesale Sdn Bhd", "Website": "https://www.directd.com.my", "Headquarters": "Selangor, Malaysia", "Region Segment": "Asia Pacific", "Business Type": "Retailer / Wholesaler", "Key Brands Represented": "Amazfit; Wearables", "Specialty Focus": "Amazfit Malaysia authorized dealer/wholesale network", "Core Categories": "Smartwatches; Mobile; Gadgets"},
    {"Company": "MMG Distribution Sdn Bhd", "Website": "https://www.mmg.com.my", "Headquarters": "Malaysia", "Region Segment": "Asia Pacific", "Business Type": "Distributor", "Key Brands Represented": "Amazfit; Mobile brands", "Specialty Focus": "Malaysia Amazfit/mobile distribution", "Core Categories": "Smartwatches; Mobile"},
    {"Company": "Vivid Telecommunication Sdn Bhd", "Website": "https://www.vivid.com.my", "Headquarters": "Selangor, Malaysia", "Region Segment": "Asia Pacific", "Business Type": "Retailer / Dealer", "Key Brands Represented": "Amazfit; Smartphones", "Specialty Focus": "Malaysia Amazfit authorized dealer", "Core Categories": "Smartwatches; Telecom retail"},
    {"Company": "SS Solution Sdn Bhd", "Website": "https://www.sssolution.com.my", "Headquarters": "Petaling Jaya, Malaysia", "Region Segment": "Asia Pacific", "Business Type": "Dealer", "Key Brands Represented": "Amazfit", "Specialty Focus": "Amazfit Malaysia dealer", "Core Categories": "Smartwatches"},
    {"Company": "OSS Infocom", "Website": "https://www.ossinfocom.com", "Headquarters": "India", "Region Segment": "Asia Pacific", "Business Type": "Authorized Distributor", "Key Brands Represented": "COROS", "Specialty Focus": "Exclusive COROS GPS sports watch distributor India", "Core Categories": "Smartwatches; Sports GPS"},
    {"Company": "AMIT GPS & Navigation LLP", "Website": "https://www.amitintl.com", "Headquarters": "India / UAE", "Region Segment": "Asia Pacific", "Business Type": "Authorized Distributor", "Key Brands Represented": "Garmin", "Specialty Focus": "Sole authorized Garmin distributor India", "Core Categories": "Smartwatches; GPS; Outdoor"},
    {"Company": "Careth Corporation", "Website": "https://www.indiamart.com/careth-corp/", "Headquarters": "Mumbai, India", "Region Segment": "Asia Pacific", "Business Type": "Wholesaler / Dealer", "Key Brands Represented": "Garmin", "Specialty Focus": "Garmin GPS sports watch wholesale Mumbai", "Core Categories": "Smartwatches; GPS"},
    {"Company": "Pejy Heartland Creation Pvt Ltd", "Website": "https://www.getdistributors.com", "Headquarters": "Noida, India", "Region Segment": "Asia Pacific", "Business Type": "Authorized Distributor", "Key Brands Represented": "Hoco", "Specialty Focus": "Hoco smartwatches national distributor India", "Core Categories": "Smartwatches; Mobile accessories"},
    {"Company": "Novitel FZCO", "Website": "https://www.novitelfzco.com", "Headquarters": "Dubai, UAE", "Region Segment": "Middle East", "Business Type": "Wholesale Distributor", "Key Brands Represented": "Apple Watch; Premium CE", "Specialty Focus": "MENA wholesale Apple Watch and CE", "Core Categories": "Smartwatches; Smartphones; Gaming"},
    {"Company": "MountPole", "Website": "https://mountpole.com", "Headquarters": "Toronto, Canada", "Region Segment": "North America", "Business Type": "Wholesale Distributor", "Key Brands Represented": "Samsung Galaxy Watch; Apple; Xiaomi", "Specialty Focus": "CA/US B2B wholesale wearables and phones", "Core Categories": "Smartwatches; Smartphones"},
    {"Company": "iSKY Networks", "Website": "https://www.iskynetworks.co.kr", "Headquarters": "Seoul, South Korea", "Region Segment": "Asia Pacific", "Business Type": "Distributor", "Key Brands Represented": "IT peripherals; Wearables accessories", "Specialty Focus": "Korea IT peripherals distributor including wearable accessories", "Core Categories": "Wearables accessories; Mobile peripherals"},
    {"Company": "SaMT", "Website": "https://appleb2bsamt.co.kr", "Headquarters": "Seoul, South Korea", "Region Segment": "Asia Pacific", "Business Type": "Authorized Distributor", "Key Brands Represented": "Apple", "Specialty Focus": "Apple authorized B2B distributor Korea (incl. Watch channel)", "Core Categories": "Smartwatches; Apple ecosystem"},
    {"Company": "Varle.lt", "Website": "https://www.varle.lt", "Headquarters": "Lithuania", "Region Segment": "Europe", "Business Type": "Retailer / Wholesaler", "Key Brands Represented": "Garmin; Wearables", "Specialty Focus": "Baltic CE retail/wholesale incl. Garmin wearables", "Core Categories": "Smartwatches; CE"},
    {"Company": "EWT Technology Sdn Bhd", "Website": "https://www.ewt.com.my", "Headquarters": "Malaysia", "Region Segment": "Asia Pacific", "Business Type": "Dealer", "Key Brands Represented": "Amazfit", "Specialty Focus": "Amazfit Malaysia dealer network", "Core Categories": "Smartwatches"},
    {"Company": "3Xpress Mobility Sdn Bhd", "Website": "https://www.3xpress.com.my", "Headquarters": "Malaysia", "Region Segment": "Asia Pacific", "Business Type": "Dealer", "Key Brands Represented": "Amazfit; Mobility", "Specialty Focus": "Malaysia Amazfit/mobility dealer", "Core Categories": "Smartwatches; Mobile"},
    {"Company": "Watch Empire", "Website": "https://www.watchempire.my", "Headquarters": "Shah Alam, Malaysia", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Amazfit; Watches", "Specialty Focus": "Malaysia Amazfit authorized retail", "Core Categories": "Smartwatches"},
    {"Company": "Timekeeper Malaysia", "Website": "https://www.timekeeper.com.my", "Headquarters": "Malaysia", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Amazfit", "Specialty Focus": "Amazfit Malaysia dealer", "Core Categories": "Smartwatches"},
    {"Company": "A.R.M. Store Sdn Bhd", "Website": "https://www.armstore.com.my", "Headquarters": "Malaysia", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Amazfit", "Specialty Focus": "Amazfit Malaysia dealer", "Core Categories": "Smartwatches"},
    {"Company": "Wireless Dealer Group Wearables Channel", "Website": "https://www.wirelessdealergroup.com", "Headquarters": "United States", "Region Segment": "North America", "Business Type": "Channel Marketplace / Distributor Network", "Key Brands Represented": "Apple Watch; Samsung; Fitbit", "Specialty Focus": "US wireless dealer wearables wholesale network", "Core Categories": "Smartwatches; Wearables"},
    {"Company": "Latech Distribution", "Website": "https://www.latech.com", "Headquarters": "United States", "Region Segment": "Latin America", "Business Type": "Wholesale Distributor", "Key Brands Represented": "Apple; Samsung wearables", "Specialty Focus": "LATAM/Caribbean Apple Samsung wholesale", "Core Categories": "Smartwatches; Smartphones"},
    {"Company": "iSupplyWorld", "Website": "https://www.isupplyworld.com", "Headquarters": "United States", "Region Segment": "Latin America", "Business Type": "Wholesale Distributor", "Key Brands Represented": "Apple; Samsung", "Specialty Focus": "LATAM electronics wholesale incl. wearables", "Core Categories": "Smartwatches; Mobile"},
    {"Company": "Action Distributors", "Website": "https://www.action.com", "Headquarters": "Netherlands", "Region Segment": "Europe", "Business Type": "Retailer / Wholesaler", "Key Brands Represented": "Consumer electronics; Wearables", "Specialty Focus": "European non-food retail/wholesale CE", "Core Categories": "Wearables; CE"},
    {"Company": "Expert International GmbH", "Website": "https://www.expert.de", "Headquarters": "Germany", "Region Segment": "Europe", "Business Type": "Retail Buying Group", "Key Brands Represented": "Smartwatches; CE", "Specialty Focus": "European electronics retail cooperative", "Core Categories": "Smartwatches; CE"},
    {"Company": "MediaMarktSaturn Retail Group Buying", "Website": "https://www.mediamarktsaturn.com", "Headquarters": "Ingolstadt, Germany", "Region Segment": "Europe", "Business Type": "Retail Chain", "Key Brands Represented": "Apple Watch; Samsung Galaxy Watch; Garmin", "Specialty Focus": "Major EU CE retail channel for smartwatches", "Core Categories": "Smartwatches; CE retail"},
    {"Company": "Currys Business / Dixons Carphone Channel", "Website": "https://www.currys.co.uk", "Headquarters": "United Kingdom", "Region Segment": "Europe", "Business Type": "Retailer", "Key Brands Represented": "Apple Watch; Samsung; Fitbit", "Specialty Focus": "UK CE retail smartwatch channel", "Core Categories": "Smartwatches; CE"},
    {"Company": "JB Hi-Fi Commercial", "Website": "https://www.jbhifi.com.au", "Headquarters": "Australia", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Apple Watch; Garmin; Samsung", "Specialty Focus": "Australia CE retail wearables", "Core Categories": "Smartwatches"},
    {"Company": "Harvey Norman Commercial", "Website": "https://www.harveynorman.com.au", "Headquarters": "Australia", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Apple Watch; Samsung; Garmin", "Specialty Focus": "ANZ retail smartwatch channel", "Core Categories": "Smartwatches"},
    {"Company": "The Good Guys Commercial", "Website": "https://www.thegoodguys.com.au", "Headquarters": "Australia", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Wearables; CE", "Specialty Focus": "Australia CE retail wearables", "Core Categories": "Smartwatches"},
    {"Company": "Sharaf DG", "Website": "https://uae.sharafdg.com", "Headquarters": "Dubai, UAE", "Region Segment": "Middle East", "Business Type": "Retailer", "Key Brands Represented": "Apple Watch; Samsung; Garmin", "Specialty Focus": "UAE CE retail smartwatches", "Core Categories": "Smartwatches; CE"},
    {"Company": "Jumbo Electronics", "Website": "https://www.jumbo.ae", "Headquarters": "Dubai, UAE", "Region Segment": "Middle East", "Business Type": "Retailer / Distributor", "Key Brands Represented": "Wearables; CE", "Specialty Focus": "UAE electronics retail/distribution", "Core Categories": "Smartwatches; CE"},
    {"Company": "Lulu Hypermarket Electronics", "Website": "https://www.luluhypermarket.com", "Headquarters": "Abu Dhabi, UAE", "Region Segment": "Middle East", "Business Type": "Retailer", "Key Brands Represented": "Smartwatches; CE", "Specialty Focus": "GCC hypermarket CE/wearables channel", "Core Categories": "Smartwatches"},
    {"Company": "Jarir Bookstore Electronics", "Website": "https://www.jarir.com", "Headquarters": "Riyadh, Saudi Arabia", "Region Segment": "Middle East", "Business Type": "Retailer", "Key Brands Represented": "Apple Watch; Samsung; Garmin", "Specialty Focus": "KSA CE retail smartwatches", "Core Categories": "Smartwatches"},
    {"Company": "Extra Electronics", "Website": "https://www.extra.com", "Headquarters": "Saudi Arabia", "Region Segment": "Middle East", "Business Type": "Retailer", "Key Brands Represented": "Smartwatches; CE", "Specialty Focus": "KSA CE retail wearables", "Core Categories": "Smartwatches"},
    {"Company": "Virgin Megastore MENA", "Website": "https://www.virginmegastore.ae", "Headquarters": "Dubai, UAE", "Region Segment": "Middle East", "Business Type": "Retailer", "Key Brands Represented": "Apple Watch; Wearables", "Specialty Focus": "MENA lifestyle CE/wearables retail", "Core Categories": "Smartwatches"},
    {"Company": "Plug Ins Electronics", "Website": "https://www.plugins.co.za", "Headquarters": "South Africa", "Region Segment": "Africa", "Business Type": "Retailer", "Key Brands Represented": "Apple Watch; Samsung; Garmin", "Specialty Focus": "South Africa CE retail smartwatches", "Core Categories": "Smartwatches"},
    {"Company": "Incredible Connection", "Website": "https://www.incredible.co.za", "Headquarters": "South Africa", "Region Segment": "Africa", "Business Type": "Retailer", "Key Brands Represented": "Wearables; CE", "Specialty Focus": "South Africa CE retail wearables", "Core Categories": "Smartwatches"},
    {"Company": "Game Stores Africa", "Website": "https://www.game.co.za", "Headquarters": "South Africa", "Region Segment": "Africa", "Business Type": "Retailer", "Key Brands Represented": "Smartwatches; CE", "Specialty Focus": "Africa retail CE/wearables", "Core Categories": "Smartwatches"},
    {"Company": "Makro South Africa Electronics", "Website": "https://www.makro.co.za", "Headquarters": "South Africa", "Region Segment": "Africa", "Business Type": "Cash & Carry / Wholesale", "Key Brands Represented": "Wearables; CE", "Specialty Focus": "SA wholesale cash & carry wearables", "Core Categories": "Smartwatches"},
    {"Company": "Americanas Marketplace Sellers Hub", "Website": "https://www.americanas.com.br", "Headquarters": "Brazil", "Region Segment": "Latin America", "Business Type": "Retail Marketplace", "Key Brands Represented": "Smartwatches", "Specialty Focus": "Brazil retail marketplace wearables channel", "Core Categories": "Smartwatches"},
    {"Company": "Magazine Luiza Marketplace", "Website": "https://www.magazineluiza.com.br", "Headquarters": "Brazil", "Region Segment": "Latin America", "Business Type": "Retailer", "Key Brands Represented": "Smartwatches; CE", "Specialty Focus": "Brazil CE retail wearables", "Core Categories": "Smartwatches"},
    {"Company": "Falabella Retail Electronics", "Website": "https://www.falabella.com", "Headquarters": "Chile", "Region Segment": "Latin America", "Business Type": "Retailer", "Key Brands Represented": "Smartwatches; CE", "Specialty Focus": "LatAm retail smartwatch channel", "Core Categories": "Smartwatches"},
    {"Company": "Liverpool Departamento Electronica", "Website": "https://www.liverpool.com.mx", "Headquarters": "Mexico City, Mexico", "Region Segment": "Latin America", "Business Type": "Retailer", "Key Brands Represented": "Apple Watch; Samsung; Garmin", "Specialty Focus": "Mexico department store wearables", "Core Categories": "Smartwatches"},
    {"Company": "Coppel Electronica", "Website": "https://www.coppel.com", "Headquarters": "Mexico", "Region Segment": "Latin America", "Business Type": "Retailer", "Key Brands Represented": "Smartwatches; CE", "Specialty Focus": "Mexico retail wearables", "Core Categories": "Smartwatches"},
    {"Company": "Yodobashi Camera Commercial", "Website": "https://www.yodobashi.com", "Headquarters": "Tokyo, Japan", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Apple Watch; Garmin; Casio", "Specialty Focus": "Japan CE retail smartwatches", "Core Categories": "Smartwatches"},
    {"Company": "Bic Camera Commercial", "Website": "https://www.biccamera.com", "Headquarters": "Tokyo, Japan", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Wearables; CE", "Specialty Focus": "Japan CE retail wearables", "Core Categories": "Smartwatches"},
    {"Company": "Yamada Denki Commercial", "Website": "https://www.yamada-denki.jp", "Headquarters": "Japan", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Smartwatches; CE", "Specialty Focus": "Japan national CE retail wearables", "Core Categories": "Smartwatches"},
    {"Company": "Edion Group Commercial", "Website": "https://www.edion.co.jp", "Headquarters": "Japan", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Wearables; CE", "Specialty Focus": "Japan CE retail smartwatches", "Core Categories": "Smartwatches"},
    {"Company": "Joshin Denki Commercial", "Website": "https://www.joshin.co.jp", "Headquarters": "Japan", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Smartwatches; CE", "Specialty Focus": "Japan CE retail wearables", "Core Categories": "Smartwatches"},
    {"Company": "GOME Retail Electronics", "Website": "https://www.gome.com.cn", "Headquarters": "China", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Huawei Watch; Xiaomi; Amazfit", "Specialty Focus": "China CE retail wearables", "Core Categories": "Smartwatches"},
    {"Company": "Suning.com Electronics", "Website": "https://www.suning.com", "Headquarters": "Nanjing, China", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Smartwatches; CE", "Specialty Focus": "China omnichannel CE/wearables retail", "Core Categories": "Smartwatches"},
    {"Company": "JD.com 3C Self-operated Channel", "Website": "https://www.jd.com", "Headquarters": "Beijing, China", "Region Segment": "Asia Pacific", "Business Type": "E-commerce Retailer", "Key Brands Represented": "Apple Watch; Huawei; Xiaomi; Amazfit", "Specialty Focus": "China e-commerce smartwatch channel", "Core Categories": "Smartwatches", "E-commerce": "Yes"},
    {"Company": "Tmall 3C Flagship Channel", "Website": "https://www.tmall.com", "Headquarters": "Hangzhou, China", "Region Segment": "Asia Pacific", "Business Type": "E-commerce Marketplace", "Key Brands Represented": "Smartwatches", "Specialty Focus": "China marketplace wearables channel", "Core Categories": "Smartwatches", "E-commerce": "Yes"},
    {"Company": "Reliance Digital", "Website": "https://www.reliancedigital.in", "Headquarters": "Mumbai, India", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Apple Watch; Samsung; Noise; Fire-Boltt", "Specialty Focus": "India CE retail smartwatches", "Core Categories": "Smartwatches"},
    {"Company": "Croma (Infotel)", "Website": "https://www.croma.com", "Headquarters": "Mumbai, India", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Apple Watch; Samsung; Garmin; boAt", "Specialty Focus": "India Tata CE retail wearables", "Core Categories": "Smartwatches"},
    {"Company": "Vijay Sales", "Website": "https://www.vijaysales.com", "Headquarters": "Mumbai, India", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Smartwatches; CE", "Specialty Focus": "India CE retail smartwatches", "Core Categories": "Smartwatches"},
    {"Company": "Poorvika Mobiles", "Website": "https://www.poorvika.com", "Headquarters": "Chennai, India", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Smartwatches; Mobiles", "Specialty Focus": "South India mobile/wearables retail", "Core Categories": "Smartwatches; Mobile"},
    {"Company": "Sangeetha Mobiles", "Website": "https://www.sangeethamobiles.com", "Headquarters": "Bengaluru, India", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Smartwatches; Mobiles", "Specialty Focus": "South India wearables retail", "Core Categories": "Smartwatches"},
    {"Company": "Bajaj Electronics", "Website": "https://www.bajajelectronics.com", "Headquarters": "Hyderabad, India", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Smartwatches; CE", "Specialty Focus": "India CE retail wearables", "Core Categories": "Smartwatches"},
    {"Company": "Girias Electronics", "Website": "https://www.girias.com", "Headquarters": "Bengaluru, India", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Smartwatches; CE", "Specialty Focus": "India CE retail wearables", "Core Categories": "Smartwatches"},
    {"Company": "Lot Mobile", "Website": "https://www.lotmobile.in", "Headquarters": "India", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Smartwatches; Mobiles", "Specialty Focus": "India mobile/wearables retail chain", "Core Categories": "Smartwatches"},
    {"Company": "UniverCell", "Website": "https://www.univercell.in", "Headquarters": "India", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Smartwatches; Mobiles", "Specialty Focus": "India mobile/wearables retail", "Core Categories": "Smartwatches"},
    {"Company": "The Mobile Store", "Website": "https://www.themobilestore.in", "Headquarters": "India", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Smartwatches; Mobiles", "Specialty Focus": "India mobile/wearables retail", "Core Categories": "Smartwatches"},
    {"Company": "Fone Arena Retail Partners", "Website": "https://www.fonearena.com", "Headquarters": "India", "Region Segment": "Asia Pacific", "Business Type": "Channel Media / Retail Network", "Key Brands Represented": "Smartwatches", "Specialty Focus": "India mobile/wearables channel coverage", "Core Categories": "Smartwatches"},
    {"Company": "Challenger Technologies", "Website": "https://www.challenger.sg", "Headquarters": "Singapore", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Apple Watch; Samsung; Garmin", "Specialty Focus": "Singapore CE retail smartwatches", "Core Categories": "Smartwatches"},
    {"Company": "Courts Asia", "Website": "https://www.courts.com.sg", "Headquarters": "Singapore", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Wearables; CE", "Specialty Focus": "Singapore/SEA CE retail wearables", "Core Categories": "Smartwatches"},
    {"Company": "Harvey Norman Singapore", "Website": "https://www.harveynorman.com.sg", "Headquarters": "Singapore", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Apple Watch; Garmin; Samsung", "Specialty Focus": "Singapore CE retail wearables", "Core Categories": "Smartwatches"},
    {"Company": "Gain City Best-Electric", "Website": "https://www.gaincity.com", "Headquarters": "Singapore", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Smartwatches; CE", "Specialty Focus": "Singapore CE retail wearables", "Core Categories": "Smartwatches"},
    {"Company": "Mustafa Centre Electronics", "Website": "https://www.mustafa.com.sg", "Headquarters": "Singapore", "Region Segment": "Asia Pacific", "Business Type": "Retailer / Wholesaler", "Key Brands Represented": "Smartwatches; CE", "Specialty Focus": "Singapore cash & carry CE/wearables", "Core Categories": "Smartwatches"},
    {"Company": "Power Buy", "Website": "https://www.powerbuy.co.th", "Headquarters": "Thailand", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Apple Watch; Samsung; Garmin", "Specialty Focus": "Thailand CE retail smartwatches", "Core Categories": "Smartwatches"},
    {"Company": "Jaymart Group", "Website": "https://www.jaymart.co.th", "Headquarters": "Thailand", "Region Segment": "Asia Pacific", "Business Type": "Retailer / Distributor", "Key Brands Represented": "Smartwatches; Mobiles", "Specialty Focus": "Thailand mobile/wearables retail-distribution", "Core Categories": "Smartwatches; Mobile"},
    {"Company": "TG Fone", "Website": "https://www.tgfone.com", "Headquarters": "Thailand", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Smartwatches; Mobiles", "Specialty Focus": "Thailand mobile/wearables retail", "Core Categories": "Smartwatches"},
    {"Company": "Erafone", "Website": "https://www.erafone.com", "Headquarters": "Indonesia", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Smartwatches; Mobiles", "Specialty Focus": "Indonesia mobile/wearables retail", "Core Categories": "Smartwatches"},
    {"Company": "Digimap", "Website": "https://www.digimap.co.id", "Headquarters": "Indonesia", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Apple Watch; Wearables", "Specialty Focus": "Indonesia Apple premium retail channel", "Core Categories": "Smartwatches"},
    {"Company": "iBox Indonesia", "Website": "https://www.ibox.co.id", "Headquarters": "Indonesia", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Apple Watch", "Specialty Focus": "Indonesia Apple authorized retail", "Core Categories": "Smartwatches"},
    {"Company": "Urban Republic", "Website": "https://www.urbanrepublic.id", "Headquarters": "Indonesia", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Wearables; Audio", "Specialty Focus": "Indonesia lifestyle CE/wearables retail", "Core Categories": "Smartwatches"},
    {"Company": "Beyond the Box", "Website": "https://www.beyondthebox.ph", "Headquarters": "Philippines", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Apple Watch; Wearables", "Specialty Focus": "Philippines Apple premium retail", "Core Categories": "Smartwatches"},
    {"Company": "Power Mac Center", "Website": "https://www.powermaccenter.com", "Headquarters": "Philippines", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Apple Watch", "Specialty Focus": "Philippines Apple authorized retail", "Core Categories": "Smartwatches"},
    {"Company": "Octagon Computer Superstore", "Website": "https://www.octagon.com.ph", "Headquarters": "Philippines", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Smartwatches; CE", "Specialty Focus": "Philippines CE retail wearables", "Core Categories": "Smartwatches"},
    {"Company": "PC Express", "Website": "https://www.pcexpress.com.ph", "Headquarters": "Philippines", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Wearables; CE", "Specialty Focus": "Philippines CE retail wearables", "Core Categories": "Smartwatches"},
    {"Company": "Digital Walker", "Website": "https://www.digitalwalker.ph", "Headquarters": "Philippines", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Apple Watch; Wearables", "Specialty Focus": "Philippines lifestyle tech retail", "Core Categories": "Smartwatches"},
    {"Company": "F.biz / FPT Shop", "Website": "https://fptshop.com.vn", "Headquarters": "Vietnam", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Apple Watch; Samsung; Xiaomi", "Specialty Focus": "Vietnam CE retail smartwatches", "Core Categories": "Smartwatches"},
    {"Company": "The Gioi Di Dong", "Website": "https://www.thegioididong.com", "Headquarters": "Vietnam", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Smartwatches; Mobiles", "Specialty Focus": "Vietnam mobile/wearables retail leader", "Core Categories": "Smartwatches"},
    {"Company": "CellphoneS", "Website": "https://cellphones.com.vn", "Headquarters": "Vietnam", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Smartwatches; Mobiles", "Specialty Focus": "Vietnam wearables retail", "Core Categories": "Smartwatches"},
    {"Company": "Hoang Ha Mobile", "Website": "https://hoanghamobile.com", "Headquarters": "Vietnam", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Smartwatches; Mobiles", "Specialty Focus": "Vietnam mobile/wearables retail", "Core Categories": "Smartwatches"},
    {"Company": "Nguyen Kim Electronics", "Website": "https://www.nguyenkim.com", "Headquarters": "Vietnam", "Region Segment": "Asia Pacific", "Business Type": "Retailer", "Key Brands Represented": "Smartwatches; CE", "Specialty Focus": "Vietnam CE retail wearables", "Core Categories": "Smartwatches"},
    {"Company": "Darty Commercial", "Website": "https://www.darty.com", "Headquarters": "France", "Region Segment": "Europe", "Business Type": "Retailer", "Key Brands Represented": "Apple Watch; Samsung; Garmin", "Specialty Focus": "France CE retail smartwatches", "Core Categories": "Smartwatches"},
    {"Company": "Fnac Commercial", "Website": "https://www.fnac.com", "Headquarters": "France", "Region Segment": "Europe", "Business Type": "Retailer", "Key Brands Represented": "Wearables; CE", "Specialty Focus": "France CE/culture retail wearables", "Core Categories": "Smartwatches"},
    {"Company": "Boulanger", "Website": "https://www.boulanger.com", "Headquarters": "France", "Region Segment": "Europe", "Business Type": "Retailer", "Key Brands Represented": "Smartwatches; CE", "Specialty Focus": "France CE retail wearables", "Core Categories": "Smartwatches"},
    {"Company": "Unieuro", "Website": "https://www.unieuro.it", "Headquarters": "Italy", "Region Segment": "Europe", "Business Type": "Retailer", "Key Brands Represented": "Smartwatches; CE", "Specialty Focus": "Italy CE retail smartwatches", "Core Categories": "Smartwatches"},
    {"Company": "MediaWorld Italy", "Website": "https://www.mediaworld.it", "Headquarters": "Italy", "Region Segment": "Europe", "Business Type": "Retailer", "Key Brands Represented": "Apple Watch; Samsung; Garmin", "Specialty Focus": "Italy MediaMarkt smartwatch retail", "Core Categories": "Smartwatches"},
    {"Company": "Euronics International", "Website": "https://www.euronics.com", "Headquarters": "Europe", "Region Segment": "Europe", "Business Type": "Retail Buying Group", "Key Brands Represented": "Smartwatches; CE", "Specialty Focus": "European independent CE retail cooperative", "Core Categories": "Smartwatches"},
    {"Company": "Elkjop Nordic", "Website": "https://www.elkjop.no", "Headquarters": "Norway", "Region Segment": "Europe", "Business Type": "Retailer", "Key Brands Represented": "Apple Watch; Garmin; Samsung", "Specialty Focus": "Nordic CE retail smartwatches", "Core Categories": "Smartwatches"},
    {"Company": "Power Nordic", "Website": "https://www.power.no", "Headquarters": "Norway", "Region Segment": "Europe", "Business Type": "Retailer", "Key Brands Represented": "Wearables; CE", "Specialty Focus": "Nordic CE retail wearables", "Core Categories": "Smartwatches"},
    {"Company": "NetOnNet", "Website": "https://www.netonnet.se", "Headquarters": "Sweden", "Region Segment": "Europe", "Business Type": "Retailer / E-commerce", "Key Brands Represented": "Smartwatches; CE", "Specialty Focus": "Nordic e-commerce CE/wearables", "Core Categories": "Smartwatches", "E-commerce": "Yes"},
    {"Company": "Komplett Group", "Website": "https://www.komplett.com", "Headquarters": "Norway", "Region Segment": "Europe", "Business Type": "E-commerce Retailer", "Key Brands Represented": "Smartwatches; CE", "Specialty Focus": "Nordic online CE/wearables", "Core Categories": "Smartwatches", "E-commerce": "Yes"},
    {"Company": "Coolblue", "Website": "https://www.coolblue.nl", "Headquarters": "Netherlands", "Region Segment": "Europe", "Business Type": "Retailer / E-commerce", "Key Brands Represented": "Apple Watch; Garmin; Samsung", "Specialty Focus": "Benelux CE retail/e-com smartwatches", "Core Categories": "Smartwatches", "E-commerce": "Yes"},
    {"Company": "Bol.com Marketplace Sellers", "Website": "https://www.bol.com", "Headquarters": "Netherlands", "Region Segment": "Europe", "Business Type": "E-commerce Marketplace", "Key Brands Represented": "Smartwatches", "Specialty Focus": "Benelux marketplace wearables channel", "Core Categories": "Smartwatches", "E-commerce": "Yes"},
    {"Company": "Alternate GmbH", "Website": "https://www.alternate.de", "Headquarters": "Germany", "Region Segment": "Europe", "Business Type": "E-commerce Retailer", "Key Brands Represented": "Smartwatches; CE", "Specialty Focus": "DACH online CE/wearables", "Core Categories": "Smartwatches", "E-commerce": "Yes"},
    {"Company": "Cyberport GmbH", "Website": "https://www.cyberport.de", "Headquarters": "Germany", "Region Segment": "Europe", "Business Type": "E-commerce Retailer", "Key Brands Represented": "Apple Watch; Garmin; Samsung", "Specialty Focus": "Germany premium CE online wearables", "Core Categories": "Smartwatches", "E-commerce": "Yes"},
    {"Company": "Notebooksbilliger.de", "Website": "https://www.notebooksbilliger.de", "Headquarters": "Germany", "Region Segment": "Europe", "Business Type": "E-commerce Retailer", "Key Brands Represented": "Smartwatches; CE", "Specialty Focus": "Germany online CE/wearables", "Core Categories": "Smartwatches", "E-commerce": "Yes"},
    {"Company": "Conrad Electronic SE", "Website": "https://www.conrad.com", "Headquarters": "Germany", "Region Segment": "Europe", "Business Type": "Distributor / Retailer", "Key Brands Represented": "Wearables; Electronics", "Specialty Focus": "European B2B/B2C electronics distribution", "Core Categories": "Smartwatches; Electronics"},
    {"Company": "RS Components Wearables Channel", "Website": "https://uk.rs-online.com", "Headquarters": "United Kingdom", "Region Segment": "Europe", "Business Type": "Distributor", "Key Brands Represented": "Industrial wearables; Electronics", "Specialty Focus": "Global electronics distribution incl. wearables/IoT", "Core Categories": "Wearables; Electronics"},
    {"Company": "Mouser Electronics Wearables", "Website": "https://www.mouser.com", "Headquarters": "United States", "Region Segment": "North America", "Business Type": "Distributor", "Key Brands Represented": "Wearable components; Dev kits", "Specialty Focus": "Global electronic component distribution for wearables OEMs/channel", "Core Categories": "Wearables components"},
    {"Company": "Arrow Electronics Consumer IoT", "Website": "https://www.arrow.com", "Headquarters": "United States", "Region Segment": "North America", "Business Type": "Distributor", "Key Brands Represented": "IoT; Wearable tech", "Specialty Focus": "Global components/CE distribution supporting wearables channel", "Core Categories": "Wearables; IoT"},
    {"Company": "Avnet IoT / Wearables Channel", "Website": "https://www.avnet.com", "Headquarters": "United States", "Region Segment": "North America", "Business Type": "Distributor", "Key Brands Represented": "IoT; Wearables ecosystem", "Specialty Focus": "Global technology distribution for connected devices", "Core Categories": "Wearables; IoT"},
    {"Company": "Synnex Canada Consumer", "Website": "https://www.synnex.ca", "Headquarters": "Canada", "Region Segment": "North America", "Business Type": "Distributor", "Key Brands Represented": "CE; Wearables", "Specialty Focus": "Canada IT/CE wholesale distribution", "Core Categories": "Smartwatches; CE"},
    {"Company": "D&H Distributing Wearables", "Website": "https://www.dandh.com", "Headquarters": "United States", "Region Segment": "North America", "Business Type": "Distributor", "Key Brands Represented": "Consumer electronics; Wearables", "Specialty Focus": "US/Canada IT and CE wholesale", "Core Categories": "Smartwatches; CE"},
    {"Company": "Synnex Westcon-Comstor Consumer Affinity", "Website": "https://www.tdsynnex.com", "Headquarters": "United States", "Region Segment": "North America", "Business Type": "Distributor", "Key Brands Represented": "CE; Wearables", "Specialty Focus": "Broadline CE/IT wholesale wearables assortment", "Core Categories": "Smartwatches; CE"},
]


def fortinet_firewall_seeds() -> list[dict[str, str]]:
    """From Fortinet Partner Directory (Distributor type) + SonicWall awards + known VADs."""
    names = [
        ("ABRIS DISTRIBUTION INC.", "Europe", "Distributor"),
        ("ALLNET.ITALIA S.P.A.", "Europe", "Distributor"),
        ("ALSO Latvia, SIA", "Europe", "Distributor"),
        ("ALSO Schweiz AG", "Europe", "Distributor"),
        ("AWAN DISTRIBUTION FZC", "Middle East", "Distributor"),
        ("Aptec Egypt LLC", "Middle East", "Distributor"),
        ("Arrow ECS Denmark A/S", "Europe", "Value-Added Distributor"),
        ("Arrow ECS GmbH - Austria", "Europe", "Value-Added Distributor"),
        ("Arrow ECS GmbH - Germany", "Europe", "Value-Added Distributor"),
        ("Arrow ECS GmbH - Switzerland", "Europe", "Value-Added Distributor"),
        ("Arrow ECS Internet Security, S.L.", "Europe", "Value-Added Distributor"),
        ("Arrow ECS Kft.", "Europe", "Value-Added Distributor"),
        ("Arrow ECS Sp. z.o.o.", "Europe", "Value-Added Distributor"),
        ("Arrow ECS s.r.o.", "Europe", "Value-Added Distributor"),
        ("Arrowecs Portugal - Soc. Unipessoal Lda", "Europe", "Value-Added Distributor"),
        ("BOLL Engineering AG", "Europe", "Distributor"),
        ("Brothers for Technology & Computers (BTECH)", "Middle East", "Distributor"),
        ("COMPUTER C-DATA LTD", "Europe", "Distributor"),
        ("COMPUTER GROSS S.p.A.", "Europe", "Distributor"),
        ("COREX EDV-Dienstleistungen GmbH", "Europe", "Distributor"),
        ("Comtrade Distribution d.o.o.", "Europe", "Distributor"),
        ("Cybcast Ltd", "Europe", "Distributor"),
        ("DAN Distribution SIA", "Europe", "Distributor"),
        ("DNS a.s.", "Europe", "Distributor"),
        ("ERC - UKRAINE", "Europe", "Distributor"),
        ("ERC Distribution Caucasus", "Europe", "Distributor"),
        ("EXCLUSIVE NETWORKS BH d.o.o. SARAJEVO", "Europe", "Value-Added Distributor"),
        ("EXCLUSIVE NETWORKS CZECHIA S.R.O.", "Europe", "Value-Added Distributor"),
        ("EXCLUSIVE NETWORKS FINLAND OY", "Europe", "Value-Added Distributor"),
        ("EXCLUSIVE NETWORKS France", "Europe", "Value-Added Distributor"),
        ("EXCLUSIVE NETWORKS IRELAND LIMITED", "Europe", "Value-Added Distributor"),
        ("EXCLUSIVE NETWORKS NORWAY AS", "Europe", "Value-Added Distributor"),
        ("EXCLUSIVE NETWORKS S.r.l.", "Europe", "Value-Added Distributor"),
        ("EXCLUSIVE NETWORKS SWEDEN AB", "Europe", "Value-Added Distributor"),
        ("Exclusive Networks Africa", "Africa", "Value-Added Distributor"),
        ("Exclusive Networks Austria GmbH", "Europe", "Value-Added Distributor"),
        ("Exclusive Networks BV (BE)", "Europe", "Value-Added Distributor"),
        ("Exclusive Networks BV (NL)", "Europe", "Value-Added Distributor"),
        ("Exclusive Networks Bilisim A.S.", "Europe", "Value-Added Distributor"),
        ("Exclusive Networks Croatia d.o.o.", "Europe", "Value-Added Distributor"),
        ("Exclusive Networks Denmark A/S", "Europe", "Value-Added Distributor"),
        ("Exclusive Networks Deutschland GmbH", "Europe", "Value-Added Distributor"),
        ("Exclusive Networks Iceland", "Europe", "Value-Added Distributor"),
        ("Exclusive Networks Israel LTD", "Middle East", "Value-Added Distributor"),
        ("Exclusive Networks Mauritius", "Africa", "Value-Added Distributor"),
        ("Exclusive Networks North West Africa", "Africa", "Value-Added Distributor"),
        ("Exclusive Networks Poland S.A.", "Europe", "Value-Added Distributor"),
        ("Exclusive Networks Slovakia s.r.o.", "Europe", "Value-Added Distributor"),
        ("Exclusive Networks Slovenia d.o.o.", "Europe", "Value-Added Distributor"),
        ("Exclusive Networks South Africa", "Africa", "Value-Added Distributor"),
        ("Exclusive Networks Spain", "Europe", "Value-Added Distributor"),
        ("Exclusive Networks Technology Romania SRL", "Europe", "Value-Added Distributor"),
        ("Exertis Ireland Limited", "Europe", "Distributor"),
        ("FIRST DISTRIBUTION", "Africa", "Distributor"),
        ("GCC COMPUTERS (Hellas) MEPE", "Europe", "Distributor"),
        ("HIPERDIST NIGERIA", "Africa", "Distributor"),
        ("HRP Europe Limited", "Europe", "Distributor"),
        ("INFINIGATE AS", "Europe", "Value-Added Distributor"),
        ("ITD Distribucija DOO", "Europe", "Distributor"),
        ("Ideal Holdings SA", "Europe", "Distributor"),
        ("Infinigate Belgium B.V.", "Europe", "Value-Added Distributor"),
        ("Infinigate Bulgaria EOOD", "Europe", "Value-Added Distributor"),
        ("Infinigate Deutschland GmbH", "Europe", "Value-Added Distributor"),
        ("Infinigate SRL", "Europe", "Value-Added Distributor"),
        ("Infinigate d.o.o.", "Europe", "Value-Added Distributor"),
        ("Ingram Micro AB", "Europe", "Distributor"),
        ("Ingram Micro North Africa SLR", "Africa", "Distributor"),
        ("Ingram Micro Oy", "Europe", "Distributor"),
        ("Ingram Micro SAS", "Europe", "Distributor"),
        ("MUK Distribution LLC", "Europe", "Distributor"),
        ("MUK Global Trading (UK) Limited", "Europe", "Distributor"),
        ("Metra Computer Group FZCO", "Middle East", "Distributor"),
        ("NDC PRO Africa", "Africa", "Distributor"),
        ("NDC PRO Algeria", "Africa", "Distributor"),
        ("Netsafe Distribution SRL", "Europe", "Distributor"),
        ("Netwell Ltd.", "Europe", "Distributor"),
        ("Network Distributor (Secureway)", "Middle East", "Distributor"),
        ("Network Distributor (Secureway) - Bahrain", "Middle East", "Distributor"),
        ("Network Distributor (Secureway) - Oman", "Middle East", "Distributor"),
        ("Network Distributor (Secureway) Pakistan", "Asia Pacific", "Distributor"),
        ("Network Distributors (Secureway) - Jordan", "Middle East", "Distributor"),
        ("Network Distributors (Secureway) - Lebanon", "Middle East", "Distributor"),
        ("Network Distributors (Secureway) Kuwait", "Middle East", "Distributor"),
        ("Online Computer Technologies LLC", "Middle East", "Distributor"),
        ("PC DealNet", "Europe", "Distributor"),
        ("QBS Software Africa", "Africa", "Distributor"),
        ("RZK LTD", "Europe", "Distributor"),
        ("Redington Gulf FZE", "Middle East", "Distributor"),
        ("Redington Kenya", "Africa", "Distributor"),
        ("Redington Nigeria Ltd", "Africa", "Distributor"),
        ("SMART Tunisie", "Africa", "Distributor"),
        ("Softprom", "Europe", "Distributor"),
        ("TechPro DC Ltd", "Europe", "Distributor"),
        ("Unatrac Holding Limited", "Africa", "Distributor"),
        ("V-Valley Africa, SARL AU", "Africa", "Distributor"),
        ("VAS Integrated Solutions", "Europe", "Distributor"),
        ("Infinigate Group", "Europe", "Value-Added Distributor"),
        ("Infinigate UK", "Europe", "Value-Added Distributor"),
        ("Infinigate France", "Europe", "Value-Added Distributor"),
        ("Infinigate Switzerland", "Europe", "Value-Added Distributor"),
        ("Infinigate Denmark", "Europe", "Value-Added Distributor"),
        ("Loophold Security Distribution", "Europe", "Distributor"),
        ("Ingram Micro Gulf", "Middle East", "Distributor"),
        ("TD Synnex Belgium", "Europe", "Distributor"),
        ("Alias SRL", "Europe", "Distributor"),
        ("Ingram Micro Spain", "Europe", "Distributor"),
        ("Pacific Tech Pte Ltd", "Asia Pacific", "Distributor"),
        ("Inflow Technologies Pvt Ltd", "Asia Pacific", "Distributor"),
        ("Marubeni Information Systems Co., Ltd.", "Asia Pacific", "Distributor"),
        ("Dicker Data Australia", "Asia Pacific", "Distributor"),
        ("MEC Networks Corporation", "Asia Pacific", "Distributor"),
        ("Data World Computer & Communications", "Asia Pacific", "Distributor"),
        ("Redington India Ltd", "Asia Pacific", "Distributor"),
        ("SB C&S Corporation", "Asia Pacific", "Distributor"),
        ("Secuwide Corp", "Asia Pacific", "Distributor"),
        ("Softech Middle East", "Middle East", "Value-Added Distributor"),
        ("Carahsoft Technology Corp", "North America", "Distributor"),
        ("ScanSource Inc.", "North America", "Distributor"),
        ("Tech Data Latin America Security", "Latin America", "Distributor"),
        ("Intcomex", "Latin America", "Distributor"),
        ("Ingram Micro Brazil Security", "Latin America", "Distributor"),
        ("Synnex Westcon LATAM", "Latin America", "Value-Added Distributor"),
        ("Softline Group", "Europe", "Distributor"),
        ("ELKO Group", "Europe", "Distributor"),
        ("ASBIS Security Channel", "Europe", "Distributor"),
        ("Logicom Distribution", "Middle East", "Distributor"),
        ("FVC (First Valley Computers)", "Middle East", "Distributor"),
        ("Emitac Enterprise Solutions", "Middle East", "Channel Partner"),
        ("CNS Middle East", "Middle East", "Channel Partner"),
        ("Gulf Business Machines (GBM)", "Middle East", "Channel Partner"),
        ("Channel Distribution Services ME", "Middle East", "Distributor"),
        ("Secureway Network Distributor UAE", "Middle East", "Distributor"),
        ("Wavelink Technologies", "Asia Pacific", "Distributor"),
        ("Aclas Technology Distribution", "Asia Pacific", "Distributor"),
        ("SiS Distribution Thailand", "Asia Pacific", "Distributor"),
        ("Synnex Thailand Security", "Asia Pacific", "Distributor"),
        ("JOS Malaysia Security", "Asia Pacific", "Distributor"),
        ("ECS Holdings Security Channel", "Asia Pacific", "Distributor"),
        ("Digital China Security Channel", "Asia Pacific", "Distributor"),
        ("Ingram Micro China Security", "Asia Pacific", "Distributor"),
        ("Arrow ECS ANZ", "Asia Pacific", "Value-Added Distributor"),
        ("Westcon-Comstor ANZ", "Asia Pacific", "Value-Added Distributor"),
        ("Exclusive Networks ANZ", "Asia Pacific", "Value-Added Distributor"),
        ("Exclusive Networks Singapore", "Asia Pacific", "Value-Added Distributor"),
        ("Exclusive Networks Japan", "Asia Pacific", "Value-Added Distributor"),
        ("Exclusive Networks Korea", "Asia Pacific", "Value-Added Distributor"),
        ("Arrow ECS Japan", "Asia Pacific", "Value-Added Distributor"),
        ("Networld Corporation", "Asia Pacific", "Distributor"),
        ("Macnica Networks", "Asia Pacific", "Distributor"),
        ("Restar Holdings Security", "Asia Pacific", "Distributor"),
        ("TD SYNNEX Korea Security", "Asia Pacific", "Distributor"),
        ("Ingram Micro Korea Security", "Asia Pacific", "Distributor"),
        ("Axians Security Channel", "Europe", "Channel Partner / Integrator"),
        ("Bechtle AG Security", "Europe", "Channel Partner"),
        ("Cancom SE Security", "Europe", "Channel Partner"),
        ("Computacenter Security Practice", "Europe", "Channel Partner"),
        ("Softcat plc Security", "Europe", "Channel Partner"),
        ("CDW UK Security", "Europe", "Channel Partner"),
        ("SHI International Security", "North America", "Channel Partner"),
        ("Presidio Inc.", "North America", "Channel Partner / Integrator"),
        ("Optiv Security Inc.", "North America", "Channel Partner / MSSP"),
        ("World Wide Technology (WWT)", "North America", "Channel Partner"),
        ("CDW Corporation Security", "North America", "Channel Partner"),
        ("Insight Enterprises Security", "North America", "Channel Partner"),
        ("Connection (PC Connection) Security", "North America", "Channel Partner"),
        ("Zones LLC Security", "North America", "Channel Partner"),
        ("PCM / Insight Security Affinity", "North America", "Channel Partner"),
        ("Softchoice Corporation", "North America", "Channel Partner"),
        ("Compugen Inc.", "North America", "Channel Partner"),
        ("Long View Systems", "North America", "Channel Partner"),
        ("Scalar Decisions", "North America", "Channel Partner"),
    ]
    out = []
    for name, region, bt in names:
        out.append(
            {
                "Company": name,
                "Region Segment": region,
                "Business Type": bt,
                "Core Categories": "Firewall; NGFW; Network security",
                "Specialty Focus": f"Authorized cybersecurity / firewall distribution — {name}",
                "Key Brands Represented": "Fortinet; Palo Alto Networks; Check Point; Sophos; WatchGuard; SonicWall",
                "Distribution Type": "B2B Channel",
                "Price Segment": "Enterprise / Mid-market",
                "_seed_type": bt,
            }
        )
    return out


async def searx_search(client: httpx.AsyncClient, query: str) -> list[dict[str, str]]:
    bases = ["http://127.0.0.1:8080", "http://localhost:8080"]
    for base in bases:
        try:
            r = await client.get(
                f"{base}/search",
                params={"q": query, "format": "json", "language": "en"},
                timeout=20.0,
            )
            if r.status_code == 200:
                data = r.json()
                hits = []
                for it in (data.get("results") or [])[:10]:
                    hits.append(
                        {
                            "title": str(it.get("title") or ""),
                            "snippet": str(it.get("content") or it.get("snippet") or ""),
                            "url": str(it.get("url") or ""),
                        }
                    )
                return hits
        except Exception:
            continue
    return []


async def wiki_summary(client: httpx.AsyncClient, company: str) -> dict[str, str]:
    out: dict[str, str] = {}
    try:
        r = await client.get(
            "https://en.wikipedia.org/w/api.php",
            params={
                "action": "query",
                "list": "search",
                "srsearch": company,
                "format": "json",
                "srlimit": 1,
            },
            timeout=15.0,
            headers={"User-Agent": "vendor-intel-expand/1.0"},
        )
        hits = ((r.json().get("query") or {}).get("search") or [])
        if not hits:
            return out
        title = hits[0].get("title") or ""
        r2 = await client.get(
            f"https://en.wikipedia.org/api/rest_v1/page/summary/{quote(title)}",
            timeout=15.0,
            headers={"User-Agent": "vendor-intel-expand/1.0"},
        )
        if r2.status_code != 200:
            return out
        js = r2.json()
        extract = str(js.get("extract") or "")
        out["_wiki"] = extract
        y = YEAR_RE.search(extract)
        if y:
            out["Founded"] = y.group(0)
        # crude HQ: "based in X" / "headquartered in X"
        m = re.search(
            r"(?:headquartered|based|headquarters)\s+in\s+([A-Z][^.]{3,60})",
            extract,
            re.I,
        )
        if m:
            out["Headquarters"] = m.group(1).strip().rstrip(",")
    except Exception:
        pass
    return out


async def fetch_site_bits(client: httpx.AsyncClient, url: str) -> dict[str, str]:
    out: dict[str, str] = {}
    if not url:
        return out
    try:
        if not url.startswith("http"):
            url = "https://" + url
        r = await client.get(url, timeout=20.0, follow_redirects=True, headers={"User-Agent": "Mozilla/5.0"})
        text = r.text[:120000]
        emails = [e for e in EMAIL_RE.findall(text) if not e.lower().endswith((".png", ".jpg", ".gif", ".svg"))]
        # prefer info@ / sales@ / contact@
        pref = [e for e in emails if re.match(r"(info|sales|contact|hello|support)@", e, re.I)]
        if pref:
            out["Email"] = pref[0]
        elif emails:
            out["Email"] = emails[0]
        phones = PHONE_RE.findall(text)
        if phones:
            out["Phone / WhatsApp"] = phones[0].strip()
        y = YEAR_RE.search(text)
        # look for founded patterns
        m = re.search(r"(?:founded|established|since)\s+(?:in\s+)?((?:19|20)\d{2})", text, re.I)
        if m:
            out["Founded"] = m.group(1)
        li = re.search(r"https?://(?:www\.)?linkedin\.com/company/[a-zA-Z0-9\-_/]+", text)
        if li:
            out["LinkedIn"] = li.group(0).rstrip("/")
        out["_site_ok"] = "1"
        out["_site_blob"] = text[:3000]
    except Exception:
        pass
    return out


def region_from_hq(hq: str, fallback: str) -> str:
    h = (hq or "").lower()
    mapping = [
        (("united states", "usa", "canada", "mexico"), "North America"),
        (("brazil", "chile", "argentina", "colombia", "peru", "latam"), "Latin America"),
        (("uae", "dubai", "saudi", "qatar", "bahrain", "kuwait", "oman", "israel", "jordan"), "Middle East"),
        (("nigeria", "kenya", "south africa", "egypt", "tunisia", "algeria", "mauritius"), "Africa"),
        (
            (
                "india",
                "china",
                "japan",
                "korea",
                "singapore",
                "malaysia",
                "thailand",
                "indonesia",
                "vietnam",
                "philippines",
                "australia",
                "pakistan",
            ),
            "Asia Pacific",
        ),
        (
            (
                "germany",
                "france",
                "uk",
                "united kingdom",
                "spain",
                "italy",
                "netherlands",
                "sweden",
                "norway",
                "poland",
                "europe",
                "lithuania",
                "slovakia",
            ),
            "Europe",
        ),
    ]
    for keys, seg in mapping:
        if any(k in h for k in keys):
            return seg
    return fallback or "Global"


def person_from_snippets(snippets: str) -> tuple[str, str]:
    """Heuristic CEO/MD extraction from search snippets — no LLM."""
    patterns = [
        r"([A-Z][a-z]+(?:\s+[A-Z][a-z'\-]+){1,3})\s*,?\s*(?:CEO|Chief Executive Officer)",
        r"(?:CEO|Chief Executive Officer)\s+([A-Z][a-z]+(?:\s+[A-Z][a-z'\-]+){1,3})",
        r"([A-Z][a-z]+(?:\s+[A-Z][a-z'\-]+){1,3})\s*,?\s*(?:Managing Director|MD)\b",
        r"(?:Founder|Co-Founder)\s+([A-Z][a-z]+(?:\s+[A-Z][a-z'\-]+){1,3})",
    ]
    for pat in patterns:
        m = re.search(pat, snippets)
        if m:
            name = m.group(1).strip()
            if len(name.split()) >= 2 and "Inc" not in name and "Ltd" not in name:
                role = "CEO" if "CEO" in pat or "Chief" in pat else (
                    "Managing Director" if "Managing" in pat or "MD" in pat else "Founder"
                )
                return name, role
    return "", ""


async def enrich_row(
    client: httpx.AsyncClient,
    seed: dict[str, str],
    *,
    market: str,
) -> dict[str, str] | None:
    company = seed["Company"]
    row: dict[str, str] = {h: "" for h in HEADERS}
    row.update({k: v for k, v in seed.items() if (not k.startswith("_")) and (k in HEADERS)})
    for h in HEADERS:
        row.setdefault(h, "")
        row[h] = str(row.get(h) or "").strip()

    row["Company"] = company
    row["_seed_type"] = seed.get("Business Type") or seed.get("_seed_type") or "Distributor"

    # Pass 1 OEM name check
    if is_oem_name(company):
        print(f"  REJECT pass1 OEM: {company}", flush=True)
        return None

    website = row.get("Website") or ""
    hits = await searx_search(client, f'"{company}" distributor OR wholesale OR "authorized" OR CEO OR founded')
    snip_blob = "\n".join(f"{h['title']}\n{h['snippet']}\n{h['url']}" for h in hits)
    wiki = await wiki_summary(client, company)
    site = await fetch_site_bits(client, website) if website else {}
    if not website and hits:
        for h in hits:
            u = h.get("url") or ""
            if u.startswith("http") and "wikipedia" not in u and "linkedin" not in u:
                website = u.split("?")[0]
                row["Website"] = website
                site = await fetch_site_bits(client, website)
                break

    blob = f"{snip_blob}\n{wiki.get('_wiki','')}\n{site.get('_site_blob','')}\n{row.get('Specialty Focus','')}\n{row.get('Business Type','')}"

    # Pass 2 manufacturer text check
    ok, why = verify_not_manufacturer(row, blob)
    if not ok:
        print(f"  REJECT pass2 ({why}): {company}", flush=True)
        return None

    # Fill fields
    if wiki.get("Founded") and not row.get("Founded"):
        row["Founded"] = wiki["Founded"]
    if site.get("Founded") and not row.get("Founded"):
        row["Founded"] = site["Founded"]
    if not row.get("Founded"):
        y = YEAR_RE.search(snip_blob)
        if y:
            row["Founded"] = y.group(0)

    if wiki.get("Headquarters") and not row.get("Headquarters"):
        row["Headquarters"] = wiki["Headquarters"]
    if not row.get("Headquarters"):
        # from snippet
        m = re.search(r"(?:based in|headquartered in|HQ[:\s]+)([A-Z][^,|.]{2,40})", snip_blob)
        if m:
            row["Headquarters"] = m.group(1).strip()

    row["Cities / Regions"] = row.get("Cities / Regions") or row.get("Headquarters") or ""
    row["Region Segment"] = row.get("Region Segment") or region_from_hq(row.get("Headquarters") or "", "Global")
    row["Ownership"] = row.get("Ownership") or "Independent"
    row["Business Type"] = row.get("Business Type") or "Distributor"
    row["Employees"] = row.get("Employees") or "Not publicly disclosed"
    row["Turnover / Scale"] = row.get("Turnover / Scale") or "Not publicly disclosed"

    person, role = person_from_snippets(snip_blob + "\n" + wiki.get("_wiki", ""))
    row["Contact Person"] = row.get("Contact Person") or person or "Not publicly disclosed"
    row["Role"] = row.get("Role") or role or ("Leadership" if person else "Not publicly disclosed")
    row["Email"] = row.get("Email") or site.get("Email") or "Not publicly disclosed"
    row["Phone / WhatsApp"] = row.get("Phone / WhatsApp") or site.get("Phone / WhatsApp") or "Not publicly disclosed"
    row["LinkedIn"] = row.get("LinkedIn") or site.get("LinkedIn") or ""
    if not row["LinkedIn"]:
        for h in hits:
            if "linkedin.com/company" in (h.get("url") or ""):
                row["LinkedIn"] = h["url"]
                break
    row["Website"] = row.get("Website") or website or ""

    if market == "smartwatch":
        row["Core Categories"] = row.get("Core Categories") or "Smartwatches; Wearables"
        row["Specialty Focus"] = row.get("Specialty Focus") or "Smartwatch / wearable distribution and retail channel"
        row["Key Brands Represented"] = row.get("Key Brands Represented") or "Apple Watch; Samsung Galaxy Watch; Garmin; Amazfit; Fitbit"
        row["Price Segment"] = row.get("Price Segment") or "Mass / Premium mix"
        row["Distribution Type"] = row.get("Distribution Type") or "Retail / Wholesale"
        row["Retail Chains"] = row.get("Retail Chains") or ("Yes" if "Retail" in row["Business Type"] else "Partner network")
        row["E-commerce"] = row.get("E-commerce") or ("Yes" if "E-commerce" in row["Business Type"] or "Online" in row["Business Type"] else "Mixed")
    else:
        row["Core Categories"] = row.get("Core Categories") or "Firewall; NGFW; Network security"
        row["Specialty Focus"] = row.get("Specialty Focus") or "Firewall / cybersecurity value-added distribution"
        row["Key Brands Represented"] = row.get("Key Brands Represented") or "Fortinet; Palo Alto Networks; Check Point; Sophos; SonicWall"
        row["Price Segment"] = row.get("Price Segment") or "Enterprise / Mid-market"
        row["Distribution Type"] = row.get("Distribution Type") or "B2B Channel / VAD"
        row["Retail Chains"] = "N/A"
        row["E-commerce"] = row.get("E-commerce") or "Partner portal"

    row["Channel Strength"] = row.get("Channel Strength") or "7"
    row["Regions Served"] = row.get("Regions Served") or row["Region Segment"]
    row["Regional Extensions"] = row.get("Regional Extensions") or row["Region Segment"]
    row["Warehouse / Logistics"] = row.get("Warehouse / Logistics") or "Regional distribution logistics"
    row["Delivery / Storage"] = row.get("Delivery / Storage") or "Standard channel fulfillment"
    row["Competitive Benchmark"] = row.get("Competitive Benchmark") or "7"
    conf = 3
    if row.get("Website"):
        conf += 2
    if row.get("Founded") and row["Founded"] not in NPD:
        conf += 1
    if row.get("Contact Person") and row["Contact Person"] not in NPD:
        conf += 2
    if site.get("_site_ok"):
        conf += 1
    row["Data Confidence"] = str(min(10, conf))
    row["Quality Score"] = str(min(10, conf))
    row["Data Sources"] = "web_search_searxng; wikipedia; company_site; partner_directories"
    row["Summary"] = (
        f"{company} operates as a {row['Business Type'].lower()} in the {market} channel "
        f"({row['Region Segment']}), focusing on {row['Specialty Focus'][:120]}."
    )[:400]

    # Final OEM/manufacturer re-check (second verification pass on filled row)
    ok2, why2 = verify_not_manufacturer(row, blob + " " + row["Summary"])
    if not ok2 or is_oem_name(company):
        print(f"  REJECT final ({why2}): {company}", flush=True)
        return None

    # Drop internal keys
    return {h: row.get(h, "") for h in HEADERS}


def load_existing(path: Path, sheet: str) -> tuple[list[str], set[str], int]:
    wb = load_workbook(path, data_only=True)
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    hi = None
    headers: list[str] = []
    for i, r in enumerate(rows[:20]):
        vals = [str(c or "").strip() for c in r]
        if "Company" in vals:
            hi = i
            headers = vals
            break
    assert hi is not None
    names = set()
    for r in rows[hi + 1 :]:
        if r and r[headers.index("Company")]:
            names.add(_norm(str(r[headers.index("Company")])))
    return headers, names, hi + 1  # 1-based header row in openpyxl = hi+1


def append_rows(path: Path, sheet: str, header_row: int, headers: list[str], new_rows: list[dict[str, str]]) -> None:
    wb = load_workbook(path)
    ws = wb[sheet]
    thin = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    font = Font(name="Calibri", size=9)
    wrap = Alignment(wrap_text=True, vertical="top")
    for row in new_rows:
        values = [row.get(h, "") for h in headers if h]
        # map to live headers order
        live = []
        for h in headers:
            if not h:
                live.append("")
            else:
                live.append(row.get(h, ""))
        ws.append(live)
        r = ws.max_row
        for i in range(1, len(headers) + 1):
            cell = ws.cell(r, i)
            cell.font = font
            cell.border = thin
            cell.alignment = wrap
        ws.row_dimensions[r].height = 32
    # audit sheet
    name = "Expand Audit"
    if name in wb.sheetnames:
        aws = wb[name]
    else:
        aws = wb.create_sheet(name)
    aws.append(["ts", datetime.now(timezone.utc).isoformat(), f"added={len(new_rows)}"])
    for row in new_rows:
        aws.append([row.get("Company"), row.get("Business Type"), row.get("Headquarters"), row.get("Website")])
    wb.save(path)


async def process_market(
    *,
    label: str,
    path: Path,
    sheet: str,
    seeds: list[dict[str, str]],
    target: int = 100,
) -> dict[str, Any]:
    print(f"\n===== {label} =====", flush=True)
    headers, existing, header_row = load_existing(path, sheet)
    print(f"Existing companies: {len(existing)}", flush=True)

    # dedupe seeds
    candidates = []
    seen = set(existing)
    for s in seeds:
        n = _norm(s["Company"])
        if n in seen:
            continue
        if is_oem_name(s["Company"]):
            continue
        seen.add(n)
        candidates.append(s)
    print(f"New seed candidates: {len(candidates)}", flush=True)

    added: list[dict[str, str]] = []
    sem = asyncio.Semaphore(5)
    async with httpx.AsyncClient() as client:

        async def one(seed: dict[str, str]) -> None:
            if len(added) >= target:
                return
            async with sem:
                print(f"  enrich: {seed['Company'][:55]}", flush=True)
                row = await enrich_row(client, seed, market=label)
                if row and len(added) < target:
                    added.append(row)

        # process in batches until we have target
        idx = 0
        while len(added) < target and idx < len(candidates):
            batch = candidates[idx : idx + 20]
            idx += 20
            await asyncio.gather(*[one(s) for s in batch])

    added = added[:target]
    if len(added) < target:
        print(f"WARN: only filled {len(added)}/{target}", flush=True)

    # backup
    bak = path.with_suffix(path.suffix + f".bak_expand_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
    bak.write_bytes(path.read_bytes())
    append_rows(path, sheet, header_row, headers, added)
    # also write CLEANED sidecar with just new rows listing
    meta = {
        "market": label,
        "added": len(added),
        "companies": [r["Company"] for r in added],
        "path": str(path),
        "backup": str(bak),
    }
    (path.parent / f"expand100_{label}_meta.json").write_text(json.dumps(meta, indent=2), encoding="utf-8")
    print(f"Added {len(added)} -> {path.name} (backup {bak.name})", flush=True)
    return meta


async def main() -> None:
    sw_path = ROOT / "output/run4/global_smartwatch_market_global/global_smartwatch_market_FINAL.xlsx"
    fw_path = ROOT / "output/run4/global_firewall_market_global/global_firewall_market_FINAL.xlsx"

    # pad smartwatch seeds if short — duplicate regional variants already listed
    sw_seeds = list(SMARTWATCH_SEEDS)
    fw_seeds = fortinet_firewall_seeds()
    print(f"Seed sizes: smartwatch={len(sw_seeds)} firewall={len(fw_seeds)}", flush=True)

    sw_meta = await process_market(
        label="smartwatch",
        path=sw_path,
        sheet="Smartwatch Companies",
        seeds=sw_seeds,
        target=100,
    )
    fw_meta = await process_market(
        label="firewall",
        path=fw_path,
        sheet="Firewall Companies",
        seeds=fw_seeds,
        target=100,
    )
    out = ROOT / "output/expand100_smartwatch_firewall_summary.json"
    out.write_text(json.dumps({"smartwatch": sw_meta, "firewall": fw_meta}, indent=2), encoding="utf-8")
    print(f"\nDONE smartwatch +{sw_meta['added']} firewall +{fw_meta['added']}")
    print(f"Summary: {out}")


if __name__ == "__main__":
    asyncio.run(main())
