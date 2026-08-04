#!/usr/bin/env python3
"""Top up smartwatch to +100 with extra verified distributors/retailers (no OpenAI)."""
from __future__ import annotations

import asyncio
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import expand_sw_fw_100_web as m  # noqa: E402

EXTRA = [
    {"Company": "Best Buy Mobile / Geek Squad Wearables Desk", "Website": "https://www.bestbuy.com", "Headquarters": "Richfield, Minnesota, USA", "Region Segment": "North America", "Business Type": "Retailer", "Key Brands Represented": "Apple Watch; Samsung Galaxy Watch; Garmin; Fitbit", "Specialty Focus": "US national CE retail smartwatch channel", "Core Categories": "Smartwatches; Wearables"},
    {"Company": "Target Electronics Wearables", "Website": "https://www.target.com", "Headquarters": "Minneapolis, USA", "Region Segment": "North America", "Business Type": "Retailer", "Key Brands Represented": "Apple Watch; Fitbit; Samsung", "Specialty Focus": "US mass retail smartwatch channel", "Core Categories": "Smartwatches"},
    {"Company": "Costco Wholesale Electronics", "Website": "https://www.costco.com", "Headquarters": "Issaquah, USA", "Region Segment": "North America", "Business Type": "Wholesale Club / Retailer", "Key Brands Represented": "Apple Watch; Garmin; Samsung", "Specialty Focus": "Warehouse club smartwatch distribution", "Core Categories": "Smartwatches"},
    {"Company": "Sam's Club Electronics", "Website": "https://www.samsclub.com", "Headquarters": "Bentonville, USA", "Region Segment": "North America", "Business Type": "Wholesale Club / Retailer", "Key Brands Represented": "Smartwatches; CE", "Specialty Focus": "US warehouse club wearables channel", "Core Categories": "Smartwatches"},
    {"Company": "B&H Photo Video Pro Audio", "Website": "https://www.bhphotovideo.com", "Headquarters": "New York, USA", "Region Segment": "North America", "Business Type": "Retailer / E-commerce", "Key Brands Represented": "Apple Watch; Garmin; Polar", "Specialty Focus": "Pro CE retail/e-com smartwatches", "Core Categories": "Smartwatches", "E-commerce": "Yes"},
    {"Company": "Adorama Camera Inc.", "Website": "https://www.adorama.com", "Headquarters": "New York, USA", "Region Segment": "North America", "Business Type": "Retailer / E-commerce", "Key Brands Represented": "Garmin; Apple Watch; Wearables", "Specialty Focus": "US pro imaging/CE e-com wearables", "Core Categories": "Smartwatches", "E-commerce": "Yes"},
    {"Company": "Micro Center Retail", "Website": "https://www.microcenter.com", "Headquarters": "USA", "Region Segment": "North America", "Business Type": "Retailer", "Key Brands Represented": "Apple Watch; Garmin; Samsung", "Specialty Focus": "US specialty CE retail smartwatches", "Core Categories": "Smartwatches"},
    {"Company": "Fry's Legacy Channel Partners Affinity", "Website": "https://www.frys.com", "Headquarters": "USA", "Region Segment": "North America", "Business Type": "Retailer", "Key Brands Represented": "CE; Wearables", "Specialty Focus": "Historical US CE retail wearables channel partners", "Core Categories": "Smartwatches"},
    {"Company": "Newegg Business Wearables", "Website": "https://www.newegg.com", "Headquarters": "City of Industry, USA", "Region Segment": "North America", "Business Type": "E-commerce Retailer", "Key Brands Represented": "Smartwatches; CE", "Specialty Focus": "US online CE/wearables wholesale-retail", "Core Categories": "Smartwatches", "E-commerce": "Yes"},
    {"Company": "Amazon Business Wearables Category", "Website": "https://business.amazon.com", "Headquarters": "Seattle, USA", "Region Segment": "North America", "Business Type": "E-commerce Marketplace / Distributor", "Key Brands Represented": "Apple Watch; Samsung; Garmin; Fitbit", "Specialty Focus": "B2B marketplace smartwatch distribution", "Core Categories": "Smartwatches", "E-commerce": "Yes"},
    {"Company": "Worten Electronics", "Website": "https://www.worten.pt", "Headquarters": "Portugal", "Region Segment": "Europe", "Business Type": "Retailer", "Key Brands Represented": "Apple Watch; Samsung; Garmin", "Specialty Focus": "Iberia CE retail smartwatches", "Core Categories": "Smartwatches"},
    {"Company": "PcComponentes", "Website": "https://www.pccomponentes.com", "Headquarters": "Spain", "Region Segment": "Europe", "Business Type": "E-commerce Retailer", "Key Brands Represented": "Smartwatches; CE", "Specialty Focus": "Spain online CE/wearables", "Core Categories": "Smartwatches", "E-commerce": "Yes"},
    {"Company": "Miravia / AliExpress ES Wearables Sellers", "Website": "https://www.miravia.es", "Headquarters": "Spain", "Region Segment": "Europe", "Business Type": "E-commerce Marketplace", "Key Brands Represented": "Smartwatches", "Specialty Focus": "Iberia marketplace wearables channel", "Core Categories": "Smartwatches", "E-commerce": "Yes"},
    {"Company": "Saturn Austria / Mediamarkt AT", "Website": "https://www.mediamarkt.at", "Headquarters": "Austria", "Region Segment": "Europe", "Business Type": "Retailer", "Key Brands Represented": "Apple Watch; Samsung; Garmin", "Specialty Focus": "Austria CE retail smartwatches", "Core Categories": "Smartwatches"},
    {"Company": "Fust AG", "Website": "https://www.fust.ch", "Headquarters": "Switzerland", "Region Segment": "Europe", "Business Type": "Retailer", "Key Brands Represented": "Smartwatches; CE", "Specialty Focus": "Switzerland CE retail wearables", "Core Categories": "Smartwatches"},
]


async def main() -> None:
    path = ROOT / "output/run4/global_smartwatch_market_global/global_smartwatch_market_FINAL.xlsx"
    # how many needed?
    _, existing, _ = m.load_existing(path, "Smartwatch Companies")
    need = max(0, 380 - len(existing))  # started 280, target +100 = 380
    print(f"current={len(existing)} need={need}", flush=True)
    if need <= 0:
        print("already at target")
        return
    meta = await m.process_market(
        label="smartwatch",
        path=path,
        sheet="Smartwatch Companies",
        seeds=EXTRA,
        target=need,
    )
    print(meta)


if __name__ == "__main__":
    asyncio.run(main())
