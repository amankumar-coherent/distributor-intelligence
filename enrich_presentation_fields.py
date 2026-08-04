#!/usr/bin/env python3
import asyncio
import json
import os
import sys
import time
import glob
from pathlib import Path
from openai import AsyncOpenAI

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT))

from backend.config import get_settings
from vendor_intel.export.fmcg_landscape import export_fmcg_xlsx
from vendor_intel.pipeline.orchestrator import save_pipeline_csv

# Load settings to verify API config
cfg = get_settings()
if not cfg.openai_api_key:
    print("ERROR: OPENAI_API_KEY is not set in environment or .env file.")
    sys.exit(1)

client = AsyncOpenAI(api_key=cfg.openai_api_key)
model_name = cfg.model_extract or "gpt-4o-mini"

# Concurrency limit for OpenAI calls
SEM_LIMIT = 10
semaphore = asyncio.Semaphore(SEM_LIMIT)

def get_market_name(result: dict, default_val: str) -> str:
    ctx = result.get("query_context") or {}
    name = ctx.get("industry") or result.get("query")
    if name:
        return str(name).replace("Global ", "").replace(" Market", "").strip()
    return default_val

async def enrich_company_presentation(company_dict: dict, market_name: str, idx: int, total: int, json_path: Path, lock: asyncio.Lock):
    name = company_dict.get("company") or company_dict.get("brand") or "Unknown"
    domain = company_dict.get("domain") or company_dict.get("website") or ""
    
    # If already enriched, skip
    if company_dict.get("_presentation_enrich") and company_dict["_presentation_enrich"].get("ownership"):
        print(f"  [{idx}/{total}] Skipping {name} (already presentation-enriched)")
        return

    _crawl = company_dict.get("_crawl") or {}
    crawl_text = _crawl.get("crawl_text") or ""
    summary = company_dict.get("company_summary") or company_dict.get("polished_summary") or company_dict.get("summary") or ""
    products = company_dict.get("key_products") or _crawl.get("products") or ""
    raw_specialty = _crawl.get("certifications") or _crawl.get("industries") or ""

    if len(crawl_text) > 8000:
        crawl_text = crawl_text[:8000] + " ... [truncated]"

    system_prompt = (
        "You are an expert corporate researcher and market analyst.\n"
        "Your task is to extract precise corporate structure, sales channels, and product focus details for a company.\n"
        "CRITICAL RULE: Extract information ONLY if it is explicitly stated in the provided scraped website text snippet. "
        "Do NOT use general knowledge, do NOT guess, and do NOT make assumptions. If a specific detail is not directly "
        "found in the text, you must return null or 'No' / 'Private' as requested. Zero hallucination is allowed."
    )

    user_prompt = f"""Company Name: {name}
Website Domain: {domain}
Market Area: {market_name}
Current Products: {products}
Current Summary: {summary}
Raw Certs/Industries: {raw_specialty}

Scraped Website Text (snippet):
{crawl_text}

JSON Output Format (Strictly return ONLY this JSON object):
{{
  "ownership": "Corporate ownership. Is it a Subsidiary of another firm (specify e.g. 'Subsidiary of [Parent Name]'), Publicly Traded (specify e.g. 'Public (NYSE: [TICKER])'), or Private? If unknown, choose 'Private'.",
  "ecommerce": "Specific online store names or platforms (e.g. 'Company Online Store, Amazon, Shopify Store, Walmart.com') if they sell directly to consumers online. If none or not applicable, output 'No'.",
  "foodservice": "Specific foodservice distributors, restaurant chains, hotel chains, or catering clients (e.g. 'Sysco, US Foods, Marriott Hotels, regional restaurants') if they supply the foodservice/HoReCa channel. If none or not applicable, output 'No'.",
  "retail_chains": "Specific names of major grocery or retail chains where their products are sold (e.g. 'Whole Foods, Costco, Kroger, Target, Safeway'). If none or not applicable, output 'No'.",
  "specialty_grocers": "Specific names of specialty, organic, or health grocers where their products are sold (e.g. 'Trader Joe's, Sprouts Farmers Market, local health food co-ops'). If none or not applicable, output 'No'.",
  "specialty_focus": "The specific product focus, technical focus, or niche of the company in this market (e.g. 'Organic extra virgin avocado oil', 'OT/CPS cybersecurity systems', 'Reusable gel ice packs'). Do NOT include quality standards like ISO 9001, HACCP, GMP, etc."
}}
"""

    async with semaphore:
        for attempt in range(2):
            try:
                response = await client.chat.completions.create(
                    model=model_name,
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    response_format={"type": "json_object"},
                    temperature=0.0,
                    max_tokens=600
                )
                content = response.choices[0].message.content
                res = json.loads(content)

                # Save values back into the company dict
                company_dict["_presentation_enrich"] = {
                    "ownership": res.get("ownership") or "Private",
                    "ecommerce": res.get("ecommerce") or "No",
                    "foodservice": res.get("foodservice") or "No",
                    "retail_chains": res.get("retail_chains") or "No",
                    "specialty_grocers": res.get("specialty_grocers") or "No",
                    "specialty_focus": res.get("specialty_focus") or ""
                }

                print(f"  [{idx}/{total}] Enriched {name}: Ownership={res.get('ownership')}, Specialty={res.get('specialty_focus')}")

                # Save checkpoint to JSON
                async with lock:
                    try:
                        with open(json_path, "r", encoding="utf-8") as f:
                            current_data = json.load(f)
                        # Find and update the company in the list
                        for c in current_data.get("relevant_companies", []):
                            if (c.get("company") == company_dict.get("company") or 
                                c.get("domain") == company_dict.get("domain")):
                                c.update(company_dict)
                                break
                        with open(json_path, "w", encoding="utf-8") as f:
                            json.dump(current_data, f, indent=2, default=str)
                    except Exception as ex:
                        print(f"Error saving JSON checkpoint: {ex}")
                return
            except Exception as e:
                print(f"  [{idx}/{total}] Error enriching {name} (attempt {attempt+1}/2): {e}")
                await asyncio.sleep(2.0)

async def enrich_market_file(json_path: Path):
    print(f"\n========================================================")
    print(f"ENRICHING PRESENTATION FIELDS: {json_path.parent.name}")
    print(f"File: {json_path.name}")
    print(f"========================================================")

    try:
        with open(json_path, "r", encoding="utf-8") as f:
            result = json.load(f)
    except Exception as e:
        print(f"Failed to load JSON {json_path}: {e}")
        return

    market_name = get_market_name(result, json_path.parent.name.replace("_", " ").title())
    companies = result.get("relevant_companies") or []
    todo_companies = [c for c in companies if c.get("is_relevant", True)]
    total = len(todo_companies)
    if total == 0:
        print("No relevant companies to enrich. Skipping.")
        return

    print(f"Found {total} relevant companies to enrich.")
    
    lock = asyncio.Lock()
    tasks = []
    for i, company in enumerate(todo_companies):
        tasks.append(enrich_company_presentation(company, market_name, i + 1, total, json_path, lock))

    await asyncio.gather(*tasks)

    # Reload full updated data to run exports
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            updated_result = json.load(f)
        
        print("Regenerating Excel export with enriched presentation details...")
        xlsx_path = json_path.with_suffix(".xlsx")
        export_fmcg_xlsx(updated_result, xlsx_path)
        print(f"Saved Excel: {xlsx_path.resolve()}")

        print("Regenerating CSV export...")
        csv_path = json_path.with_suffix(".csv")
        save_pipeline_csv(updated_result, str(csv_path))
        print(f"Saved CSV: {csv_path.resolve()}")
    except Exception as e:
        print(f"Failed to save final JSON or export: {e}")

async def main():
    start_time = time.perf_counter()
    
    json_files = glob.glob("output/*/global_*.json")
    if not json_files:
        print("No global JSON files found under output/*/")
        return

    print(f"Found {len(json_files)} market file(s) to enrich: {[Path(p).name for p in json_files]}")

    for json_path in json_files:
        await enrich_market_file(Path(json_path))

    elapsed = time.perf_counter() - start_time
    print(f"\n========================================================")
    print(f"ALL PRESENTATION ENRICHMENTS COMPLETED in {elapsed/60.0:.2f} minutes.")
    print(f"========================================================")

if __name__ == "__main__":
    if sys.platform == 'win32':
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    asyncio.run(main())
