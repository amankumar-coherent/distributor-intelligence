#!/usr/bin/env python3
import asyncio
import json
import os
import sys
import time
import glob
import re
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

SEM_LIMIT = 10
semaphore = asyncio.Semaphore(SEM_LIMIT)

def get_market_name(result: dict, default_val: str) -> str:
    ctx = result.get("query_context") or {}
    name = ctx.get("industry") or result.get("query")
    if name:
        return str(name).replace("Global ", "").replace(" Market", "").strip()
    return default_val

async def audit_company(company_dict: dict, market_name: str, idx: int, total: int):
    name = company_dict.get("company") or company_dict.get("brand") or "Unknown"
    domain = company_dict.get("domain") or company_dict.get("website") or ""
    
    # If already audited, skip
    if company_dict.get("audit_completed"):
        print(f"  [{idx}/{total}] Skipping {name} (already audited)")
        return
        
    crawl_text = ""
    _crawl = company_dict.get("_crawl") or {}
    if _crawl:
        crawl_text = _crawl.get("crawl_text") or ""
    
    # Trim crawl text if too long to save tokens
    if len(crawl_text) > 8000:
        crawl_text = crawl_text[:8000] + " ... [truncated]"
        
    current_products = company_dict.get("key_products") or _crawl.get("products") or ""
    summary = company_dict.get("company_summary") or company_dict.get("polished_summary") or company_dict.get("summary") or ""
    
    system_prompt = (
        "You are an expert market research analyst auditing a list of companies to filter out noise and ensure high data quality.\n"
        f"The target market is: \"{market_name}\"\n\n"
        "Your task is to audit the company and determine if it is a relevant participant in the target market.\n"
        "A company is a relevant participant if it:\n"
        "1. Manufactures, designs, or develops products/systems directly for the target market.\n"
        "2. Supplies specialized components, raw materials, or sub-systems designed for the target market.\n"
        "3. Integrates or compiles systems specifically for the target market.\n"
        "4. Directly distributes, retail sells, or rents out products in this market.\n"
        "5. Provides specialized engineering, testing, or maintenance services for this market.\n\n"
        "A company is NOT relevant (noise) if it:\n"
        "1. Is completely unrelated (e.g. sports blog, tickets site, consumer retailer, finance/investment fund, generic real estate).\n"
        "2. Is a software/IT consulting or general services company that has no specialized offering for this market.\n"
        "3. Is a generic machine shop or fabricator that does not specifically focus on or offer products/services for this market.\n"
        "4. Is a research paper, article, or market research directory rather than an actual company.\n\n"
        "Evaluate the company based on the provided website crawl, products, and your general knowledge.\n"
        "You MUST respond with a raw JSON object containing the exact structure below."
    )
    
    user_prompt = f"""Company Name: {name}
Website Domain: {domain}
Current Summary: {summary}
Current Products: {current_products}

Scraped Website Text (snippet):
{crawl_text}

JSON Output Format (Strictly return ONLY this JSON object):
{{
  "is_relevant": true / false,
  "relevance_score": integer from 0 to 5,
  "role": "Manufacturer" / "Distributor" / "Supplier" / "Integrator" / "Service Provider" / "Unrelated",
  "reasoning": "A concise explanation of why the company is relevant or irrelevant, citing specific products/activities in the target market.",
  "corrected_products": "A comma-separated list of 2-5 specific products they offer in this market (or null if irrelevant or none). Keep it brief.",
  "corrected_hq": "Their headquarters location (City, State, Country or City, Country) if you can verify it, otherwise null.",
  "corrected_founded_year": integer representing the year founded (or null if unknown),
  "corrected_employees": "approximate number of employees or range (e.g. '10-50', '100-500', '1000+') if known, otherwise null."
}}
"""

    async with semaphore:
        for attempt in range(3):
            try:
                response = await client.chat.completions.create(
                    model=model_name,
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    response_format={"type": "json_object"},
                    temperature=0.1,
                    max_tokens=800
                )
                content = response.choices[0].message.content
                res = json.loads(content)
                
                # Apply changes to company_dict
                relevance_score = res.get("relevance_score", 0)
                is_relevant = res.get("is_relevant", False)
                
                # We enforce is_relevant is true only if relevance_score >= 3
                if relevance_score < 3:
                    is_relevant = False
                
                company_dict["is_relevant"] = is_relevant
                company_dict["role"] = res.get("role") or company_dict.get("role") or "Other"
                company_dict["confidence"] = round(relevance_score / 5.0, 2)
                company_dict["audit_reasoning"] = res.get("reasoning", "")
                company_dict["audit_score"] = relevance_score
                
                if is_relevant:
                    company_dict["is_relevant"] = True
                    # Update products
                    corr_prod = res.get("corrected_products")
                    if corr_prod:
                        company_dict["key_products"] = corr_prod
                        if "_crawl" not in company_dict:
                            company_dict["_crawl"] = {}
                        company_dict["_crawl"]["products"] = corr_prod
                    
                    # Update HQ
                    corr_hq = res.get("corrected_hq")
                    if corr_hq:
                        if "_crawl" not in company_dict:
                            company_dict["_crawl"] = {}
                        company_dict["_crawl"]["headquarters"] = corr_hq
                        
                    # Update Founded
                    corr_founded = res.get("corrected_founded_year")
                    if corr_founded:
                        if "_crawl" not in company_dict:
                            company_dict["_crawl"] = {}
                        company_dict["_crawl"]["founded_year"] = int(corr_founded)
                        
                    # Update Employees
                    corr_emp = res.get("corrected_employees")
                    if corr_emp:
                        if "_crawl" not in company_dict:
                            company_dict["_crawl"] = {}
                        company_dict["_crawl"]["employee_count"] = str(corr_emp)
                else:
                    company_dict["is_relevant"] = False
                    
                company_dict["audit_completed"] = True
                company_dict["audit_timestamp"] = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
                
                status_str = "KEPT" if is_relevant else "REJECTED"
                print(f"  [{idx}/{total}] {name} ({domain}) -> {status_str} | Score={relevance_score} | Role={company_dict['role']} | HQ={res.get('corrected_hq') or '—'}")
                return
            except Exception as e:
                print(f"  [{idx}/{total}] Error auditing {name} (attempt {attempt+1}/3): {e}")
                await asyncio.sleep(2)
        
        # If all attempts fail, leave unchanged but print warning
        print(f"  [{idx}/{total}] Failed all attempts to audit {name}. Leaving as-is.")

async def audit_market_file(json_path: Path):
    print(f"\n========================================================")
    print(f"AUDITING MARKET: {json_path.parent.name}")
    print(f"File: {json_path.name}")
    print(f"========================================================")
    
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            result = json.load(f)
    except Exception as e:
        print(f"Failed to load JSON {json_path}: {e}")
        return
        
    market_name = get_market_name(result, json_path.parent.name.replace("_", " ").title())
    print(f"Detected Market Name: {market_name}")
    
    companies = result.get("relevant_companies") or []
    total = len(companies)
    if total == 0:
        print("No companies found in relevant_companies. Skipping.")
        return
        
    print(f"Found {total} companies to audit.")
    
    # Process in parallel using our semaphore and asyncio.gather
    tasks = []
    for i, company in enumerate(companies):
        tasks.append(audit_company(company, market_name, i + 1, total))
        
    await asyncio.gather(*tasks)
    
    # Save the updated JSON file
    try:
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(result, f, indent=2, default=str)
        print(f"Saved updated master JSON file.")
    except Exception as e:
        print(f"Failed to save JSON {json_path}: {e}")
        return
        
    # Regenerate exports
    try:
        print("Regenerating Excel export...")
        xlsx_path = json_path.with_suffix(".xlsx")
        export_fmcg_xlsx(result, xlsx_path)
        print(f"Saved Excel: {xlsx_path.resolve()}")
        
        print("Regenerating CSV export...")
        csv_path = json_path.with_suffix(".csv")
        save_pipeline_csv(result, str(csv_path))
        print(f"Saved CSV: {csv_path.resolve()}")
    except Exception as e:
        print(f"Failed to export CSV/XLSX: {e}")

async def main():
    start_time = time.perf_counter()
    
    # Find all output directories containing a global JSON file
    json_files = glob.glob("output/*/global_*.json")
    if not json_files:
        print("No global JSON files found under output/*/")
        return
        
    print(f"Found {len(json_files)} market file(s) to audit: {[Path(p).name for p in json_files]}")
    
    for json_path in json_files:
        await audit_market_file(Path(json_path))
        
    elapsed = time.perf_counter() - start_time
    print(f"\n========================================================")
    print(f"ALL AUDITS COMPLETED in {elapsed/60.0:.2f} minutes.")
    print(f"========================================================")

if __name__ == "__main__":
    if sys.platform == 'win32':
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    asyncio.run(main())
