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
from vendor_intel.config import Settings
from vendor_intel.clients.search_router import FreeSearchRouter
from vendor_intel.export.fmcg_landscape import export_fmcg_xlsx
from vendor_intel.pipeline.orchestrator import save_pipeline_csv

# Load settings to verify API config
cfg = get_settings()
if not cfg.openai_api_key:
    print("ERROR: OPENAI_API_KEY is not set in environment or .env file.")
    sys.exit(1)

client = AsyncOpenAI(api_key=cfg.openai_api_key)
model_name = cfg.model_extract or "gpt-4o-mini"

# Set settings object for SearchRouter
proj_settings = Settings.load()
router = FreeSearchRouter(proj_settings)

# Semaphore to control rate limits
SEM_LIMIT = 5
semaphore = asyncio.Semaphore(SEM_LIMIT)

def get_market_name(result: dict, default_val: str) -> str:
    ctx = result.get("query_context") or {}
    name = ctx.get("industry") or result.get("query")
    if name:
        return str(name).replace("Global ", "").replace(" Market", "").strip()
    return default_val

async def audit_company_via_search(company_dict: dict, market_name: str, idx: int, total: int, json_path: Path, lock: asyncio.Lock):
    name = company_dict.get("company") or company_dict.get("brand") or "Unknown"
    domain = company_dict.get("domain") or company_dict.get("website") or ""
    
    # Strip domain to get base domain for query fallback
    clean_domain = domain.removeprefix("https://").removeprefix("http://").removeprefix("www.").split("/")[0] if domain else ""

    # Check if we already have presentation data to keep it or refine it
    pe_old = company_dict.get("_presentation_enrich") or {}

    # Query to check what the company actually does
    query = f'"{name}" company products services overview'
    if clean_domain:
        query_alt = f'"{name}" site:{clean_domain} what we do'
    else:
        query_alt = f'"{name}" business profile overview'

    hits = []
    async with semaphore:
        try:
            # Try primary query
            hits = await router.search(query, discovery_mode=True)
            if not hits:
                await asyncio.sleep(1.0)
                hits = await router.search(query_alt, discovery_mode=True)
        except Exception as e:
            # Suppress connection errors
            pass

        # Prepare context text (from search snippets or crawl fallback)
        context_source = "web_search"
        context_text = ""
        
        if hits:
            lines = []
            for h in hits[:5]:
                lines.append(f"- {h.title}: {h.snippet} ({h.link})")
            context_text = "\n".join(lines)
        else:
            # Fallback to scraped website text if search fails
            _crawl = company_dict.get("_crawl") or {}
            context_text = _crawl.get("crawl_text") or ""
            if context_text:
                context_source = "crawl_text_fallback"
                if len(context_text) > 8000:
                    context_text = context_text[:8000] + " ... [truncated]"
            else:
                context_source = "general_knowledge_fallback"

        system_prompt = (
            "You are a critical market research auditor. Your job is to strictly verify if a company is a real participant in the target market.\n"
            "Analyze the provided business details and determine if they manufacture, distribute, or provide specialized products/services in this market.\n"
            "CRITICAL RULES:\n"
            "1. Base your decisions strictly on facts found in the provided search snippets/text. If the company is unrelated, you must reject it.\n"
            "2. Do NOT guess or hallucinate. If you lack direct evidence of their presence in this market, set 'is_relevant' to false.\n"
            "3. Reject directories, portals, event pages, unrelated consumers, or general software consultancies with no specific market offering.\n"
            "4. Return a raw JSON object matching the exact structure below."
        )

        user_prompt = f"""Target Market: {market_name}
Company Name: {name}
Website Domain: {domain}
Context Source: {context_source}

Provided Source Text:
{context_text if context_text else "[No text available, use strict factual knowledge only. If not 100% certain, mark is_relevant=false]"}

JSON Output Format (Strictly return ONLY this JSON object):
{{
  "is_relevant": true / false,
  "relevance_score": integer from 0 to 5,
  "role": "Manufacturer" / "Distributor" / "Supplier" / "Integrator" / "Service Provider" / "Unrelated",
  "reasoning": "Explain exactly what products/services they provide in the target market based on the text. If irrelevant, explain why.",
  "ownership": "Corporate ownership. Choose from: 'Private', 'Public (Exchange: [Ticker])', or 'Subsidiary of [Parent Name]'.",
  "ecommerce": "Specific platforms or e-store names (e.g. 'Amazon, Shopify Store') if they sell directly online in this market. If none/unknown, write 'No'.",
  "foodservice": "Specific foodservice clients/distributors (e.g. 'Sysco, US Foods, hotel chains') if they supply HoReCa. If none/unknown, write 'No'.",
  "retail_chains": "Specific supermarket/retail chains (e.g. 'Whole Foods, Costco, Kroger'). If none/unknown, write 'No'.",
  "specialty_grocers": "Specific specialty/organic grocers (e.g. 'Trader Joe's, Sprouts'). If none/unknown, write 'No'.",
  "specialty_focus": "The precise product or technical focus of the company in this market (e.g. 'OT/CPS cybersecurity systems', 'Organic extra virgin avocado oil'). Do NOT include standards like ISO 9001, HACCP, GMP, etc."
}}
"""

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
                res = json.loads(response.choices[0].message.content)
                
                # Check relevance
                relevance_score = res.get("relevance_score", 0)
                is_relevant = res.get("is_relevant", False)
                if relevance_score < 3:
                    is_relevant = False

                # Apply updates to company dict
                company_dict["is_relevant"] = is_relevant
                company_dict["role"] = res.get("role") or company_dict.get("role") or "Other"
                company_dict["confidence"] = round(relevance_score / 5.0, 2)
                company_dict["audit_reasoning"] = res.get("reasoning", "")
                company_dict["audit_score"] = relevance_score
                company_dict["final_search_audited"] = True
                company_dict["audit_source"] = context_source

                if is_relevant:
                    company_dict["_presentation_enrich"] = {
                        "ownership": res.get("ownership") or pe_old.get("ownership") or "Private",
                        "ecommerce": res.get("ecommerce") or pe_old.get("ecommerce") or "No",
                        "foodservice": res.get("foodservice") or pe_old.get("foodservice") or "No",
                        "retail_chains": res.get("retail_chains") or pe_old.get("retail_chains") or "No",
                        "specialty_grocers": res.get("specialty_grocers") or pe_old.get("specialty_grocers") or "No",
                        "specialty_focus": res.get("specialty_focus") or pe_old.get("specialty_focus") or ""
                    }
                    print(f"  [{idx}/{total}] Re-verified {name}: Relevant (Score={relevance_score}, Source={context_source})")
                else:
                    print(f"  [{idx}/{total}] Flagged IRRELEVANT {name}: (Score={relevance_score}, Source={context_source}) - Reason: {res.get('reasoning')}")

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

                # Polite delay to prevent DDGS rate limiting
                await asyncio.sleep(1.5)
                return
            except Exception as e:
                print(f"  [{idx}/{total}] Error auditing {name} (attempt {attempt+1}/2): {e}")
                await asyncio.sleep(1.5)

async def audit_market_file(json_path: Path):
    print(f"\n========================================================")
    print(f"FINAL AUDIT OF MARKET: {json_path.parent.name}")
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
    # Audit currently relevant ones to refine/verify
    todo_companies = [c for c in companies if c.get("is_relevant", True)]
    total = len(todo_companies)
    if total == 0:
        print("No relevant companies to audit. Skipping.")
        return

    print(f"Auditing {total} companies...")
    
    lock = asyncio.Lock()
    tasks = []
    for i, company in enumerate(todo_companies):
        tasks.append(audit_company_via_search(company, market_name, i + 1, total, json_path, lock))

    await asyncio.gather(*tasks)

    # Reload full updated data to run exports
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            updated_result = json.load(f)
        
        print("Regenerating Excel export...")
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

    print(f"Found {len(json_files)} market file(s) to audit: {[Path(p).name for p in json_files]}")

    for json_path in json_files:
        await audit_market_file(Path(json_path))

    elapsed = time.perf_counter() - start_time
    print(f"\n========================================================")
    print(f"FINAL AUDITING COMPLETED in {elapsed/60.0:.2f} minutes.")
    print(f"========================================================")

if __name__ == "__main__":
    if sys.platform == 'win32':
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    asyncio.run(main())
