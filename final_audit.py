#!/usr/bin/env python3
"""
Final Audit Script - Verifies each company against web search results
to ensure they actually operate in the claimed market.
Uses OpenAI to evaluate relevance from search snippets.
"""
import asyncio
import json
import os
import sys
import time
import csv
import io
import re
from pathlib import Path
from typing import Any

from openai import AsyncOpenAI

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

# Load .env
env_path = ROOT / ".env"
if env_path.exists():
    with open(env_path, encoding="utf-8") as _ef:
        for _line in _ef:
            _line = _line.strip()
            if _line and not _line.startswith("#") and "=" in _line:
                _k, _v = _line.split("=", 1)
                os.environ.setdefault(_k.strip(), _v.strip())

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
client = AsyncOpenAI(api_key=OPENAI_API_KEY)

MARKETS = [
    "avocado_oil",
    "technical_foam",
    "spectral_sensor",
    "gel_ice_pack",
    "cyber_physical_systems_security",
    "high_energy_laser",
]

MARKET_DESCRIPTIONS = {
    "avocado_oil": "avocado oil production, processing, bottling, or distribution as a food or cosmetic ingredient",
    "technical_foam": "technical/industrial foam manufacturing, polyurethane foam, acoustic foam, thermal insulation foam, memory foam for industrial uses",
    "spectral_sensor": "spectral sensors, hyperspectral imaging, NIR sensors, spectrometers, multispectral cameras for industrial/scientific use",
    "gel_ice_pack": "gel ice packs, cold packs, reusable ice packs, cold chain packaging, temperature-controlled packaging",
    "cyber_physical_systems_security": "cybersecurity for cyber-physical systems (CPS), OT/ICS security, industrial control system security, SCADA security, IoT security",
    "high_energy_laser": "high energy laser systems, directed energy weapons, laser cutting/welding industrial systems, fiber lasers, solid-state lasers",
}


def load_companies() -> dict[str, list[dict]]:
    """Load companies from CSV files."""
    all_companies = {}
    for m in MARKETS:
        csv_path = ROOT / f"output/{m}/global_{m}_market_global.csv"
        if not csv_path.exists():
            print(f"[WARN] {csv_path} not found")
            continue
        with open(csv_path, encoding="utf-8", errors="replace") as f:
            lines = f.readlines()
        header = lines[0]
        data_lines = [l for l in lines[1:] if not l.strip().startswith("===")]
        reader = csv.DictReader(io.StringIO(header + "".join(data_lines)))
        rows = list(reader)
        companies = []
        for r in rows:
            companies.append({
                "name": r.get("Company", "").strip(),
                "website": r.get("Website", "").strip(),
                "role": r.get("Role", "").strip(),
                "functionality": r.get("Functionality", "").strip(),
                "key_products": r.get("Key_Products", "").strip(),
                "confidence": r.get("Confidence", "").strip(),
                "summary": r.get("Summary", "").strip()[:300],
                "industry": r.get("Industry", "").strip(),
            })
        all_companies[m] = companies
        print(f"  Loaded {m}: {len(companies)} companies")
    return all_companies


async def audit_company_batch(market: str, companies: list[dict], semaphore: asyncio.Semaphore) -> list[dict]:
    """Audit a batch of companies for a market using LLM."""
    market_desc = MARKET_DESCRIPTIONS.get(market, market)
    results = []
    
    async def audit_one(company: dict) -> dict:
        async with semaphore:
            name = company["name"]
            website = company["website"]
            functionality = company["functionality"]
            key_products = company["key_products"]
            summary = company["summary"]
            role = company["role"]
            
            # Build company profile text
            profile = f"""Company: {name}
Website: {website}
Role: {role}
Functionality: {functionality}
Key Products/Services: {key_products}
Summary: {summary}"""
            
            prompt = f"""You are a market research auditor. Your task is to verify whether a company truly operates in a specific market.

MARKET: {market.replace('_', ' ').upper()}
MARKET DESCRIPTION: {market_desc}

COMPANY PROFILE:
{profile}

Based ONLY on the information provided (no hallucination), evaluate:
1. Is this company genuinely active in the stated market? (yes/partial/no/unclear)
2. What does this company ACTUALLY do based on its name, website, products?
3. Is there any red flag? (e.g., wrong industry, generic name, no clear products)
4. Confidence in their market relevance: high/medium/low
5. Recommended action: keep/review/remove

RULES:
- If company name or website is a generic domain (like "lasers.com", "industrialcybersec.com") without clear identity, flag as unclear
- If products/functionality don't match market, flag as no/review  
- If company is clearly in the market, mark as yes/keep
- Never invent information - only use what's in the profile
- Large well-known companies in adjacent markets (e.g., IBM in cybersecurity) = yes/keep

Respond in JSON format:
{{
  "company": "{name}",
  "market_relevant": "yes|partial|no|unclear",
  "what_they_do": "1-2 sentence description of what company actually does",
  "red_flags": "none or description of issue",
  "market_confidence": "high|medium|low",
  "action": "keep|review|remove",
  "reason": "brief justification"
}}"""

            for attempt in range(3):
                try:
                    resp = await client.chat.completions.create(
                        model="gpt-4o-mini",
                        messages=[{"role": "user", "content": prompt}],
                        temperature=0,
                        max_tokens=400,
                        response_format={"type": "json_object"},
                    )
                    result = json.loads(resp.choices[0].message.content)
                    result["original_confidence"] = company.get("confidence", "")
                    result["website"] = website
                    return result
                except Exception as e:
                    if attempt < 2:
                        await asyncio.sleep(2 ** attempt)
                    else:
                        return {
                            "company": name,
                            "market_relevant": "unclear",
                            "what_they_do": functionality,
                            "red_flags": f"API error: {e}",
                            "market_confidence": "low",
                            "action": "review",
                            "reason": "Could not audit",
                            "original_confidence": company.get("confidence", ""),
                            "website": website,
                        }
    
    tasks = [audit_one(c) for c in companies]
    results = await asyncio.gather(*tasks)
    return list(results)


async def main():
    print("=" * 70)
    print("FINAL COMPANY AUDIT - Verifying market relevance for all companies")
    print("=" * 70)
    
    all_companies = load_companies()
    
    # Check if we have a cached run
    audit_cache_path = ROOT / "audit_results_cache.json"
    audit_cache = {}
    if audit_cache_path.exists():
        with open(audit_cache_path, encoding="utf-8") as f:
            audit_cache = json.load(f)
        print(f"\n[INFO] Loaded cached results for {len(audit_cache)} markets")
    
    semaphore = asyncio.Semaphore(8)  # 8 parallel requests
    all_results = {}
    
    for market, companies in all_companies.items():
        if market in audit_cache:
            print(f"\n[SKIP] {market} already audited ({len(audit_cache[market])} results cached)")
            all_results[market] = audit_cache[market]
            continue
        
        print(f"\n[AUDIT] {market.upper()} - {len(companies)} companies...")
        start = time.time()
        results = await audit_company_batch(market, companies, semaphore)
        elapsed = time.time() - start
        print(f"  Done in {elapsed:.1f}s")
        
        all_results[market] = results
        audit_cache[market] = results
        
        # Save cache after each market
        with open(audit_cache_path, "w", encoding="utf-8") as f:
            json.dump(audit_cache, f, indent=2, ensure_ascii=False)
        print(f"  Saved cache for {market}")
    
    # Generate final report
    print("\n" + "=" * 70)
    print("GENERATING AUDIT REPORT")
    print("=" * 70)
    
    report_lines = []
    report_lines.append("# Final Market Audit Report\n")
    report_lines.append("*Audit verifies each company's actual business vs claimed market relevance*\n\n")
    
    grand_total = 0
    grand_keep = 0
    grand_review = 0
    grand_remove = 0
    
    suspicious_companies = []
    
    for market, results in all_results.items():
        market_desc = MARKET_DESCRIPTIONS.get(market, "")
        keep = [r for r in results if r.get("action") == "keep"]
        review = [r for r in results if r.get("action") == "review"]
        remove_list = [r for r in results if r.get("action") == "remove"]
        
        grand_total += len(results)
        grand_keep += len(keep)
        grand_review += len(review)
        grand_remove += len(remove_list)
        
        report_lines.append(f"## {market.replace('_', ' ').title()} Market\n")
        report_lines.append(f"**Market Focus:** {market_desc}\n\n")
        report_lines.append(f"| Status | Count |\n|--------|-------|\n")
        report_lines.append(f"| ✅ Keep | {len(keep)} |\n")
        report_lines.append(f"| ⚠️ Review | {len(review)} |\n")
        report_lines.append(f"| ❌ Remove | {len(remove_list)} |\n")
        report_lines.append(f"| **Total** | **{len(results)}** |\n\n")
        
        # All companies
        report_lines.append("### Company Audit Details\n\n")
        report_lines.append("| # | Company | What They Do | Market Relevant | Action | Red Flags |\n")
        report_lines.append("|---|---------|-------------|-----------------|--------|----------|\n")
        
        for i, r in enumerate(results, 1):
            name = r.get("company", "?")
            what = r.get("what_they_do", "")[:80]
            relevant = r.get("market_relevant", "?")
            action = r.get("action", "?")
            flags = r.get("red_flags", "none")[:60]
            
            # emoji for action
            emoji = {"keep": "✅", "review": "⚠️", "remove": "❌"}.get(action, "❓")
            rel_emoji = {"yes": "✅", "partial": "🟡", "no": "❌", "unclear": "❓"}.get(relevant, "❓")
            
            report_lines.append(f"| {i} | **{name}** | {what} | {rel_emoji} {relevant} | {emoji} {action} | {flags} |\n")
            
            if action in ("review", "remove"):
                suspicious_companies.append({
                    "market": market,
                    "company": name,
                    "website": r.get("website", ""),
                    "action": action,
                    "market_relevant": relevant,
                    "what_they_do": r.get("what_they_do", ""),
                    "red_flags": r.get("red_flags", ""),
                    "reason": r.get("reason", ""),
                })
        
        report_lines.append("\n")
        
        # Print console summary
        print(f"\n  {market.upper()}: {len(keep)} keep, {len(review)} review, {len(remove_list)} remove")
        if review:
            print(f"  [REVIEW]: {', '.join(r['company'] for r in review[:10])}")
        if remove_list:
            print(f"  [REMOVE]: {', '.join(r['company'] for r in remove_list[:10])}")
    
    # Summary section
    report_lines.insert(2, f"## Executive Summary\n\n")
    report_lines.insert(3, f"| Metric | Count |\n|--------|-------|\n")
    report_lines.insert(4, f"| Total Companies Audited | {grand_total} |\n")
    report_lines.insert(5, f"| ✅ Confirmed Relevant (Keep) | {grand_keep} |\n")
    report_lines.insert(6, f"| ⚠️ Needs Review | {grand_review} |\n")
    report_lines.insert(7, f"| ❌ Recommended Removal | {grand_remove} |\n\n")
    
    # Suspicious companies list
    if suspicious_companies:
        report_lines.append("## Companies Flagged for Review / Removal\n\n")
        report_lines.append("| Market | Company | Action | Issue |\n")
        report_lines.append("|--------|---------|--------|-------|\n")
        for c in suspicious_companies:
            market_label = c["market"].replace("_", " ").title()
            action_label = "[REVIEW]" if c["action"] == "review" else "[REMOVE]"
            report_lines.append(f"| {market_label} | **{c['company']}** | {action_label} {c['action']} | {c['red_flags'][:80]} |\n")
    
    # Save report
    report_path = ROOT / "output" / "final_audit_report.md"
    with open(report_path, "w", encoding="utf-8") as f:
        f.writelines(report_lines)
    
    print(f"\n\n{'='*70}")
    print(f"AUDIT COMPLETE!")
    print(f"Total Companies: {grand_total}")
    print(f"  [KEEP]:   {grand_keep} ({grand_keep/grand_total*100:.1f}%)")
    print(f"  [REVIEW]: {grand_review} ({grand_review/grand_total*100:.1f}%)")
    print(f"  [REMOVE]: {grand_remove} ({grand_remove/grand_total*100:.1f}%)")
    print(f"\nReport saved to: {report_path}")
    print(f"{'='*70}")
    
    # Also save suspicious companies as JSON for quick reference
    suspicious_path = ROOT / "output" / "flagged_companies.json"
    with open(suspicious_path, "w", encoding="utf-8") as f:
        json.dump(suspicious_companies, f, indent=2, ensure_ascii=False)
    print(f"Flagged companies saved to: {suspicious_path}")


if __name__ == "__main__":
    asyncio.run(main())
