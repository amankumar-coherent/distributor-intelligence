#!/usr/bin/env python3
import json
import csv
import io
import os
import shutil
from pathlib import Path

ROOT = Path(__file__).resolve().parent

MARKETS = [
    "avocado_oil",
    "technical_foam",
    "spectral_sensor",
    "gel_ice_pack",
    "cyber_physical_systems_security",
    "high_energy_laser",
]

def clean_domain(website: str) -> str:
    if not website:
        return ""
    d = website.strip().lower()
    for prefix in ["https://", "http://", "www."]:
        if d.startswith(prefix):
            d = d[len(prefix):]
    return d.split("/")[0]

def load_kept_companies_from_csv(market: str) -> list[dict]:
    csv_path = ROOT / f"output/{market}/global_{market}_market_global.csv"
    if not csv_path.exists():
        print(f"  [WARN] CSV not found: {csv_path}")
        return []
    with open(csv_path, encoding="utf-8", errors="replace") as f:
        lines = f.readlines()
    if not lines:
        return []
    header = lines[0]
    data_lines = [l for l in lines[1:] if not l.strip().startswith("===")]
    reader = csv.DictReader(io.StringIO(header + "".join(data_lines)))
    
    kept_companies = []
    for r in reader:
        company = r.get("Company")
        if company:
            kept_companies.append({
                "name": str(company).strip(),
                "website": str(r.get("Website", "")).strip(),
                "role": str(r.get("Role", "")).strip(),
                "summary": str(r.get("Summary", "")).strip(),
                "hq": str(r.get("Headquarters", "")).strip(),
                "founded": str(r.get("Founded", "")).strip(),
                "employees": str(r.get("Employees", "")).strip(),
                "turnover": str(r.get("Turnover / Scale", "")).strip(),
            })
    return kept_companies

def map_new_company_to_master(new_c: dict, audit_entry: dict = None) -> dict:
    company_name = new_c.get("Company") or (audit_entry.get("company") if audit_entry else "")
    website = new_c.get("Website") or (audit_entry.get("website") if audit_entry else "")
    domain = clean_domain(website)
    
    score = 5
    reasoning = ""
    if audit_entry:
        reasoning = audit_entry.get("reason", "")
        action = audit_entry.get("action", "keep")
        score = 5 if action == "keep" else (3 if action == "review" else 1)
        
    master_c = {
        "company": company_name,
        "domain": domain,
        "is_relevant": True,
        "role": new_c.get("Business Type") or (audit_entry.get("role") if audit_entry else "Manufacturer"),
        "value_chain_section": new_c.get("Specialty Focus") or "Other",
        "summary": new_c.get("Summary") or (audit_entry.get("what_they_do") if audit_entry else ""),
        "confidence": 1.0,
        "website": website,
        "is_seed": False,
        "company_summary": new_c.get("Summary") or "",
        "polished_summary": new_c.get("Summary") or "",
        "key_products": new_c.get("Core Categories") or "",
        "audit_completed": True,
        "audit_reasoning": reasoning,
        "audit_score": score,
        "audit_timestamp": audit_entry.get("audit_timestamp", "") if audit_entry else "",
        "_crawl": {
            "founded_year": new_c.get("Founded"),
            "headquarters": new_c.get("Headquarters"),
            "employee_count": new_c.get("Employees"),
            "revenue": new_c.get("Turnover / Scale"),
            "email": new_c.get("Email"),
            "phone": new_c.get("Phone / WhatsApp"),
            "linkedin": new_c.get("LinkedIn"),
            "products": new_c.get("Core Categories"),
        },
        "_presentation_enrich": {
            "ownership": new_c.get("Ownership"),
            "ecommerce": new_c.get("E-commerce"),
            "foodservice": new_c.get("Foodservice / HoReCa"),
            "retail_chains": new_c.get("Retail Chains"),
            "specialty_grocers": new_c.get("Specialty Grocers"),
            "specialty_focus": new_c.get("Specialty Focus"),
        }
    }
    return master_c

def main():
    print("=" * 80)
    print("SYNCHRONIZING JSON FILES WITH CURATED EXCEL/CSV DATA")
    print("=" * 80)
    
    for m in MARKETS:
        print(f"\nProcessing Market: {m.upper()}")
        
        json_path = ROOT / f"output/{m}/global_{m}_market_global.json"
        if not json_path.exists():
            print(f"  [ERROR] Master JSON not found: {json_path}")
            continue
            
        # 1. Load kept companies from CSV
        csv_companies = load_kept_companies_from_csv(m)
        if not csv_companies:
            print(f"  [WARN] No companies loaded from CSV. Skipping.")
            continue
        print(f"  Curated CSV company count: {len(csv_companies)}")
        
        # Create sets / lookups
        csv_names_lower = {c["name"].lower() for c in csv_companies}
        csv_lookup = {c["name"].lower(): c for c in csv_companies}
        
        # 2. Load audit cache if exists
        cache_path = ROOT / f"output/{m}/audit_cache.json"
        audit_lookup = {}
        if cache_path.exists():
            try:
                with open(cache_path, "r", encoding="utf-8") as f:
                    audit_list = json.load(f)
                audit_lookup = {a["company"].lower(): a for a in audit_list}
                print(f"  Loaded {len(audit_lookup)} audit cache entries")
            except Exception as e:
                print(f"  [WARN] Failed to load audit cache: {e}")
                
        # 3. Load discovery new companies if exists
        discovery_path = ROOT / f"output/{m}/discovery_new_companies.json"
        discovery_lookup = {}
        if discovery_path.exists():
            try:
                with open(discovery_path, "r", encoding="utf-8") as f:
                    discovery_list = json.load(f)
                discovery_lookup = {d["Company"].lower(): d for d in discovery_list}
                print(f"  Loaded {len(discovery_lookup)} discovery new company entries")
            except Exception as e:
                print(f"  [WARN] Failed to load discovery new companies: {e}")
                
        # 4. Load master JSON
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                master_data = json.load(f)
        except Exception as e:
            print(f"  [ERROR] Failed to load master JSON: {e}")
            continue
            
        old_relevant = master_data.get("relevant_companies") or []
        old_relevant_lookup = {c["company"].lower(): c for c in old_relevant}
        print(f"  Original JSON company count: {len(old_relevant)}")
        
        # 5. Build new relevant_companies list
        new_relevant = []
        for csv_c in csv_companies:
            name_lower = csv_c["name"].lower()
            audit_entry = audit_lookup.get(name_lower)
            
            # Case A: Company is already in the master JSON
            if name_lower in old_relevant_lookup:
                master_c = old_relevant_lookup[name_lower]
                master_c["is_relevant"] = True
                master_c["audit_completed"] = True
                if audit_entry:
                    master_c["audit_reasoning"] = audit_entry.get("reason", master_c.get("audit_reasoning"))
                    master_c["audit_score"] = 5 if audit_entry.get("action") == "keep" else (3 if audit_entry.get("action") == "review" else 1)
                    if "audit_timestamp" in audit_entry:
                        master_c["audit_timestamp"] = audit_entry["audit_timestamp"]
                new_relevant.append(master_c)
                
            # Case B: Company was newly discovered
            elif name_lower in discovery_lookup:
                new_c = discovery_lookup[name_lower]
                master_c = map_new_company_to_master(new_c, audit_entry)
                new_relevant.append(master_c)
                
            # Case C: Fallback, create minimal representation from CSV fields
            else:
                print(f"    - Creating fallback object for: {csv_c['name']}")
                fallback_c = {
                    "company": csv_c["name"],
                    "domain": clean_domain(csv_c["website"]),
                    "is_relevant": True,
                    "role": csv_c["role"] or "Manufacturer",
                    "value_chain_section": "Other",
                    "summary": csv_c["summary"],
                    "confidence": 1.0,
                    "website": csv_c["website"],
                    "is_seed": False,
                    "company_summary": csv_c["summary"],
                    "polished_summary": csv_c["summary"],
                    "key_products": "",
                    "audit_completed": True,
                    "_crawl": {
                        "founded_year": csv_c["founded"],
                        "headquarters": csv_c["hq"],
                        "employee_count": csv_c["employees"],
                        "revenue": csv_c["turnover"],
                    }
                }
                if audit_entry:
                    fallback_c["audit_reasoning"] = audit_entry.get("reason")
                    fallback_c["audit_score"] = 5 if audit_entry.get("action") == "keep" else 3
                new_relevant.append(fallback_c)
                
        print(f"  New JSON company count: {len(new_relevant)}")
        
        # Sanity check
        if len(new_relevant) != len(csv_companies):
            print(f"  [ERROR] Count mismatch! Expected {len(csv_companies)}, got {len(new_relevant)}")
            continue
            
        # 6. Update master JSON keys and values
        master_data["relevant_companies"] = new_relevant
        master_data["enriched_count"] = len(new_relevant)
        master_data["classified_count"] = len(new_relevant)
        master_data["after_entity_gate"] = len(new_relevant)
        if "phase2_company_count" in master_data:
            master_data["phase2_company_count"] = len(new_relevant)
            
        # Filter all_classified if it exists
        if "all_classified" in master_data:
            old_classified = master_data.get("all_classified") or []
            old_classified_lookup = {c["company"].lower(): c for c in old_classified}
            new_classified = []
            for csv_c in csv_companies:
                name_lower = csv_c["name"].lower()
                if name_lower in old_classified_lookup:
                    new_classified.append(old_classified_lookup[name_lower])
                else:
                    # Find in new_relevant
                    match_c = next((c for c in new_relevant if c["company"].lower() == name_lower), None)
                    if match_c:
                        new_classified.append(match_c)
            master_data["all_classified"] = new_classified
            
        # 7. Backup the original JSON file
        backup_path = json_path.with_suffix(".json.bak")
        try:
            shutil.copy2(json_path, backup_path)
            print(f"  Backup created: {backup_path.name}")
        except Exception as e:
            print(f"  [WARN] Could not create backup: {e}")
            
        # 8. Write updated JSON back
        try:
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(master_data, f, indent=2, ensure_ascii=False)
            print(f"  [SUCCESS] Updated master JSON file successfully.")
        except Exception as e:
            print(f"  [ERROR] Failed to save updated JSON: {e}")

if __name__ == "__main__":
    main()
