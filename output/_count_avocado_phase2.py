import json
import re
from collections import Counter
from pathlib import Path

p = Path(r"D:\vendor-intel\output\phase2\phase2_companies_avocado_oil_global.json")
d = json.loads(p.read_text(encoding="utf-8"))
companies = d.get("companies") or []
print("phase2_count", len(companies))
print("company_count_field", d.get("company_count"))
stats = d.get("discovery_stats") or {}
for k, v in stats.items():
    if isinstance(v, (int, float, str, bool)):
        print(f"stat {k}: {v}")

secs = Counter(str(c.get("section") or "") for c in companies)
print("\nsections:")
for k, v in secs.most_common(20):
    print(f"  {v:4d}  {k!r}")

# Heuristic proxy (NOT classifier): name/domain signals
MFG = re.compile(
    r"\b(manufacturer|producer|mill|press|oil\s*co\.?|oilco|brand|farm|grower|"
    r"refiner|extraction|pressed|bottler)\b",
    re.I,
)
DIST = re.compile(
    r"\b(distributor|distribution|wholesale|wholesaler|importer|import|"
    r"trading|trader|foodservice|food\s*service|channel|reseller|"
    r"ingredients?\s*(distributor|supply)|supply\s*co)\b",
    re.I,
)
RETAIL = re.compile(
    r"\b(market|supermarket|grocery|hypermarket|retail|store|shop)\b",
    re.I,
)
JUNK = re.compile(
    r"(list of|top \d+|market size|report|benefits|characteristics|"
    r"what is|buy avocado|amazon\.|wikipedia|linkedin\.com/pulse)",
    re.I,
)

buckets = Counter()
for c in companies:
    name = str(c.get("name") or c.get("company_name") or "")
    dom = str(c.get("domain") or "")
    blob = f"{name} {dom} {c.get('section') or ''}"
    if JUNK.search(blob):
        buckets["junk_name"] += 1
    elif DIST.search(blob):
        buckets["name_looks_distributor"] += 1
    elif MFG.search(blob):
        buckets["name_looks_manufacturer"] += 1
    elif RETAIL.search(blob):
        buckets["name_looks_retail"] += 1
    else:
        buckets["unclear_needs_classify"] += 1

print("\nheuristic_name_proxy (not final classifier):")
for k, v in buckets.most_common():
    print(f"  {v:4d}  {k}")

print("\nsample unclear:")
n = 0
for c in companies:
    name = str(c.get("name") or "")
    blob = f"{name} {c.get('domain') or ''} {c.get('section') or ''}"
    if not (JUNK.search(blob) or DIST.search(blob) or MFG.search(blob) or RETAIL.search(blob)):
        print(" -", name, "|", c.get("domain"), "|", c.get("section"))
        n += 1
        if n >= 15:
            break
