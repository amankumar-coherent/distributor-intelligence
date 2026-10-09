"""Rough quality estimate on Phase2 avocado list (pre-classifier)."""
from __future__ import annotations

import json
import re
from collections import Counter
from pathlib import Path

p = Path(r"D:\vendor-intel\output\phase2\phase2_companies_avocado_oil_global.json")
companies = json.loads(p.read_text(encoding="utf-8")).get("companies") or []

MARKETPLACE = {
    "tridge.com", "tradewheel.com", "alibaba.com", "indiamart.com", "tradeindia.com",
    "thomasnet.com", "kompass.com", "go4worldbusiness.com", "getdistributors.com",
    "europages.com", "exportersindia.com", "made-in-china.com", "globalsources.com",
    "amazon.com", "ebay.com", "walmart.com", "cookeatshare.com", "wikipedia.org",
    "linkedin.com", "facebook.com", "youtube.com",
}
KNOWN_MFG = {
    "chosenfoods.com", "olivado.com", "latourangelle.com", "kevala.net",
    "primalkitchen.com", "bellavado.com", "calpure.com", "westfaliafruit.com",
    "groveavocadooil.co.nz", "avocadoreal.com", "connoils.com", "spectrumorganics.com",
    "mantova.com", "avopure.com",
}
KNOWN_DIST = {
    "sysco.com", "usfoods.com", "kehe.com", "unfi.com", "brenntag.com",
    "azelis.com", "univarsolutions.com", "bidfood.com", "imcdgroup.com",
}

MFG_NAME = re.compile(
    r"\b(oils?|press(?:ery|ed)?|mill|farm|grower|producer|brand|kitchen|"
    r"organics?|avocado\s*oil\s*(co|company|direct|essentials|international)?)\b",
    re.I,
)
DIST_NAME = re.compile(
    r"\b(distribut|wholesale|import|trading|foodservice|ingredients?|"
    r"supply|supplies|reseller|dealer)\b",
    re.I,
)
JUNK_NAME = re.compile(
    r"(market\s+size|report|benefits|characteristics|buyers?\s+guide|"
    r"list\s+of|top\s+\d+|importers\s*&\s*|persea|health\s+benefits)",
    re.I,
)


def host(dom: str) -> str:
    d = (dom or "").strip().lower()
    d = re.sub(r"^https?://", "", d)
    d = d.split("/")[0]
    if d.startswith("www."):
        d = d[4:]
    return d


buckets = Counter()
examples: dict[str, list[str]] = {k: [] for k in [
    "marketplace_directory", "known_manufacturer", "known_distributor",
    "junk_name", "name_suggests_mfg_brand", "name_suggests_channel", "unknown",
]}

for c in companies:
    name = str(c.get("name") or "").strip()
    dom = host(str(c.get("domain") or ""))
    if not name:
        buckets["empty"] += 1
        continue
    if JUNK_NAME.search(name):
        key = "junk_name"
    elif dom in MARKETPLACE or any(dom.endswith("." + m) for m in MARKETPLACE):
        key = "marketplace_directory"
    elif dom in KNOWN_MFG:
        key = "known_manufacturer"
    elif dom in KNOWN_DIST:
        key = "known_distributor"
    elif DIST_NAME.search(name):
        key = "name_suggests_channel"
    elif MFG_NAME.search(name) and "distribut" not in name.lower():
        key = "name_suggests_mfg_brand"
    else:
        key = "unknown"
    buckets[key] += 1
    if len(examples[key]) < 8:
        examples[key].append(f"{name} ({dom})")

print(f"TOTAL_PHASE2={len(companies)}")
print("\nRough pre-classify breakdown:")
for k, v in buckets.most_common():
    pct = 100.0 * v / max(len(companies), 1)
    print(f"  {v:4d}  ({pct:5.1f}%)  {k}")

print("\nExamples:")
for k, rows in examples.items():
    if not rows:
        continue
    print(f"\n[{k}]")
    for r in rows:
        print("  -", r)

# Honest estimate bands based on prior avocado channel runs in this project:
# raw ~600 -> classified relevant channel often ~10-25% after strict gate
likely_channelish = buckets["known_distributor"] + buckets["name_suggests_channel"]
mfgish = buckets["known_manufacturer"] + buckets["name_suggests_mfg_brand"]
junkish = buckets["marketplace_directory"] + buckets["junk_name"]
unknown = buckets["unknown"]
print("\nESTIMATE (until Phase4 classify finishes):")
print(f"  clearly_not_genuine_now ~= {junkish + mfgish} (marketplaces/junk + mfg-looking)")
print(f"  channel-looking_by_name ~= {likely_channelish}")
print(f"  unknown_need_enrich_classify ~= {unknown}")
print("  expected_genuine_distributors_after_classify ~= 60-150 of 568")
print("  (strict CHANNEL_PARTNER_MODE; prior similar avocado run ~67 export-relevant)")
