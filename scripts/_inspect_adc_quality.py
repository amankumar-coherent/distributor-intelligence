import json
from collections import Counter
from pathlib import Path

p = Path("output/web_expand/global_antibody_drug_conjugates_market_global/chatgpt_checkpoint_batch_all.json")
d = json.loads(p.read_text(encoding="utf-8"))
rows = (d.get("data") or {}).get("filled_all") or []
print("n", len(rows))
print("Specialty sample:")
for r in rows[:20]:
    print("-", r.get("Company"), "|", (r.get("Specialty Focus") or "")[:80], "|", (r.get("Summary") or "")[:60])

# OEM / known manufacturer names in biotech reagents
OEM_HINTS = (
    "thermo fisher", "sigma-aldrich", "merck kgaa", "abcam", "bio-techne",
    "miltenyi", "genscript", "cell signaling", "santa cruz", "peprotech",
    "cayman chemical", "biolegend", "fujifilm wako", "roche", "novartis",
    "astrazeneca", "pfizer", "gilead", "seagen", "daiichi", "immunogen",
    "adc therapeutics", "genentech", "amgen", "bms", "bristol", "eli lilly",
    "sanofi", "takeda", "abbvie", "johnson", "janssen", "glaxosmithkline", "gsk",
)
hits = []
for r in rows:
    n = (r.get("Company") or "").lower()
    for h in OEM_HINTS:
        if h in n:
            hits.append(r.get("Company"))
            break
print("\nlikely OEM/manufacturer name hits", len(hits))
print(hits[:40])

# HQ fill rate
NPD = {"", "n/a", "na", "none", "-", "not publicly disclosed", "unknown"}
print("HQ empty", sum(1 for r in rows if str(r.get("Headquarters") or "").strip().lower() in NPD))
print("Website empty", sum(1 for r in rows if str(r.get("Website") or "").strip().lower() in NPD))
print("Summary empty", sum(1 for r in rows if str(r.get("Summary") or "").strip().lower() in NPD))
print("HQ sample", Counter((r.get("Headquarters") or "")[:40] for r in rows).most_common(8))
