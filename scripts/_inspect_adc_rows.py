import json
from collections import Counter
from pathlib import Path

p = Path("output/web_expand/global_antibody_drug_conjugates_market_global/chatgpt_checkpoint_batch_all.json")
d = json.loads(p.read_text(encoding="utf-8"))
data = d.get("data") or {}
for k in ["filled_part", "filled_all", "to_fill", "merged"]:
    v = data.get(k)
    print(k, len(v) if isinstance(v, list) else v)

rows = data.get("filled_all") or data.get("filled_part") or data.get("to_fill") or []
print("using", "filled_all" if data.get("filled_all") else "filled_part" if data.get("filled_part") else "to_fill", "n=", len(rows))

NPD = {"", "n/a", "na", "none", "-", "—", "not publicly disclosed", "unknown", "tbd"}

def empty(v):
    return str(v or "").strip().lower() in NPD

if rows:
    print("BT:", Counter(r.get("Business Type") or "" for r in rows).most_common(15))
    cols = list(rows[0].keys())
    print("cols", len(cols))
    for c in cols:
        e = sum(1 for r in rows if empty(r.get(c)))
        if e:
            print(f"  empty {c}: {e}/{len(rows)}")
    # sample manufacturers-ish
    mfr = [r.get("Company") for r in rows if "manufactur" in (r.get("Business Type") or "").lower()]
    print("mfr BT", len(mfr), mfr[:10])
    print("sample companies:", [r.get("Company") for r in rows[:15]])
