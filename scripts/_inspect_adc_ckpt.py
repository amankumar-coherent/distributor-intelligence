import json
from pathlib import Path

p = Path("output/web_expand/global_antibody_drug_conjugates_market_global/chatgpt_checkpoint_batch_all.json")
d = json.loads(p.read_text(encoding="utf-8"))
print("step", d.get("step"), "sub", d.get("substep"), "status", d.get("status"))
print("completed", d.get("completed"))
print("progress", d.get("progress"))
data = d.get("data") or {}
print("data keys", list(data.keys()))
for k in [
    "merged",
    "to_fill",
    "filled",
    "verified",
    "extracted",
    "raw",
    "discovered",
    "recalled",
    "rows",
    "final_rows",
]:
    v = data.get(k)
    if isinstance(v, list):
        print(k, "len", len(v))
        if v and isinstance(v[0], dict):
            print("  sample keys", list(v[0].keys())[:25])
            c = v[0]
            print(
                "  sample company=",
                c.get("Company") or c.get("company") or c.get("name"),
                "bt=",
                c.get("Business Type") or c.get("business_type") or c.get("role"),
            )
    elif v is not None:
        print(k, type(v).__name__, str(v)[:100])

out = Path("output/web_expand/global_antibody_drug_conjugates_market_global")
print("files:", sorted(x.name for x in out.iterdir()))
