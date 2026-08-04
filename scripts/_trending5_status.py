import json
import sqlite3
from pathlib import Path

c = sqlite3.connect("data/batch.db")
rows = c.execute(
    "select query, status, started_at, finished_at, companies_exported, last_error, market_slug "
    "from markets where batch_id=? order by id",
    ("trending5-expand-900",),
).fetchall()
print("=== trending5-expand-900 ===")
for r in rows:
    print("---")
    print("query:", r[0])
    print("status:", r[1], "| exported:", r[4])
    print("slug:", r[6])
    print("started:", r[2])
    print("finished:", r[3])
    print("err:", (r[5] or "")[:300])

base = Path("output/web_expand")
slugs = [
    "global_antibody_drug_conjugates_market_global",
    "global_edge_data_center_market_global",
    "global_renewable_energy_technologies_market_global",
    "global_semiconductor_foundry_market_global",
    "global_oncology_radiopharmaceuticals_market_global",
]
print("\n=== checkpoints ===")
for slug in slugs:
    p = base / slug / "chatgpt_checkpoint_batch_all.json"
    print(slug, "exists=", p.exists())
    if p.exists():
        d = json.loads(p.read_text(encoding="utf-8"))
        prog = d.get("progress") or {}
        print(
            "  step=", d.get("step"),
            "sub=", d.get("substep"),
            "status=", d.get("status"),
            "fill=", prog.get("fill_count"),
            "backend=", prog.get("fill_backend"),
        )
