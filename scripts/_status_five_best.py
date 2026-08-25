# -*- coding: utf-8 -*-
import json
from pathlib import Path

ROOT = Path(r"D:\vendor-intel")

def summarize(folder: Path):
    print("\n==", folder.name, "==")
    if not folder.exists():
        print("MISSING")
        return
    audit = folder / "chatgpt_expand_batch_all_audit.json"
    ckpt = folder / "chatgpt_checkpoint_batch_all.json"
    final = list(folder.glob("*FINAL*.xlsx"))
    final = [f for f in final if "pre_" not in f.name and "BEFORE" not in f.name and "CORRUPT" not in f.name]
    if audit.exists():
        d = json.loads(audit.read_text(encoding="utf-8"))
        print("audit:", {k: d.get(k) for k in d if k in (
            "status", "step", "phase", "done", "total", "filled", "companies",
            "kept", "rejected", "query", "batch_id", "rows", "final_rows",
            "message", "current_company", "fill_done", "fill_total"
        ) or "row" in k.lower() or "fill" in k.lower() or "step" in k.lower()})
        print("audit_keys:", list(d.keys())[:30])
    if ckpt.exists():
        d = json.loads(ckpt.read_text(encoding="utf-8"))
        print("ckpt_keys:", list(d.keys())[:40])
        for k in ("step", "phase", "status", "current_step", "filled_part", "verified", "candidates",
                  "discover", "fill_index", "fill_total", "done_steps", "progress"):
            if k in d:
                v = d[k]
                if isinstance(v, list):
                    print(f"  {k}: list[{len(v)}]")
                elif isinstance(v, dict):
                    print(f"  {k}: dict keys={list(v.keys())[:12]}")
                else:
                    print(f"  {k}: {v}")
        # common nested
        data = d.get("data") if isinstance(d.get("data"), dict) else d
        for k in ("filled_part", "candidates", "verified_rows", "kept"):
            if isinstance(data.get(k), list):
                print(f"  data.{k}: {len(data[k])}")
        steps = data.get("steps_done") or data.get("completed_steps") or d.get("steps_done")
        if steps:
            print("  steps_done:", steps)
    if final:
        for f in sorted(final, key=lambda x: x.stat().st_mtime, reverse=True)[:3]:
            print("xlsx:", f.name, "mtime", f.stat().st_mtime)

folders = [
    ROOT / "output/web_expand/global_flexible_packaging_market_global",
    ROOT / "output/web_expand/global_wearable_medical_devices_market_global",
    ROOT / "output/web_expand/global_semiconductor_market_global",
    ROOT / "output/web_expand/global_liquefied_natural_gas_market_global",
    ROOT / "output/web_expand/global_glp_1_receptor_agonist_market_global",
    ROOT / "output/web_expand/global_glp1_receptor_agonist_market_global",
]
# discover lng/glp
for p in (ROOT / "output/web_expand").glob("*lng*"):
    folders.append(p)
for p in (ROOT / "output/web_expand").glob("*glp*"):
    folders.append(p)
for p in (ROOT / "output/web_expand").glob("*liquefied*"):
    folders.append(p)
for p in (ROOT / "output/web_expand").glob("*agonist*"):
    folders.append(p)

seen = set()
for f in folders:
    if f in seen:
        continue
    seen.add(f)
    summarize(f)

# flexible add distributor ckpt
fp = ROOT / "output/web_expand/global_flexible_packaging_market_global/add_distributor_batch_ckpt.json"
if fp.exists():
    d = json.loads(fp.read_text(encoding="utf-8"))
    print("\nflexible add_distributor ckpt keys", list(d.keys())[:20])
    print(json.dumps({k: (len(v) if isinstance(v, list) else v) for k, v in d.items() if k in d}, indent=2)[:1500])
