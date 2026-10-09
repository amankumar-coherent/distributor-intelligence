#!/usr/bin/env python3
"""Merge the parallel workers' shard checkpoints into one checkpoint + Excel.

Each worker writes its own ckpt_N.json so the processes never contend for a
single file. This folds them back into ceo_chro_india.json (the file every
other script in this flow reads) and rewrites the workbook.

Merge rule: a filled field never loses to an empty one. If two shards hold
the same company, whichever has a name wins, so re-running a worker can
only add.
"""
import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

OUT = ROOT / "output" / "ai_nodata_core3"
FIELDS = ["CEO", "AI HEAD", "CTO"]


def merge_entry(old: dict, new: dict) -> dict:
    out = dict(old or {})
    out["website"] = (old or {}).get("website") or (new or {}).get("website", "")
    for field in FIELDS:
        a = (old or {}).get(field) or {}
        b = (new or {}).get(field) or {}
        out[field] = a if a.get("name") else (b if b.get("name") else (a or b))
    return out


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--companies-file", required=True)
    parser.add_argument("--merged", default=str(OUT / "ai_nodata_core3.json"))
    parser.add_argument("--output", default=str(OUT / "ai_nodata_core3.xlsx"))
    args = parser.parse_args()

    from run_core3 import write_excel

    merged_path = Path(args.merged)
    data = json.loads(merged_path.read_text(encoding="utf-8")) if merged_path.exists() else {}
    before = sum(1 for v in data.values() if any((v.get(f) or {}).get("name") for f in FIELDS))

    shard_dir = OUT / "shards"
    shards = sorted(shard_dir.glob("ckpt_*.json")) if shard_dir.is_dir() else []
    for shard in shards:
        try:
            part = json.loads(shard.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            print(f"  skipped unreadable {shard.name}")
            continue
        for key, entry in part.items():
            data[key] = merge_entry(data.get(key), entry)
        print(f"  merged {shard.name}: {len(part)} companies")

    merged_path.parent.mkdir(parents=True, exist_ok=True)
    merged_path.write_text(json.dumps(data, indent=1, ensure_ascii=False), encoding="utf-8")

    names = [l.strip() for l in Path(args.companies_file).read_text(encoding="utf-8").splitlines() if l.strip()]
    seen, companies = set(), []
    for n in names:
        if n.lower() in seen:
            continue
        seen.add(n.lower())
        companies.append(n)

    write_excel(data, companies, Path(args.output))

    after = sum(1 for v in data.values() if any((v.get(f) or {}).get("name") for f in FIELDS))
    ceo = sum(1 for v in data.values() if (v.get("CEO") or {}).get("name"))
    chro = sum(1 for v in data.values() if (v.get("CHRO") or {}).get("name"))
    web = sum(1 for v in data.values() if v.get("website"))
    print(f"\nshards merged : {len(shards)}")
    print(f"with a name   : {before} -> {after}")
    print(f"  CEO {ceo} | CHRO {chro} | websites {web}")
    print(f"Excel: {args.output}")


if __name__ == "__main__":
    main()
