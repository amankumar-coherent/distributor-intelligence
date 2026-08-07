import json
from pathlib import Path

rows = [
    json.loads(l)
    for l in Path("scripts/_weak129_fix_progress.jsonl").read_text(encoding="utf-8").splitlines()
    if l.strip()
]
npd = [r for r in rows if r.get("action") == "npd"]
todo = []
for r in npd:
    fpath = r.get("file") or r.get("filepath") or ""
    todo.append(
        {
            "company": r["company"],
            "filepath": fpath,
            "file": Path(fpath).name if fpath else "",
            "old": r.get("old") or {},
            "prev_reason": (r.get("parsed") or {}).get("reason") or "",
        }
    )
Path("scripts/_npd31_todo.json").write_text(
    json.dumps(todo, indent=2, ensure_ascii=False), encoding="utf-8"
)
print("count", len(todo))
for t in todo:
    print(f"{t['file']} | {t['company']}")
