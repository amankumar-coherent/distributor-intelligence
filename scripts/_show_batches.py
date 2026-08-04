import sqlite3
from pathlib import Path

p = Path("data/batch.db")
print("exists", p.exists())
if not p.exists():
    raise SystemExit
c = sqlite3.connect(p)
tables = [r[0] for r in c.execute("select name from sqlite_master where type='table'").fetchall()]
print("tables", tables)
for t in tables:
    cols = [r[1] for r in c.execute(f"pragma table_info({t})").fetchall()]
    print(t, cols)
    if "batch_id" in cols:
        print(" batch_ids", c.execute(f"select batch_id, count(*) from {t} group by batch_id").fetchall())
        qcols = set(cols)
        if {"query", "status"}.issubset(qcols):
            extra = ", step" if "step" in qcols else ""
            rows = c.execute(
                f"select batch_id, query, status{extra} from {t} order by rowid desc limit 30"
            ).fetchall()
            for r in rows:
                print(" ", r)
