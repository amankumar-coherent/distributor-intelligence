import json
from pathlib import Path

cache = Path(r"D:\vendor-intel\data\search_cache")
zoya_rows = []
gai_bulgaria = []

for fp in cache.glob("*.json"):
    try:
        raw = fp.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        continue
    if "avocado oil" not in raw.lower() or "bulgaria" not in raw.lower():
        continue
    try:
        data = json.loads(raw)
    except Exception:
        continue
    q = str(data.get("query") or "")
    if "bulgaria" not in q.lower():
        continue
    for h in data.get("results") or []:
        if not isinstance(h, dict):
            continue
        blob = json.dumps(h, ensure_ascii=False).lower()
        backend = str(h.get("backend") or "")
        title = str(h.get("title") or "")
        snippet = str(h.get("snippet") or "")
        link = str(h.get("link") or "")
        row = {
            "q": q[:100],
            "backend": backend,
            "title": title[:120],
            "link": link[:120],
            "snippet": snippet.replace("\n", " ")[:200],
            "has_zoya": "zoya" in blob,
        }
        if row["has_zoya"]:
            zoya_rows.append(row)
        if "google_ai" in backend:
            gai_bulgaria.append(row)

print("queries with bulgaria+avocado in cache that have Zoya in a hit:", len(zoya_rows))
print("google_ai_scraper hits for bulgaria avocado queries:", len(gai_bulgaria))
print()
print("=== Any hit containing Zoya ===")
if not zoya_rows:
    print("NONE — Zoya was NOT stored in search cache for these queries")
else:
    for r in zoya_rows[:15]:
        print(f"backend={r['backend']}")
        print(f"  title={r['title']}")
        print(f"  link={r['link']}")
        print(f"  q={r['q']}")
        print()

print("=== Google AI scraper Bulgaria payloads (Zoya inside?) ===")
if not gai_bulgaria:
    print("No google_ai_scraper rows for Bulgaria avocado queries")
else:
    for r in gai_bulgaria[:6]:
        print(f"has_zoya={r['has_zoya']} title={r['title']}")
        print(f"  snip={r['snippet'][:180]}")
        print()
