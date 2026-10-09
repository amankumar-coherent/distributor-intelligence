import json
from pathlib import Path

cache = Path(r"D:\vendor-intel\data\search_cache")
needles = ["interdrink", "eterika", "evterpa", "divisha", "zoya", "budjak"]

rows = []
for fp in cache.glob("*.json"):
    try:
        raw = fp.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        continue
    if "avocado oil" not in raw.lower() or "bulgaria" not in raw.lower():
        continue
    if "importers distributors bulgaria" not in raw.lower() and "importers distributors Bulgaria" not in raw:
        # also allow close match
        if "importers distributors" not in raw.lower() or "bulgaria" not in raw.lower():
            continue
    try:
        data = json.loads(raw)
    except Exception:
        continue
    q = str(data.get("query") or "")
    if "bulgaria" not in q.lower() or "importer" not in q.lower():
        continue
    for h in data.get("results") or []:
        if not isinstance(h, dict):
            continue
        backend = str(h.get("backend") or "")
        title = str(h.get("title") or "")
        snippet = str(h.get("snippet") or "")
        blob = (title + "\n" + snippet).lower()
        found = [n for n in needles if n in blob]
        if "google_ai" in backend or found:
            rows.append(
                {
                    "q": q,
                    "backend": backend,
                    "title": title[:100],
                    "found": found,
                    "snippet": snippet.replace("\n", " ")[:500],
                }
            )

print(f"matching result rows: {len(rows)}")
gai = [r for r in rows if "google_ai" in r["backend"]]
print(f"google_ai rows: {len(gai)}")
print()
for r in gai:
    print(f"QUERY: {r['q']}")
    print(f"FOUND in scrape: {r['found'] or 'NONE of Interdrink/Eterika/Evterpa/Divisha/Zoya'}")
    print(f"SNIPPET:\n{r['snippet']}")
    print("-" * 60)

# explicit checks
text = "\n".join(r["snippet"] for r in gai).lower()
print("\n=== Checklist for this AI Overview ===")
for name, key in [
    ("Interdrink Co. Ltd.", "interdrink"),
    ("Eterika Ltd.", "eterika"),
    ("Evterpa Cosmetic", "evterpa"),
]:
    print(f"  {name}: {'YES extracted' if key in text else 'NO / not in Google AI scrape'}")
