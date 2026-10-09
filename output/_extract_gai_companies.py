import json
import re
from collections import Counter
from pathlib import Path

cache = Path(r"D:\vendor-intel\data\search_cache")
names = Counter()
samples = []

KNOWN = [
    "Chosen Foods",
    "Olivado",
    "La Tourangelle",
    "Bella Vado",
    "Westfalia",
    "Sesajal",
    "CalPure",
    "AvoPure",
    "Primal Kitchen",
    "Mantova",
    "Spectrum",
    "Brenntag",
    "Gustav Heess",
    "Avocado Real",
    "SVA Naturals",
    "SVA Organics",
    "Raasa",
    "Nutiva",
    "Avohass",
    "Kevala",
    "Aconcagua",
    "Village Press",
    "Grove Avocado",
    "Costco",
    "Walmart",
    "Whole Foods",
    "Ahuacatlan",
    "Avocado Health",
    "Santini",
    "VedaOils",
]

files = sorted(cache.glob("*.json"), key=lambda p: p.stat().st_mtime, reverse=True)[:250]
gai_files = 0

for fp in files:
    try:
        raw = fp.read_text(encoding="utf-8", errors="ignore")
        data = json.loads(raw)
    except Exception:
        continue

    if "google_ai_scraper" not in raw and "Google Search scrape" not in raw and "Google Search (cleaned)" not in raw:
        continue
    gai_files += 1
    query = str(data.get("query") or data.get("q") or "")
    hits = data.get("hits") or data.get("results") or []
    if not isinstance(hits, list):
        hits = []

    gai_blobs = []
    for h in hits:
        if not isinstance(h, dict):
            continue
        src = str(h.get("source") or "").lower()
        title = str(h.get("title") or "")
        snippet = str(h.get("snippet") or "")
        link = str(h.get("link") or h.get("url") or "")
        if (
            "google_ai" in src
            or "Google Search scrape" in title
            or "Google Search (cleaned)" in title
        ):
            gai_blobs.append(f"{title}\n{snippet}")
            if link and "google.com/search" not in link:
                host = re.sub(r"^https?://(www\.)?", "", link).split("/")[0]
                if host and "." in host and "google." not in host:
                    names[f"cite:{host}"] += 1

    blob = "\n".join(gai_blobs) if gai_blobs else raw
    for k in KNOWN:
        if re.search(re.escape(k), blob, re.I):
            names[k] += 1
            if len(samples) < 30:
                m = re.search(
                    r".{0,40}" + re.escape(k) + r".{0,60}",
                    blob,
                    re.I | re.S,
                )
                ctx = (m.group(0).replace("\n", " ") if m else "")[:120]
                samples.append((k, query[:70], ctx))

print(f"gai_cache_files={gai_files}")
print("\nCompanies found in Google AI scraper text:")
for name, cnt in names.most_common(40):
    if name.startswith("cite:"):
        continue
    print(f"  {cnt:3d}  {name}")

print("\nCitation domains from Google AI hits:")
for name, cnt in names.most_common(30):
    if name.startswith("cite:"):
        print(f"  {cnt:3d}  {name[5:]}")

print("\nSamples:")
seen = set()
for company, q, ctx in samples:
    key = company
    if key in seen:
        continue
    seen.add(key)
    print(f"- {company}")
    if q:
        print(f"  query: {q}")
    if ctx:
        print(f"  context: {ctx}")
