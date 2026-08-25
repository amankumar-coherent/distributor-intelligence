"""Audit Flexible Packaging FINAL: names, websites, location-split dupes."""
from __future__ import annotations

import re
from collections import defaultdict
from pathlib import Path
from urllib.parse import urlparse

from openpyxl import load_workbook

ROOT = Path(r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global")
files = sorted(ROOT.glob("*FINAL*.xlsx"))
if not files:
    files = sorted(ROOT.glob("*.xlsx"))
fp = files[0]
print("USING", fp)

wb = load_workbook(fp, read_only=True, data_only=True)
ws = wb.active
headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
print("N_COLS", len(headers))
for i, h in enumerate(headers):
    print(f"{i:02d} {h}")

# column map (case-insensitive contains)
def col(*names: str) -> int:
    for i, h in enumerate(headers):
        hs = str(h or "").strip().lower()
        for n in names:
            if hs == n.lower() or n.lower() in hs:
                return i
    return -1

ci = col("Company Name", "company")
wi = col("Website", "Company Website", "domain", "url")
print("company_col", ci, "website_col", wi)

rows = list(ws.iter_rows(min_row=2, values_only=True))
wb.close()
print("ROW_COUNT", len(rows))

names = []
websites = []
for r in rows:
    name = (r[ci] if ci >= 0 and ci < len(r) else None) or ""
    web = (r[wi] if wi >= 0 and wi < len(r) else None) or ""
    names.append(str(name).strip())
    websites.append(str(web).strip())

# exact unique
exact = set(n for n in names if n)
print("EXACT_UNIQUE_NAMES", len(exact))

# location/country suffixes often appended
LOC_WORDS = {
    "india", "usa", "us", "u.s.", "u.s.a.", "uk", "u.k.", "uae", "china", "japan",
    "germany", "france", "italy", "spain", "brazil", "mexico", "canada", "australia",
    "korea", "south korea", "thailand", "vietnam", "indonesia", "malaysia", "singapore",
    "taiwan", "philippines", "turkey", "poland", "netherlands", "belgium", "switzerland",
    "sweden", "norway", "denmark", "finland", "austria", "ireland", "portugal", "greece",
    "russia", "saudi arabia", "south africa", "egypt", "argentina", "chile", "colombia",
    "peru", "pakistan", "bangladesh", "sri lanka", "new zealand", "hong kong",
    "europe", "asia", "apac", "emea", "latam", "north america", "south america",
    "middle east", "global", "worldwide", "international",
    "california", "texas", "ohio", "illinois", "florida", "new york", "pennsylvania",
    "michigan", "georgia", "north carolina", "new jersey", "massachusetts",
    "mumbai", "delhi", "chennai", "bangalore", "bengaluru", "hyderabad", "pune",
    "ahmedabad", "kolkata", "gujarat", "maharashtra", "tamil nadu",
}

SUFFIX_LEGAL = re.compile(
    r"\b(inc\.?|incorporated|llc|l\.l\.c\.|ltd\.?|limited|corp\.?|corporation|"
    r"co\.?|plc|gmbh|ag|sa|bv|nv|pty|pvt\.?|private|public|holdings?|group|"
    r"company|companies)\b\.?$",
    re.I,
)

LOC_TAIL = re.compile(
    r"[\s,\-/|]+("
    + "|".join(re.escape(w) for w in sorted(LOC_WORDS, key=len, reverse=True))
    + r")\s*$",
    re.I,
)

# also "… Packaging India" / "… Packaging USA"
PACK_LOC = re.compile(
    r"([\s,\-/]+(?:flexible\s+)?packaging[\s,\-/]+("
    + "|".join(re.escape(w) for w in sorted(LOC_WORDS, key=len, reverse=True))
    + r"))\s*$",
    re.I,
)


def strip_loc(name: str) -> str:
    n = name.strip()
    n = PACK_LOC.sub("", n).strip(" ,-/|")
    # strip trailing country after legal suffix: "Foo Inc India"
    n2 = LOC_TAIL.sub("", n).strip(" ,-/|")
    if n2 != n:
        n = n2
    # strip trailing country before legal: "Foo India Inc"
    n = re.sub(
        r"[\s,\-/]+("
        + "|".join(re.escape(w) for w in sorted(LOC_WORDS, key=len, reverse=True))
        + r")[\s,\-/]+(inc\.?|llc|ltd\.?|corp\.?|gmbh|ag|sa|plc)\s*$",
        r" \2",
        n,
        flags=re.I,
    )
    return re.sub(r"\s+", " ", n).strip(" ,-/|")


def canon_key(name: str) -> str:
    n = strip_loc(name)
    n = SUFFIX_LEGAL.sub("", n).strip(" ,./")
    n = re.sub(r"[^\w\s]", " ", n.lower())
    n = re.sub(r"\s+", " ", n).strip()
    # drop trailing packaging / flexible packaging noise for grouping only
    n = re.sub(r"\b(flexible\s+)?packaging\b$", "", n).strip()
    return n


def domain_of(url: str) -> str:
    u = (url or "").strip()
    if not u:
        return ""
    if not re.match(r"^https?://", u, re.I):
        u = "https://" + u
    try:
        host = urlparse(u).netloc.lower()
    except Exception:
        return ""
    host = host.split("@")[-1]
    if host.startswith("www."):
        host = host[4:]
    return host


# Group by canon
groups: dict[str, list[tuple[int, str, str]]] = defaultdict(list)
for i, (n, w) in enumerate(zip(names, websites), start=2):
    if not n:
        continue
    groups[canon_key(n)].append((i, n, w))

multi = {k: v for k, v in groups.items() if len(v) > 1}
print("CANON_UNIQUE", len(groups))
print("MULTI_GROUPS", len(multi))
print("--- sample multi groups (first 40) ---")
shown = 0
for k, v in sorted(multi.items(), key=lambda x: -len(x[1])):
    shown += 1
    if shown > 40:
        break
    print(f"\n[{len(v)}] key={k!r}")
    for row, n, w in v[:8]:
        print(f"  r{row}: {n!r} | {w!r}")
    if len(v) > 8:
        print(f"  ... +{len(v)-8} more")

# names ending oddly with Inc / location patterns
print("\n--- names with location-ish tails ---")
locish = []
for i, n in enumerate(names, start=2):
    if not n:
        continue
    sn = strip_loc(n)
    if sn != n:
        locish.append((i, n, sn))
print("LOCISH_COUNT", len(locish))
for row, n, sn in locish[:60]:
    print(f"  r{row}: {n!r} -> {sn!r}")

print("\n--- Inc/LLC endings sample ---")
incish = [n for n in names if re.search(r"\b(inc|llc|ltd|corp)\b\.?\s*$", n, re.I)]
print("INCISH", len(incish))
for n in sorted(set(incish))[:40]:
    print(" ", n)

# website issues: empty, non-http junk, mismatch vs name
print("\n--- website domain samples / issues ---")
bad_web = []
for i, (n, w) in enumerate(zip(names, websites), start=2):
    d = domain_of(w)
    if not w:
        bad_web.append((i, n, w, "empty"))
    elif not d:
        bad_web.append((i, n, w, "unparseable"))
    elif any(x in d for x in ("linkedin.com", "facebook.com", "twitter.com", "wikipedia.org", "crunchbase.com", "zoominfo.com")):
        bad_web.append((i, n, w, "social_or_directory"))
print("BAD_WEB_COUNT", len(bad_web))
for item in bad_web[:50]:
    print(" ", item)

# print all company+website for user review dump
out = ROOT / "_company_website_audit.tsv"
with out.open("w", encoding="utf-8") as f:
    f.write("row\tcompany\twebsite\tdomain\tcanon\tstripped\n")
    for i, (n, w) in enumerate(zip(names, websites), start=2):
        f.write(f"{i}\t{n}\t{w}\t{domain_of(w)}\t{canon_key(n)}\t{strip_loc(n)}\n")
print("WROTE", out)
