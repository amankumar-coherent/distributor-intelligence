"""Estimate how many discovery-cache names look like real avocado-oil distributors."""
from __future__ import annotations

import json
import re
from collections import Counter
from pathlib import Path

CACHE = Path(r"D:\vendor-intel\data\search_cache")

CHANNEL_RE = re.compile(
    r"\b(?:distribut(?:or|ion|ing)|wholesal(?:er|e)|importer|reseller|"
    r"channel\s+partner|dealer|trader|trading\s+co|bulk\s+supplier|"
    r"foodservice|ingredient\s+distribut)\b",
    re.I,
)
MFG_RE = re.compile(
    r"\b(?:manufacturer|manufacturing|producer|processor|refinery|"
    r"cold[\s-]?press(?:ed|ing)?\s+mill|factory|oem\s+maker)\b",
    re.I,
)
JUNK_NAME_RE = re.compile(
    r"(?:market\s+(?:size|share|report|outlook|forecast|analysis)|"
    r"imarc|mordor|grand\s+view|fortune\s+business|research\s+and\s+markets|"
    r"wikipedia|linkedin|indiamart|tradeindia|alibaba|thomasnet|kompass|"
    r"justdial|top\s+\d+|list\s+of|buyers?\s+guide|pdf\b|"
    r"global\s+avocado\s+oil\s+market|"
    r"^(?:united\s+states|india|china|germany|france|brazil|mexico|nigeria|"
    r"australia|canada|japan|spain|italy|turkey|egypt|israel)$)",
    re.I,
)
COUNTRY_ONLY = re.compile(
    r"^(?:united\s+states|india|china|germany|france|brazil|mexico|nigeria|"
    r"australia|canada|japan|spain|italy|turkey|egypt|israel|singapore|"
    r"switzerland|sweden|poland|belgium|thailand|vietnam|malaysia|"
    r"philippines|indonesia|netherlands|saudi\s+arabia|uae|"
    r"united\s+arab\s+emirates|south\s+africa|argentina|chile|peru)$",
    re.I,
)

TITLE_CO_RE = re.compile(
    r"^(.{2,80}?)(?:\s*[|\-–—:]\s*|\s+[-–—]\s+|\s+\|\s+).+",
)


def clean_title(title: str) -> str:
    t = (title or "").strip()
    t = re.sub(r"\s+", " ", t)
    # strip common suffixes
    for sep in [" | ", " - ", " – ", " — ", ": "]:
        if sep in t:
            left = t.split(sep, 1)[0].strip()
            if 2 < len(left) < 70:
                t = left
                break
    t = re.sub(
        r"\b(?:official\s+site|home|about|products?|wholesale|distributor|suppliers?)\b.*$",
        "",
        t,
        flags=re.I,
    ).strip(" -|:")
    return t[:80]


def classify(name: str, blob: str) -> str:
    n = (name or "").strip()
    if len(n) < 3 or COUNTRY_ONLY.match(n) or JUNK_NAME_RE.search(n):
        return "junk"
    text = f"{n} {blob}"
    ch = bool(CHANNEL_RE.search(text))
    mf = bool(MFG_RE.search(text))
    oil = bool(re.search(r"avocado\s+oil|avocado\s+oils", text, re.I))
    if ch and not mf:
        return "likely_channel"
    if ch and mf:
        return "mixed_channel_mfg"
    if mf and not ch:
        return "likely_manufacturer"
    if oil and re.search(r"\b(?:ltd|llc|inc|gmbh|bv|sa|pvt|limited|group|oils?|foods?)\b", n, re.I):
        return "possible_participant"
    if oil:
        return "avocado_mention_weak"
    return "unclear"


def main() -> None:
    files = sorted(CACHE.glob("*.json"), key=lambda p: p.stat().st_mtime, reverse=True)[:400]
    # name -> best blob / evidence
    names: dict[str, dict] = {}
    avocado_files = 0

    for fp in files:
        try:
            raw = fp.read_text(encoding="utf-8", errors="ignore")
            data = json.loads(raw)
        except Exception:
            continue
        q = str(data.get("query") or "")
        if "avocado oil" not in q.lower():
            continue
        # Prefer distributor-ish queries (closer to current sweep)
        if not re.search(
            r"distribut|wholesale|importer|channel\s+partner|reseller|authorized",
            q,
            re.I,
        ):
            # still include some general avocado oil company queries
            if "avocado oil" not in q.lower():
                continue
        avocado_files += 1
        results = data.get("results") or data.get("hits") or []
        if not isinstance(results, list):
            continue
        for h in results:
            if not isinstance(h, dict):
                continue
            title = str(h.get("title") or "")
            snippet = str(h.get("snippet") or "")
            link = str(h.get("link") or h.get("url") or "")
            backend = str(h.get("backend") or "")
            name = clean_title(title)
            if not name:
                continue
            key = re.sub(r"[^a-z0-9]", "", name.lower())
            if len(key) < 3:
                continue
            blob = f"{snippet} {link} {q} {backend}"
            label = classify(name, blob)
            prev = names.get(key)
            # prefer stronger labels
            rank = {
                "likely_channel": 5,
                "mixed_channel_mfg": 4,
                "possible_participant": 3,
                "avocado_mention_weak": 2,
                "likely_manufacturer": 1,
                "unclear": 0,
                "junk": -1,
            }
            if not prev or rank.get(label, 0) > rank.get(prev["label"], 0):
                names[key] = {
                    "name": name,
                    "label": label,
                    "query": q[:90],
                    "backend": backend,
                    "snippet": snippet[:160],
                }

    counts = Counter(v["label"] for v in names.values())
    total = len(names)
    print(f"avocado_cache_files_scanned={avocado_files}")
    print(f"unique_cleaned_names={total}")
    print("--- classification ---")
    for k, n in counts.most_common():
        pct = 100.0 * n / total if total else 0
        print(f"  {n:4d}  ({pct:5.1f}%)  {k}")

    channelish = counts["likely_channel"] + counts["mixed_channel_mfg"]
    print()
    print(f"LIKELY distributors/channel (channel or mixed): {channelish}")
    print(f"  of which pure-ish channel: {counts['likely_channel']}")
    print(f"  mixed channel+mfg: {counts['mixed_channel_mfg']}")
    print(f"Likely manufacturers (not channel): {counts['likely_manufacturer']}")
    print(f"Weak / unclear / junk: {total - channelish - counts['likely_manufacturer']}")

    # Scale note vs terminal 527
    print()
    print("NOTE: Terminal shows ~527 unique discovery names in-memory.")
    print("This audit uses search_cache titles (proxy), not the live hit list.")
    print("Final channel-partner export gate will drop many more manufacturers/junk.")

    print("\n--- sample likely_channel ---")
    for v in [x for x in names.values() if x["label"] == "likely_channel"][:15]:
        print(f"  - {v['name']}")
    print("\n--- sample likely_manufacturer ---")
    for v in [x for x in names.values() if x["label"] == "likely_manufacturer"][:12]:
        print(f"  - {v['name']}")
    print("\n--- sample junk ---")
    for v in [x for x in names.values() if x["label"] == "junk"][:12]:
        print(f"  - {v['name']}")


if __name__ == "__main__":
    main()
