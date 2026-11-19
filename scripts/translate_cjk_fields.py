"""Translate Japanese / Chinese / Korean text in a market's discovery data to
English, in place, before verification and export.

Why: the deliverable Excel and JSON must carry ENGLISH names (user, jpf
2026-10-29), but AI Mode returns many Japan companies under their Japanese
registered name ("株式会社クオリア"). Those names also broke verification
(the verifier's ASCII name key reduced them all to ""), so 286 of 825 Japan
fashion companies never got a verdict.

What it does, for every company in data/market_research/<slug>.json:
  - Company -> the company's official English name (its own English usage
    when it has one, else a faithful romanisation with "Co., Ltd." /
    "Inc." for 株式会社 / 有限会社), using the website as a hint;
  - Contact Person -> romanised, given name first;
  - every other text field with CJK characters -> plain English.
Calls DeepSeek (the repo's configured OpenAI-compatible client) in batches,
backs up the JSON first, and is safe to re-run (only CJK values are sent).

Usage: python scripts/translate_cjk_fields.py --market "Japan Fashion E-commerce Market"
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT))

from vendor_intel.config import Settings  # noqa: E402
from vendor_intel.pipeline.chatgpt_expand import _chat_json, _client, _model  # noqa: E402

CJK = re.compile(r"[぀-ヿ㐀-鿿가-힯！-～ｦ-ﾟ]")
BATCH = 10

SYSTEM = (
    "You translate business data from Japanese, Chinese or Korean into English. "
    "Reply with JSON only: {\"items\": [{\"id\": <id>, \"en\": \"<english>\"}]}, "
    "one item per input id, never skip one.\n"
    "Rules by field:\n"
    "- Company: the company's OFFICIAL ENGLISH NAME as the company itself uses it "
    "(website, English site, annual report). If it has none, romanise the name "
    "faithfully (Hepburn) and render the legal form in English: 株式会社 -> "
    "\"Co., Ltd.\", 有限会社 -> \"Ltd.\", 合同会社 -> \"LLC\". Use the website "
    "domain as a hint (e.g. unique-jean.co.jp -> \"Unique Jean\"). Never translate "
    "a brand name into its meaning. No Japanese characters in the answer.\n"
    "- Contact Person: romanise (Hepburn), given name first, family name last.\n"
    "- Role: the English job title (代表取締役 -> \"Representative Director\", "
    "代表取締役社長 -> \"President and Representative Director\").\n"
    "- Any other field: a short, plain English translation.\n"
    "Keep any English already in the value."
)


def market_slug(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--market", required=True)
    args = ap.parse_args()

    path = ROOT / "data" / "market_research" / f"{market_slug(args.market)}.json"
    data = json.loads(path.read_text(encoding="utf-8"))
    companies = data["companies"]

    jobs = []  # (company index, field, value, website)
    for i, c in enumerate(companies):
        for field, value in c.items():
            if isinstance(value, str) and CJK.search(value):
                jobs.append((i, field, value, str(c.get("Website") or "")))
    print(f"{args.market}: {len(companies)} companies, {len(jobs)} CJK values to translate")
    if not jobs:
        return

    backup = path.with_suffix(f".before_translate_{time.strftime('%Y%m%d_%H%M%S')}.json")
    shutil.copy2(path, backup)
    print(f"backup -> {backup.name}")

    settings = Settings()
    client, model = _client(settings), _model(settings)
    done = 0
    for start in range(0, len(jobs), BATCH):
        batch = jobs[start:start + BATCH]
        payload = [
            {"id": n, "field": f, "value": v, **({"website": w} if f == "Company" else {})}
            for n, (_i, f, v, w) in enumerate(batch)
        ]
        try:
            out = _chat_json(client, model, SYSTEM, json.dumps(payload, ensure_ascii=False),
                             label="translate", max_tokens=8000)
        except Exception as exc:  # noqa: BLE001 - one bad batch must not stop the rest
            print(f"  batch {start // BATCH + 1}: failed ({exc}); left as is")
            continue
        items = (out or {}).get("items", []) if isinstance(out, dict) else []
        by_id = {int(x["id"]): str(x.get("en") or "").strip()
                 for x in items if isinstance(x, dict) and str(x.get("id", "")).isdigit()}
        for n, (i, f, v, _w) in enumerate(batch):
            en = by_id.get(n, "")
            if en and not CJK.search(en):
                companies[i][f] = en
                done += 1
        # Save after every batch so an interruption keeps the work.
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"  batch {start // BATCH + 1}/{-(-len(jobs) // BATCH)}: {done} translated so far")

    left = sum(1 for c in companies for v in c.values() if isinstance(v, str) and CJK.search(v))
    print(f"DONE: {done} values translated, {left} still contain CJK text")


if __name__ == "__main__":
    main()
