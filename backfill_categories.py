#!/usr/bin/env python3
"""Refill Core Categories / Specialty Focus cells that echo the market name.

Measured on live state files: 227 of 414 Dry Eyes rows (55%) and 33 of 80
Solar Rooftop rows (41%) had

    Core Categories = "Global Solar Rooftop Market - Distributors Only"

The value says only that the company is in the market we asked about. It
carries no information, and it hides whether the company sells the right
product -- which is exactly what the product scope exists to establish.

The cause is fixed for NEW companies (the prompt no longer interpolates the
market into the field it asks about, and batch_runner._clean_category
blanks an echo). This repairs the rows already banked.

IMPORTANT -- run this only when no pipeline is running. It drives the same
persistent browser profile, and two processes on one profile collide on
the profile lock and neither makes progress. google_ai_mode now refuses
with ProfileInUse rather than looping, so a mistake is loud rather than
silent, but the run still will not start.

Existing good values are never touched: only cells that echo the market
are replaced, and a cell whose replacement cannot be established is left
exactly as it was rather than blanked, so this can only add information.

Usage:
    # see what would change
    python backfill_categories.py --market "Global Dry Eyes Treatment Market - Distributors Only" --dry-run

    # apply, then re-export
    python backfill_categories.py --market "Global Dry Eyes Treatment Market - Distributors Only" --export

    # every market that has a state file
    python backfill_categories.py --all --export
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

from dotenv import load_dotenv  # noqa: E402

load_dotenv(ROOT / ".env")

from vendor_intel.pipeline.market_research import discovery  # noqa: E402
from vendor_intel.pipeline.market_research.batch_runner import (  # noqa: E402
    _echoes_market,
)
from vendor_intel.pipeline.market_research.state import (  # noqa: E402
    STATE_DIR,
    market_slug,
)
from vendor_intel.scraping import google_ai_mode as gam  # noqa: E402

CHUNK_SIZE = 8          # smaller than discovery's 10: the reply carries prose
_ASK_RETRIES = 3

TARGET_FIELDS = ("Core Categories", "Specialty Focus")


def _prompt(market: str, payload: list[dict]) -> str:
    scope = discovery.product_scope_for(market)
    scope_line = f"\n{scope}\n" if scope else ""
    return (
        f'For each company below, state what it actually DISTRIBUTES in the '
        f'"{market}" market.\n{scope_line}\n'
        "Reply with ONLY a JSON object, no markdown fences, no commentary:\n"
        '{"results": [{"company": "<name exactly as given>", '
        '"core_categories": "<product categories it distributes>", '
        '"specialty_focus": "<its specific niche>"}]}\n'
        "Rules:\n"
        "- core_categories must be PRODUCT CATEGORIES, e.g. "
        '"Ophthalmic pharmaceuticals; OTC eye care" or "PV modules; '
        'inverters; mounting systems".\n'
        f'- NEVER repeat the market name. "{market}" is not a category.\n'
        "- If you do not know what a company distributes, use \"\" for both "
        "fields. An empty answer is kept as-is; a wrong one is worse.\n"
        "- Return one entry per company, in the same order."
    )


def _extract(text: str) -> list[dict] | None:
    """Pull the results array out of AI Mode's rendered answer.

    AI Mode renders the whole conversation, so the page text begins with
    the prompt read back verbatim — and this prompt CONTAINS a JSON
    template. A naive scan therefore locks onto the echoed template and
    returns its literal placeholders:

        [{"company": "<name exactly as given>",
          "core_categories": "<product categories it distributes>"}]

    which matches no company, so every row is "left unchanged" and the
    whole pass silently refills nothing (observed: refilled 0 of 73).

    Cutting at the LAST occurrence of the template marker is what makes
    this robust: the echo precedes the answer, so whatever follows the
    final marker is the generated payload. Same approach as
    discovery._strip_echoed_prompt.
    """
    if not text:
        return None
    cleaned = re.sub(r"^\s*```(?:json)?|```\s*$", "", text.strip(),
                     flags=re.I | re.M)

    # Drop everything up to the LAST echoed "results" opening.
    markers = [m.start() for m in re.finditer(r'\{\s*"results"', cleaned)]
    if len(markers) > 1:
        cleaned = cleaned[markers[-1]:]

    # A reply still carrying the template's placeholders is the echo, not
    # an answer.
    if "<name exactly as given>" in cleaned:
        tail = cleaned.split("<name exactly as given>")[-1]
        nxt = re.search(r'\{\s*"results"', tail)
        cleaned = tail[nxt.start():] if nxt else tail

    start = cleaned.find("{")
    while start != -1:
        depth = 0
        for i in range(start, len(cleaned)):
            if cleaned[i] == "{":
                depth += 1
            elif cleaned[i] == "}":
                depth -= 1
                if depth == 0:
                    try:
                        obj = json.loads(cleaned[start : i + 1])
                    except json.JSONDecodeError:
                        break
                    if isinstance(obj, dict) and isinstance(obj.get("results"), list):
                        return obj["results"]
                    break
        start = cleaned.find("{", start + 1)
    return None


def _norm(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(name or "").lower())


def _ask(market: str, chunk: list[dict]) -> dict[str, dict]:
    payload = [
        {"company": r.get("Company", ""), "website": r.get("Website", ""),
         "products": str(r.get("Key Products / Technologies") or "")[:200]}
        for r in chunk
    ]
    prompt = (
        _prompt(market, payload)
        + f"\n\nCompanies ({len(payload)}):\n"
        + json.dumps(payload, ensure_ascii=False)
    )
    for attempt in range(_ASK_RETRIES):
        text = gam.ask(prompt, retries=2)
        if text and not gam.is_ai_mode_refusal(text):
            results = _extract(text)
            if results:
                return {_norm(r.get("company")): r for r in results
                        if isinstance(r, dict)}
        print(f"      retry {attempt + 1}/{_ASK_RETRIES}")
        gam.shutdown()
        time.sleep(3)
    return {}


def backfill(market: str, *, dry_run: bool) -> dict:
    slug = market_slug(market)
    state_path = STATE_DIR / f"{slug}.json"
    if not state_path.is_file():
        print(f"  no state file: {state_path.name}")
        return {}

    data = json.loads(state_path.read_text(encoding="utf-8"))
    companies = data.get("companies") or []

    needs = [
        r for r in companies
        if any(_echoes_market(r.get(f), market) for f in TARGET_FIELDS)
    ]
    print(f"{market}")
    print(f"  {len(companies)} companies, {len(needs)} with an echoed category")
    if not needs:
        return {"fixed": 0, "total": len(companies)}

    if dry_run:
        print("  DRY RUN — sample of what would be refilled:")
        for r in needs[:12]:
            print(f"    {str(r.get('Company'))[:34]:36} "
                  f"CC={str(r.get('Core Categories'))[:38]}")
        return {"fixed": 0, "total": len(companies)}

    fixed = 0
    unresolved = 0
    total_chunks = (len(needs) + CHUNK_SIZE - 1) // CHUNK_SIZE
    for n, i in enumerate(range(0, len(needs), CHUNK_SIZE), 1):
        chunk = needs[i : i + CHUNK_SIZE]
        print(f"    chunk {n}/{total_chunks} ({len(chunk)} companies)...")
        answers = _ask(market, chunk)
        if not answers:
            # Leave the rows untouched. A blanked category would make the
            # row fail REQUIRED_FIELDS and drop the company at export --
            # strictly worse than an uninformative value.
            print("      no usable answer — leaving these rows unchanged")
            unresolved += len(chunk)
            continue

        for row in chunk:
            a = answers.get(_norm(row.get("Company"))) or {}
            changed = False
            if str(a.get("company") or "").startswith("<"):
                a = {}          # echoed template, not an answer
            for field, key in (("Core Categories", "core_categories"),
                               ("Specialty Focus", "specialty_focus")):
                if not _echoes_market(row.get(field), market):
                    continue          # existing value is fine, never overwrite
                new = " ".join(str(a.get(key) or "").split())
                if new.startswith("<") and new.endswith(">"):
                    continue    # template placeholder, never a real value
                if new and not _echoes_market(new, market):
                    row[field] = new
                    changed = True
            if changed:
                fixed += 1
            else:
                unresolved += 1

    print(f"  refilled {fixed}, left unchanged {unresolved}")

    if fixed:
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        backup = state_path.with_suffix(f".before-category-backfill-{stamp}.json")
        shutil.copy2(state_path, backup)
        data["companies"] = companies
        state_path.write_text(json.dumps(data, ensure_ascii=False, indent=2),
                              encoding="utf-8")
        print(f"  backup: {backup.name}")
        print(f"  state updated")

    return {"fixed": fixed, "total": len(companies)}


def main() -> None:
    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--market", action="append", default=[])
    ap.add_argument("--all", action="store_true",
                    help="every market with a state file")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--export", action="store_true",
                    help="re-export the .xlsx afterwards")
    args = ap.parse_args()

    markets = list(args.market)
    if args.all:
        from run_all_markets import MARKETS  # noqa: PLC0415

        markets = list(MARKETS)
    if not markets:
        ap.error("pass --market or --all")

    results = {}
    try:
        for market in markets:
            results[market] = backfill(market, dry_run=args.dry_run)
            print()
    finally:
        try:
            gam.shutdown()
        except Exception:
            pass

    if args.export and not args.dry_run:
        from vendor_intel.pipeline.market_research.export_xlsx import export_xlsx

        for market, r in results.items():
            if r.get("fixed"):
                print(f"re-exported: {export_xlsx(market)}")


if __name__ == "__main__":
    main()
