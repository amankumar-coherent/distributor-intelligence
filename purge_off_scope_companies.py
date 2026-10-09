#!/usr/bin/env python3
"""Delete companies that do not sell the market's actual product.

Why this exists. The pipeline's distributor definition constrains what a
company IS (a wholesaler, not a reseller) but never constrained what it
SELLS. For "Global Dry Eyes Treatment Market" that let through ophthalmic
EQUIPMENT and DEVICE distributors -- diagnostic systems, clinical lasers,
slit lamps, microsurgical instruments. Every one is a real ophthalmic
wholesale distributor. None distributes a dry-eye drug, so none belongs in
a dry-eye TREATMENT market.

Measured on the live state file: of 175 collected companies only 24 were
unambiguously drug distributors, 98 were device-only, and 49 mentioned
both. discovery.product_scope_for() stops new ones arriving; this removes
the ones already banked.

How it decides. Two stages, because neither alone is trustworthy:

  1. A keyword pass splits the obvious cases. Device-only text with no
     drug term at all is a confident delete; drug text with no device term
     is a confident keep.
  2. Everything ambiguous -- and, unless --fast, every keyword decision
     too -- goes to Google AI Mode in chunks, which is asked for the
     specific product that justifies keeping the company.

Fail-safe direction is deliberately the OPPOSITE of the verifier's. The
verifier drops on doubt because a wrong KEEP pollutes the output. Here a
wrong DELETE destroys a real company permanently, so an unreadable or
missing answer KEEPS the company and reports it for review.

Nothing is overwritten in place: the state file is backed up first, and
every decision is written to a JSON audit log.

Usage:
    # see what would go, change nothing
    python purge_off_scope_companies.py --market "Global Dry Eyes Treatment Market - Distributors Only" --dry-run

    # apply, then re-export the workbook
    python purge_off_scope_companies.py --market "Global Dry Eyes Treatment Market - Distributors Only" --export

    # keyword-only, no AI Mode calls (fast, cruder)
    python purge_off_scope_companies.py --market "..." --fast --dry-run
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
from vendor_intel.pipeline.market_research.state import (  # noqa: E402
    STATE_DIR,
    market_slug,
)
from vendor_intel.scraping import google_ai_mode as gam  # noqa: E402

CHUNK_SIZE = 10
_ASK_RETRIES = 3

# Fields describing what the company actually sells.
_PRODUCT_FIELDS = (
    "Core Categories",
    "Specialty Focus",
    "Key Products / Technologies",
)


# --- keyword pre-pass -----------------------------------------------------
#
# Per-market term sets. Only markets with a product scope defined in
# discovery are listed; anything else goes straight to AI Mode, since
# invented keywords would be worse than none.
_KEYWORDS: dict[str, tuple[str, str]] = {
    "dry eye": (
        # ON-scope: dry-eye drugs and therapeutic formulations
        r"artificial tear|lubricant(ing)? eye ?drop|eye ?drop|ophthalmic emulsion"
        r"|cyclosporin|ciclosporin|lifitegrast|xiidra|restasis|ikervis|cequa"
        r"|miebo|perfluorohexyloctane|tyrvaya|varenicline|punctal|ocular insert"
        r"|hyaluron|carboxymethylcellulose|povidone|ophthalmic (gel|ointment|solution|suspension)"
        r"|dry eye (drug|pharmaceutic|therapeut|treatment|medicat)"
        r"|ophthalmic (pharmaceutic|drug|medicat)|eye care pharmaceutic",
        # OFF-scope: equipment, devices, optical retail
        r"\b(equipment|device|instrument|laser|diagnostic|slit ?lamp|meibograph"
        r"|tear.?film analys|tomograph|\bOCT\b|phaco|microsurg|surgical instrument"
        r"|\bIPL\b|spectacle|frame|contact lens|optical retail|\bIOL\b"
        r"|intraocular lens|cataract|refractive surger|machine|apparatus)\b",
    ),
    "glucometer": (
        r"glucomet|glucose (meter|monitor|sensor)|\bCGM\b|continuous glucose"
        r"|test strip|lancet|freestyle libre|dexcom|guardian sensor|blood glucose",
        r"\b(insulin pump only|fitness tracker|smartwatch|activity band"
        r"|blood pressure monitor|nebuli[sz]er)\b",
    ),
    "silicon carbide": (
        r"silicon carbide|\bSiC\b|MOSFET|schottky|JBS diode|power module"
        r"|wafer|substrate|epitax|abrasive|refractor|technical ceramic",
        r"\b(gallium nitride only|GaN-only|silicon only|passive component only)\b",
    ),
    "solar rooftop": (
        r"\bPV\b|photovoltaic|solar (module|panel|inverter)|inverter"
        r"|mounting (structure|system)|racking|balance of system|\bBOS\b",
        r"\b(EPC|engineering procurement|installer|project developer"
        r"|solar thermal|water heater)\b",
    ),
    "polyurethane": (
        r"polyurethane|\bPU\b|\bMDI\b|\bTDI\b|polyol|isocyanate|prepolymer"
        r"|\bfoam\b|elastomer|adhesive|sealant|coating|\bTPU\b",
        r"\b(furniture retail|mattress retail|machinery only|machine only)\b",
    ),
}


def _keyword_sets(market: str) -> tuple[re.Pattern, re.Pattern] | None:
    low = (market or "").lower()
    for needle, (on, off) in _KEYWORDS.items():
        if needle in low:
            return re.compile(on, re.I), re.compile(off, re.I)
    return None


def _product_text(row: dict) -> str:
    return " | ".join(str(row.get(f) or "") for f in _PRODUCT_FIELDS)


# --- the AI Mode adjudication --------------------------------------------


def _scope_prompt(market: str, scope: str) -> str:
    """Ask for the JUSTIFYING PRODUCT, not a yes/no.

    Naming the product is what makes the answer checkable: a model that
    cannot name a qualifying product has, in effect, said no, whereas a
    bare boolean invites a confident guess.
    """
    return (
        f'You are auditing a company list for the market "{market}".\n\n'
        f"{scope}\n\n"
        "For EACH company below decide whether it genuinely distributes a "
        "product inside the REQUIRED PRODUCT SCOPE above.\n"
        "Reply with ONLY a JSON object, no markdown fences, no commentary:\n"
        '{"results": [{"company": "<name exactly as given>", '
        '"in_scope": true|false, '
        '"product": "<the specific in-scope product it distributes, or \\"\\">", '
        '"confidence": 0-100}]}\n'
        "Rules:\n"
        "- in_scope=true REQUIRES naming a specific qualifying product in "
        '"product". A generic answer like "pharmaceuticals" or "ophthalmic '
        'products" is NOT specific enough - use false.\n'
        "- A company that only sells equipment, devices, instruments or "
        "machinery for this condition is in_scope=false.\n"
        "- If you do not know the company, in_scope=false with confidence 0.\n"
        "- Return one entry per company, in the same order."
    )


def _extract(text: str) -> list[dict] | None:
    """Pull the results array out of AI Mode's rendered answer."""
    if not text:
        return None
    cleaned = re.sub(r"^\s*```(?:json)?|```\s*$", "", text.strip(),
                     flags=re.I | re.M)
    # Widest balanced object containing "results"
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
                    if isinstance(obj, dict) and "results" in obj:
                        res = obj["results"]
                        return res if isinstance(res, list) else None
                    break
        start = cleaned.find("{", start + 1)
    return None


def _norm(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(name or "").lower())


def _ask_chunk(market: str, scope: str, chunk: list[dict]) -> dict[str, dict]:
    """Returns {normalised company name: verdict}. Empty dict on failure,
    which the caller treats as KEEP."""
    payload = [
        {
            "company": r.get("Company", ""),
            "website": r.get("Website", ""),
            "categories": _product_text(r)[:400],
        }
        for r in chunk
    ]
    prompt = (
        _scope_prompt(market, scope)
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
        print(f"      retry {attempt + 1}/{_ASK_RETRIES} (empty or unparseable)")
        gam.shutdown()
        time.sleep(3)
    return {}


# --- main -----------------------------------------------------------------


def purge(market: str, *, dry_run: bool, fast: bool,
          min_confidence: int) -> dict:
    slug = market_slug(market)
    state_path = STATE_DIR / f"{slug}.json"
    if not state_path.is_file():
        print(f"no state file for {market!r}: {state_path}")
        return {}

    data = json.loads(state_path.read_text(encoding="utf-8"))
    companies = data.get("companies") or []
    if not companies:
        print("state file holds no companies")
        return {}

    scope = discovery.product_scope_for(market)
    if not scope:
        print(f"no product scope defined for {market!r} — nothing to enforce.")
        print("Add one to discovery._PRODUCT_SCOPES first, or the purge would")
        print("be filtering on invented product terms.")
        return {}

    kw = _keyword_sets(market)
    print(f"{market}")
    print(f"  {len(companies)} companies in state")
    print(f"  scope: {len(scope)} chars, keywords: {'yes' if kw else 'no'}")
    print(f"  mode: {'keyword only' if fast else 'keyword + AI Mode'}")
    print()

    keep: list[dict] = []
    removed: list[dict] = []
    audit: list[dict] = []
    needs_ai: list[dict] = []

    # Stage 1 -- keywords split the obvious cases.
    for row in companies:
        text = _product_text(row)
        if not kw or not text.strip():
            needs_ai.append(row)
            continue
        on_re, off_re = kw
        has_on = bool(on_re.search(text))
        has_off = bool(off_re.search(text))

        if has_off and not has_on and fast:
            # Device-only with NO in-scope term at all. Only trusted in
            # --fast mode: measured against the live data, the keyword
            # pass alone wrongly deleted DKSH Holding AG, which really
            # does distribute the Alcon Systane dry-eye portfolio but
            # also carries Topcon diagnostic equipment, so the off-scope
            # terms fired. A broad-line distributor legitimately sells
            # both, so only AI Mode can separate them.
            removed.append(row)
            audit.append({
                "company": row.get("Company"), "decision": "removed",
                "stage": "keyword", "confidence": 100,
                "reason": "off-scope product terms only, no in-scope product term",
            })
        elif has_on and not has_off and fast:
            keep.append(row)
            audit.append({
                "company": row.get("Company"), "decision": "kept",
                "stage": "keyword", "confidence": 100,
                "reason": "in-scope product terms, no off-scope terms",
            })
        else:
            # Mixed, or clean-but-unconfirmed in non-fast mode.
            needs_ai.append(row)

    if fast:
        for row in needs_ai:
            keep.append(row)
            audit.append({
                "company": row.get("Company"), "decision": "kept",
                "stage": "keyword", "confidence": 0,
                "reason": "ambiguous; kept because --fast makes no AI call",
            })
        needs_ai = []

    print(f"  keyword pass: {len(removed)} removed, {len(keep)} kept, "
          f"{len(needs_ai)} to adjudicate")

    # Stage 2 -- AI Mode adjudicates the rest.
    if needs_ai:
        total_chunks = (len(needs_ai) + CHUNK_SIZE - 1) // CHUNK_SIZE
        for n, i in enumerate(range(0, len(needs_ai), CHUNK_SIZE), 1):
            chunk = needs_ai[i : i + CHUNK_SIZE]
            print(f"    chunk {n}/{total_chunks} ({len(chunk)} companies)...")
            verdicts = _ask_chunk(market, scope, chunk)

            if not verdicts:
                # FAIL-SAFE KEEPS. Opposite of the verifier on purpose: a
                # wrong delete here is unrecoverable.
                print("      no usable answer — KEEPING this chunk for review")
                for row in chunk:
                    keep.append(row)
                    audit.append({
                        "company": row.get("Company"), "decision": "kept",
                        "stage": "ai-mode", "confidence": 0,
                        "reason": "backend gave no usable answer — kept fail-safe",
                        "needs_review": True,
                    })
                continue

            for row in chunk:
                v = verdicts.get(_norm(row.get("Company"))) or {}
                in_scope = bool(v.get("in_scope"))
                product = str(v.get("product") or "").strip()
                try:
                    conf = int(float(v.get("confidence") or 0))
                except (TypeError, ValueError):
                    conf = 0

                if not v:
                    keep.append(row)
                    audit.append({
                        "company": row.get("Company"), "decision": "kept",
                        "stage": "ai-mode", "confidence": 0,
                        "reason": "company missing from the answer — kept fail-safe",
                        "needs_review": True,
                    })
                elif in_scope and product and conf >= min_confidence:
                    keep.append(row)
                    audit.append({
                        "company": row.get("Company"), "decision": "kept",
                        "stage": "ai-mode", "confidence": conf,
                        "reason": f"in-scope product: {product}",
                    })
                elif in_scope and not product:
                    # Claimed in-scope but could not name a product: that
                    # is the shape of a guess, so it is not a keep.
                    removed.append(row)
                    audit.append({
                        "company": row.get("Company"), "decision": "removed",
                        "stage": "ai-mode", "confidence": conf,
                        "reason": "claimed in-scope but named no specific product",
                    })
                elif in_scope and conf < min_confidence:
                    keep.append(row)
                    audit.append({
                        "company": row.get("Company"), "decision": "kept",
                        "stage": "ai-mode", "confidence": conf,
                        "reason": f"in-scope but low confidence ({conf}) — kept fail-safe",
                        "needs_review": True,
                    })
                else:
                    removed.append(row)
                    audit.append({
                        "company": row.get("Company"), "decision": "removed",
                        "stage": "ai-mode", "confidence": conf,
                        "reason": "not in the required product scope",
                    })

    print()
    print(f"  RESULT: keep {len(keep)}, remove {len(removed)}")
    review = [a for a in audit if a.get("needs_review")]
    if review:
        print(f"  {len(review)} kept for review (no confident answer)")

    out_dir = ROOT / "output" / "market_research"
    out_dir.mkdir(parents=True, exist_ok=True)
    audit_path = out_dir / f"{slug}_scope_purge_audit.json"
    audit_path.write_text(
        json.dumps({
            "market": market,
            "when": datetime.now().isoformat(timespec="seconds"),
            "dry_run": dry_run,
            "before": len(companies),
            "kept": len(keep),
            "removed": len(removed),
            "decisions": audit,
        }, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"  audit: {audit_path}")

    if dry_run:
        print()
        print("  DRY RUN — state file unchanged. Sample of what would go:")
        for a in [x for x in audit if x["decision"] == "removed"][:20]:
            print(f"    - {a['company']}  ({a['reason'][:60]})")
        return {"kept": len(keep), "removed": len(removed)}

    if removed:
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        backup = state_path.with_suffix(f".before-scope-purge-{stamp}.json")
        shutil.copy2(state_path, backup)
        print(f"  backup: {backup.name}")

        data["companies"] = keep
        state_path.write_text(json.dumps(data, ensure_ascii=False, indent=2),
                              encoding="utf-8")
        print(f"  state updated: {len(keep)} companies remain")
    else:
        print("  nothing removed — state file untouched")

    return {"kept": len(keep), "removed": len(removed)}


def main() -> None:
    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--market", required=True)
    ap.add_argument("--dry-run", action="store_true",
                    help="report what would be removed, change nothing")
    ap.add_argument("--fast", action="store_true",
                    help="keyword pass only, no AI Mode calls (cruder)")
    ap.add_argument("--min-confidence", type=int, default=60,
                    help="below this an in-scope claim is kept for review "
                         "rather than trusted (default 60)")
    ap.add_argument("--export", action="store_true",
                    help="re-export the .xlsx after purging")
    args = ap.parse_args()

    try:
        result = purge(args.market, dry_run=args.dry_run, fast=args.fast,
                       min_confidence=args.min_confidence)
    finally:
        try:
            gam.shutdown()
        except Exception:
            pass

    if args.export and result.get("removed") and not args.dry_run:
        from vendor_intel.pipeline.market_research.export_xlsx import export_xlsx

        path = export_xlsx(args.market)
        print(f"  re-exported: {path}")


if __name__ == "__main__":
    main()
