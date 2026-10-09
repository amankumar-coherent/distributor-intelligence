#!/usr/bin/env python3
"""Second pass for the 6th+7th October list: top companies up to 3 people.

Two groups need different asks, so they get different prompts:

* Companies with NOTHING. Re-asked from scratch. Many of these are
  genuinely non-India entities (US retailers, African banks) where blank
  is the right answer, so the prompt makes refusing easy and explicit.
* Companies with 1-2 people. Re-asked with the names we already hold
  listed in the prompt, so the model looks for DIFFERENT people instead
  of returning the same CEO again.

Results are written to a separate checkpoint and merged into the main one
by the caller, so a bad run can be discarded without touching good data.

Same anti-fabrication rule as the first pass: a name is written back only
if it appears verbatim in the reply, and name-pattern emails on the
company's own domain are dropped as constructed.
"""
import argparse
import json
import re
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT))

from run_leadership_aimode_batch import norm, parse_batch, tokens  # noqa: E402
from run_oct67 import SLOTS, looks_constructed, rank  # noqa: E402
from run_retail_pharma import reset_profile  # noqa: E402
from vendor_intel.placeholders.load_keys import apply_env_overrides  # noqa: E402

OUT_DIR = ROOT / "output" / "oct67"

PROMPT_EMPTY = """India leadership of each company below, plus website domain.

Up to 3 people each, INDIA-BASED ONLY. Priority: CTO and AI Head first, \
then CEO, Founder, Co-founder, COO, MD.

Check whether the company has an India office at all. If it has no India \
presence, return an empty people list - that is the correct answer, not \
a failure. Never substitute a global leader based abroad.

JSON array only, same order:
[{{"company":"","website":"","people":[{{"name":"","title":"","email":""}}]}}]

{companies}"""

PROMPT_MORE = """More India leaders for each company below.

We already have the people listed in brackets - give DIFFERENT people, \
INDIA-BASED ONLY. Priority: CTO and AI Head first, then CEO, Founder, \
Co-founder, COO, MD.

JSON array only, same order:
[{{"company":"","website":"","people":[{{"name":"","title":"","email":""}}]}}]

Empty list if there are no other India people in these roles - a guess \
is wrong.

{companies}"""


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--companies-file", required=True)
    ap.add_argument("--source", default=str(OUT_DIR / "oct67.json"))
    ap.add_argument("--checkpoint", required=True)
    ap.add_argument("--batch-size", type=int, default=5)
    ap.add_argument("--limit", type=int, default=None)
    ap.add_argument("--delay", type=float, default=3.0)
    ap.add_argument("--auto-reset", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    apply_env_overrides()

    source = json.loads(Path(args.source).read_text(encoding="utf-8"))
    companies = []
    seen = set()
    for raw in Path(args.companies_file).read_text(encoding="utf-8").splitlines():
        c = raw.strip()
        if c and c.lower() not in seen:
            seen.add(c.lower())
            companies.append(c)

    ckpt = Path(args.checkpoint)
    found = json.loads(ckpt.read_text(encoding="utf-8")) if ckpt.exists() else {}

    def current(name: str) -> list[dict]:
        """People already known, from the first pass plus this pass."""
        out = list((source.get(name.lower()) or {}).get("people") or [])
        have = {p.get("name", "").lower() for p in out}
        for p in (found.get(name.lower()) or {}).get("people") or []:
            if p.get("name", "").lower() not in have:
                out.append(p)
        return out

    empty, partial = [], []
    for c in companies:
        n = len(current(c))
        if n == 0:
            empty.append(c)
        elif n < SLOTS:
            partial.append(c)

    batches: list[tuple[str, list[str]]] = []
    for i in range(0, len(empty), args.batch_size):
        batches.append(("empty", empty[i: i + args.batch_size]))
    for i in range(0, len(partial), args.batch_size):
        batches.append(("more", partial[i: i + args.batch_size]))
    if args.limit is not None:
        batches = batches[: args.limit]

    print(f"no people: {len(empty)}   has 1-2: {len(partial)}   batches: {len(batches)}")

    def render(kind: str, members: list[str]) -> str:
        if kind == "empty":
            listing = "\n".join(f"{i+1}. {c}" for i, c in enumerate(members))
            return PROMPT_EMPTY.format(companies=listing)
        lines = []
        for i, c in enumerate(members):
            have = ", ".join(p["name"] for p in current(c) if p.get("name"))
            lines.append(f"{i+1}. {c} [have: {have}]")
        return PROMPT_MORE.format(companies="\n".join(lines))

    if args.dry_run:
        for kind in ("empty", "more"):
            b = next((m for k, m in batches if k == kind), None)
            if b:
                p = render(kind, b)
                print(f"\n--- {kind} prompt ({len(p)} chars) ---")
                print(p)
        return
    if not batches:
        print("no gaps")
        return

    from vendor_intel.scraping.google_ai_mode import ai_mode_enabled, ask, shutdown
    if not ai_mode_enabled():
        raise SystemExit("Set GOOGLE_AI_MODE_ENABLED=true in .env")

    added = rejected = dropped = 0
    try:
        for bi, (kind, members) in enumerate(batches, start=1):
            print(f"\n[batch {bi}/{len(batches)}] {kind}: {len(members)} companies ...", flush=True)
            answer = ""
            records: list[dict] = []
            for attempt in range(4):
                answer = ask(render(kind, members)) or ""
                low = answer.lower()
                if "request limit" in low or "try again in a little while" in low:
                    print(f"   RATE LIMITED (try {attempt+1}/4)", flush=True)
                    if args.auto_reset and reset_profile():
                        print("   profile + cookies reset - relaunching", flush=True)
                        time.sleep(5)
                        continue
                    time.sleep(120)
                    continue
                if "wasn't generated" in low or "was not generated" in low:
                    print(f"   Google declined (try {attempt+1}/4)", flush=True)
                    time.sleep(6 * (attempt + 1))
                    continue
                records = [r for r in parse_batch(answer) if "<" not in str(r.get("company") or "")]
                if records:
                    break
                print(f"   unparseable ({len(answer)} chars, try {attempt+1}/4)", flush=True)
                time.sleep(5)
            if not records:
                print("   giving up on this batch", flush=True)
                continue

            wanted = {norm(c): c for c in members}
            answer_tokens = tokens(answer)
            batch_added = 0
            for rec in records:
                company = wanted.get(norm(str(rec.get("company") or "")))
                if company is None:
                    continue
                site = " ".join(str(rec.get("website") or "").split())
                site = re.sub(r"^(?:https?:)?//", "", site).strip("/")

                known = {p.get("name", "").lower() for p in current(company)}
                fresh = []
                raw_people = rec.get("people")
                if not isinstance(raw_people, list):
                    raw_people = []
                for slot in raw_people:
                    if not isinstance(slot, dict):
                        continue
                    name = " ".join(str(slot.get("name") or "").split())
                    title = " ".join(str(slot.get("title") or "").split())
                    email = " ".join(str(slot.get("email") or "").split())
                    if not name or name.lower() in known:
                        continue
                    words = re.findall(r"[a-z0-9]{2,}", name.lower())
                    if words and any(w not in answer_tokens for w in words):
                        rejected += 1
                        continue
                    if email and email.lower() not in answer.lower():
                        email = ""
                    if email and looks_constructed(email, name, site):
                        email = ""
                        dropped += 1
                    known.add(name.lower())
                    fresh.append({"name": name, "title": title, "email": email})

                if not fresh and not site:
                    continue
                entry = found.setdefault(company.lower(), {"website": "", "people": []})
                entry["website"] = entry.get("website") or site
                room = SLOTS - len(current(company))
                if room > 0 and fresh:
                    fresh.sort(key=lambda p: rank(p.get("title", "")))
                    entry["people"].extend(fresh[:room])
                    batch_added += len(fresh[:room])

            ckpt.parent.mkdir(parents=True, exist_ok=True)
            ckpt.write_text(json.dumps(found, indent=1, ensure_ascii=False), encoding="utf-8")
            added += batch_added
            print(f"   +{batch_added} people   (rejected {rejected}, constructed emails {dropped})", flush=True)
            if bi < len(batches) and args.delay:
                time.sleep(args.delay)
    finally:
        try:
            shutdown()
        except Exception:  # noqa: BLE001
            pass

    print(f"\npeople added: {added}   rejected: {rejected}   constructed emails dropped: {dropped}")


if __name__ == "__main__":
    main()
