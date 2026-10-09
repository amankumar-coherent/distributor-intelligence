#!/usr/bin/env python3
"""Second-pass gap fill for the retail/pharma leadership sheet.

The broad pass asked every company for all eight designations at once. This
asks each company ONLY for the ones it is still missing, so the model's
output budget goes to the open slots instead of re-stating known answers.

Companies are grouped by their exact missing-role signature, so one query
covers several companies that need the same thing.

Same anti-fabrication rule as the broad pass: a name is written back only
if it appears verbatim in the model's own reply. Blanks here are usually
real -- most of these firms have no Co-owner and no CHRO -- so the prompt
says so explicitly rather than pressing for an answer.
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
from run_retail_pharma import ROLES, reset_profile  # noqa: E402
from vendor_intel.placeholders.load_keys import apply_env_overrides  # noqa: E402

OUT_DIR = ROOT / "output" / "retail_pharma_india"


def build_prompt(roles: list[str], companies: list[str]) -> str:
    """Ask for exactly these roles, nothing else."""
    slots = ",".join(
        '"{0}":{{"name":"","title":"","email":""}}'.format(r) if r == "OTHER HEAD"
        else '"{0}":{{"name":"","email":""}}'.format(r)
        for r in roles
    )
    listing = "\n".join("{0}. {1}".format(i + 1, c) for i, c in enumerate(companies))
    note = ""
    if "OTHER HEAD" in roles:
        note = ("OTHER HEAD = most senior INDIA person, with their real title "
                "(Chairman, President, Director, Proprietor, GM).\n")
    return (
        "For each company below give the INDIA person for only: {0}.\n\n".format(", ".join(roles))
        + "JSON array only, same order:\n"
        + '[{{"company":"",{0}}}]\n\n'.format(slots)
        + note
        + "India arm, not the global parent. Many firms genuinely have no CHRO, Owner or Co-owner - leave those "
          "empty rather than guessing. A wrong name is worse than a blank.\n\n"
        + listing
    )


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--companies-file", required=True)
    ap.add_argument("--source", default=str(OUT_DIR / "retail_pharma_india.json"),
                    help="Merged checkpoint read to work out what is missing.")
    ap.add_argument("--checkpoint", required=True, help="Where this pass writes its own findings.")
    ap.add_argument("--roles", default="CEO,MD,Founder,CHRO",
                    help="Only chase these designations.")
    ap.add_argument("--batch-size", type=int, default=6)
    ap.add_argument("--limit", type=int, default=None)
    ap.add_argument("--delay", type=float, default=3.0)
    ap.add_argument("--auto-reset", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    apply_env_overrides()
    want = [r.strip() for r in args.roles.split(",") if r.strip() in ROLES]

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

    # Group companies by the exact set of roles they still need, so one
    # query serves everyone with the same gap.
    groups: dict[tuple, list[str]] = {}
    for c in companies:
        entry = source.get(c.lower()) or {}
        already = found.get(c.lower()) or {}
        miss = tuple(
            r for r in want
            if not (entry.get(r) or {}).get("name") and not (already.get(r) or {}).get("name")
        )
        if miss:
            groups.setdefault(miss, []).append(c)

    batches = []
    for roles, members in sorted(groups.items(), key=lambda kv: -len(kv[1])):
        for i in range(0, len(members), args.batch_size):
            batches.append((list(roles), members[i: i + args.batch_size]))
    if args.limit is not None:
        batches = batches[: args.limit]

    gaps = sum(len(m) for m in groups.values())
    print("companies needing work: {0}   batches: {1}".format(gaps, len(batches)))
    for roles, members in sorted(groups.items(), key=lambda kv: -len(kv[1]))[:8]:
        print("   {0:>4} companies missing {1}".format(len(members), ", ".join(roles)))

    if args.dry_run:
        if batches:
            roles, members = batches[0]
            p = build_prompt(roles, members)
            print("\n--- prompt ---")
            print(p)
            print("\nprompt chars: {0}".format(len(p)))
        return
    if not batches:
        print("no gaps for these roles")
        return

    from vendor_intel.scraping.google_ai_mode import ai_mode_enabled, ask, shutdown
    if not ai_mode_enabled():
        raise SystemExit("Set GOOGLE_AI_MODE_ENABLED=true in .env")

    added = rejected = 0
    try:
        for bi, (roles, members) in enumerate(batches, start=1):
            print("\n[batch {0}/{1}] {2} companies x {3} roles ...".format(
                bi, len(batches), len(members), len(roles)), flush=True)
            answer = ""
            records: list[dict] = []
            for attempt in range(4):
                answer = ask(build_prompt(roles, members)) or ""
                low = answer.lower()
                if "request limit" in low or "try again in a little while" in low:
                    print("   RATE LIMITED (try {0}/4)".format(attempt + 1), flush=True)
                    if args.auto_reset and reset_profile():
                        print("   profile + cookies reset - relaunching", flush=True)
                        time.sleep(5)
                        continue
                    time.sleep(120)
                    continue
                if "wasn't generated" in low or "was not generated" in low:
                    print("   Google declined (try {0}/4)".format(attempt + 1), flush=True)
                    time.sleep(6 * (attempt + 1))
                    continue
                records = [r for r in parse_batch(answer) if "<" not in str(r.get("company") or "")]
                if records:
                    break
                print("   unparseable ({0} chars, try {1}/4)".format(len(answer), attempt + 1), flush=True)
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
                entry = found.setdefault(company.lower(), {})
                for role in roles:
                    slot = rec.get(role) or {}
                    if not isinstance(slot, dict):
                        continue
                    name = " ".join(str(slot.get("name") or "").split())
                    email = " ".join(str(slot.get("email") or "").split())
                    if name:
                        words = re.findall(r"[a-z0-9]{2,}", name.lower())
                        if words and any(w not in answer_tokens for w in words):
                            rejected += 1
                            name = email = ""
                    if email and email.lower() not in answer.lower():
                        email = ""
                    if not name:
                        continue
                    out = {"name": name, "email": email}
                    if role == "OTHER HEAD":
                        out["title"] = " ".join(str(slot.get("title") or "").split())
                    entry[role] = out
                    batch_added += 1

            ckpt.parent.mkdir(parents=True, exist_ok=True)
            ckpt.write_text(json.dumps(found, indent=1, ensure_ascii=False), encoding="utf-8")
            added += batch_added
            print("   +{0} names   (rejected {1})".format(batch_added, rejected), flush=True)
            if bi < len(batches) and args.delay:
                time.sleep(args.delay)
    finally:
        try:
            shutdown()
        except Exception:  # noqa: BLE001
            pass

    print("\nnames added: {0}   rejected: {1}".format(added, rejected))


if __name__ == "__main__":
    main()
