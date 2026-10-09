#!/usr/bin/env python3
"""Leadership lookup via Google AI Mode for the AI_NO_DATA company list.

Roles are the ones the source sheet asks for in its "Designations to find"
column: CTO, AI Head, CEO, Founder, Co-founder, COO, MD — name and email
for each, plus the company website/domain.

Designed to be driven by run_allroles_parallel.py, which gives every
worker its own Chrome profile and its own captcha-raptor copy.

Two hard-won details are preserved from the CEO/CHRO version of this
script and matter just as much here:

* Keep the prompt SHORT. AI Mode echoes the prompt back TWICE in the page
  transcript, and the reply is capped a few thousand characters. A
  1,022-char prompt left ~250 characters for data and silently truncated
  the JSON array, abandoning ~300 companies as "unparseable" even though
  Google had answered. Seven roles need more output per company than two,
  so the batch size is smaller here.
* Google's "request limit" page is not an answer. With --auto-reset the
  worker wipes BOTH its Chrome and its Chromium profile (the run switches
  builds to load the solver; clearing one leaves the other's cookies and
  the limit in place) and relaunches.

ONLY REAL DATA: a name or email is written back only if it appears
verbatim in the model's own reply; anything else is dropped and counted.
"""
import argparse
import json
import os
import re
import shutil
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT))

from run_leadership_aimode_batch import norm, parse_batch, tokens  # noqa: E402
from vendor_intel.placeholders.load_keys import apply_env_overrides  # noqa: E402

OUT_DIR = ROOT / "output" / "ai_nodata_core3"
ROLES = ["CEO", "AI HEAD", "CTO"]

# Short on purpose — see the module docstring.
PROMPT = """CEO, AI Head and CTO of each company below, with website domain.

Search the company website About/Team/Leadership page and LinkedIn.

JSON array only, no commentary, same order:
[{{"company":"","website":"","CEO":{{"name":"","email":""}},"AI HEAD":{{"name":"","email":""}},"CTO":{{"name":"","email":""}}}}]

Empty field if unknown — a guess is wrong. If a founder runs the company and there is no separate CEO, put that person as CEO. Many small firms have no AI Head or CTO; leave those empty. Email only if really published.

{companies}"""


def reset_profile() -> bool:
    """Wipe this worker's profile + cookies so a rate limit clears."""
    from vendor_intel.scraping import google_ai_mode as gam

    try:
        gam.shutdown()
    except Exception:  # noqa: BLE001
        pass
    time.sleep(2)

    base = Path(
        (os.getenv("GOOGLE_AI_MODE_PROFILE_DIR") or "").strip()
        or str(ROOT / "data" / "ai_mode_chrome_profile")
    )
    stamp = time.strftime("%H%M%S")
    removed = False
    for path in (base, Path(f"{base}_chromium")):
        if path.is_dir():
            try:
                shutil.rmtree(path, ignore_errors=True)
                removed = True
            except Exception:  # noqa: BLE001
                try:
                    path.rename(Path(f"{path}.stale_{stamp}"))
                    removed = True
                except Exception:  # noqa: BLE001
                    pass
    for owner in base.parent.glob(f"{base.name}*.owner"):
        owner.unlink(missing_ok=True)
    return removed


def load_names(path: str) -> list[str]:
    lines = Path(path).read_text(encoding="utf-8").splitlines()
    seen: set[str] = set()
    out: list[str] = []
    for raw in lines:
        name = str(raw).strip()
        if not name or name.lower() in ("nan", "company") or name.lower() in seen:
            continue
        seen.add(name.lower())
        out.append(name)
    return out


def write_excel(data: dict, companies: list[str], output: Path) -> None:
    import pandas as pd
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    rows = []
    for company in companies:
        entry = data.get(company.lower()) or {}
        row = {"Company": company, "Website": entry.get("website", "")}
        for role in ROLES:
            person = entry.get(role) or {}
            row[role] = person.get("name", "")
            row[f"{role} EMAIL"] = person.get("email", "")
        row["Found"] = sum(1 for r in ROLES if (entry.get(r) or {}).get("name"))
        rows.append(row)

    df = pd.DataFrame(rows)
    output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Leadership")
        ws = writer.sheets["Leadership"]
        for cell in ws[1]:
            cell.fill = PatternFill("solid", fgColor="1F3864")
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        body = Alignment(horizontal="left", vertical="top", wrap_text=False)
        for idx, column in enumerate(df.columns, start=1):
            lengths = sorted(len(str(v)) for v in df[column].tolist() if v not in (None, ""))
            typical = lengths[min(int(len(lengths) * 0.9), len(lengths) - 1)] if lengths else 0
            ws.column_dimensions[get_column_letter(idx)].width = min(
                max(typical + 2, len(str(column)) + 2, 11), 34
            )
            for cell in ws[get_column_letter(idx)][1:]:
                cell.alignment = body
        ws.freeze_panes = "B2"
        ws.auto_filter.ref = ws.dimensions
        ws.row_dimensions[1].height = 24


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--companies-file", required=True)
    parser.add_argument("--checkpoint", default=str(OUT_DIR / "ai_nodata_core3.json"))
    parser.add_argument("--output", default=str(OUT_DIR / "ai_nodata_core3.xlsx"))
    parser.add_argument(
        "--batch-size",
        type=int,
        default=8,
        help="Companies per query. 8 here: only three roles are requested, "
             "so each company needs much less output than the seven-role pass.",
    )
    parser.add_argument("--limit", type=int, default=None)
    parser.add_argument("--delay", type=float, default=3.0)
    parser.add_argument("--retry-empty", action="store_true")
    parser.add_argument("--auto-reset", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    apply_env_overrides()

    companies = load_names(args.companies_file)
    ckpt = Path(args.checkpoint)
    data = json.loads(ckpt.read_text(encoding="utf-8")) if ckpt.exists() else {}

    def has_any(name: str) -> bool:
        entry = data.get(name.lower()) or {}
        return any((entry.get(r) or {}).get("name") for r in ROLES)

    if args.retry_empty:
        todo = [c for c in companies if not has_any(c)]
    else:
        todo = [c for c in companies if c.lower() not in data]
    if args.limit is not None:
        todo = todo[: args.limit]

    batches = [todo[i : i + args.batch_size] for i in range(0, len(todo), args.batch_size)]
    print(f"companies : {len(companies)}   to query: {len(todo)}   batches: {len(batches)} x {args.batch_size}")

    if args.dry_run:
        if batches:
            listing = "\n".join(f"{i+1}. {c}" for i, c in enumerate(batches[0]))
            print("\n--- prompt ---")
            print(PROMPT.format(companies=listing))
            print(f"\nprompt chars: {len(PROMPT.format(companies=listing))}")
        return
    if not batches:
        write_excel(data, companies, Path(args.output))
        print(f"nothing to do; Excel refreshed: {args.output}")
        return

    from vendor_intel.scraping.google_ai_mode import ai_mode_enabled, ask, shutdown

    if not ai_mode_enabled():
        raise SystemExit("Set GOOGLE_AI_MODE_ENABLED=true in .env")

    added = rejected = 0
    try:
        for bi, batch in enumerate(batches, start=1):
            listing = "\n".join(f"{i+1}. {c}" for i, c in enumerate(batch))
            print(f"\n[batch {bi}/{len(batches)}] {len(batch)} companies ...", flush=True)

            answer = ""
            records: list[dict] = []
            for attempt in range(4):
                answer = ask(PROMPT.format(companies=listing)) or ""
                low = answer.lower()
                if "request limit" in low or "try again in a little while" in low:
                    print(f"   RATE LIMITED (try {attempt + 1}/4)", flush=True)
                    if args.auto_reset and reset_profile():
                        print("   profile + cookies reset — relaunching", flush=True)
                        time.sleep(5)
                        continue
                    time.sleep(120)
                    continue
                if "wasn't generated" in low or "was not generated" in low:
                    print(f"   Google declined (try {attempt + 1}/4)", flush=True)
                    time.sleep(6 * (attempt + 1))
                    continue
                records = [r for r in parse_batch(answer) if "<" not in str(r.get("company") or "")]
                if records:
                    break
                print(f"   unparseable ({len(answer)} chars, try {attempt + 1}/4)", flush=True)
                time.sleep(5)
            if not records:
                print("   giving up on this batch — will retry on resume", flush=True)
                continue

            wanted = {norm(c): c for c in batch}
            answer_tokens = tokens(answer)
            batch_added = 0
            for rec in records:
                company = wanted.get(norm(str(rec.get("company") or "")))
                if company is None:
                    continue
                entry = {"website": " ".join(str(rec.get("website") or "").split())}
                for role in ROLES:
                    slot = rec.get(role) or {}
                    if not isinstance(slot, dict):
                        slot = {}
                    name = " ".join(str(slot.get("name") or "").split())
                    email = " ".join(str(slot.get("email") or "").split())
                    if name:
                        words = re.findall(r"[a-z0-9]{2,}", name.lower())
                        if words and any(w not in answer_tokens for w in words):
                            rejected += 1
                            name = email = ""
                    if email and email.lower() not in answer.lower():
                        email = ""
                    if name:
                        batch_added += 1
                    entry[role] = {"name": name, "email": email}
                data[company.lower()] = entry

            ckpt.parent.mkdir(parents=True, exist_ok=True)
            ckpt.write_text(json.dumps(data, indent=1, ensure_ascii=False), encoding="utf-8")
            added += batch_added
            print(f"   +{batch_added} names   (rejected {rejected})", flush=True)
            if bi < len(batches) and args.delay:
                time.sleep(args.delay)
    finally:
        try:
            shutdown()
        except Exception:  # noqa: BLE001
            pass

    write_excel(data, companies, Path(args.output))
    withname = sum(1 for c in companies if has_any(c))
    print(f"\nnames added : {added}   rejected: {rejected}")
    print(f"companies with a name: {withname} of {len(companies)}")
    print(f"Excel: {args.output}")


if __name__ == "__main__":
    main()
