#!/usr/bin/env python3
"""Re-collect leadership names via Google AI Mode, batched per query.

Why AI Mode rather than the AI Overview relay: AI Mode is a real
generative answer, so unlike search-backed AI Overview it CAN answer a
batched "here are N companies, give me their leadership" prompt. That
makes a full re-collection affordable — 3,545 companies in ~178 queries
at the default batch size instead of 3,545 separate lookups.

Batch size is capped at 20 by measurement, not preference: at 20 the reply
parsed 20/20 companies, while at 30 AI Mode truncated the JSON mid-object
and only 1 record survived. parse_batch() also salvages complete objects
from a partially truncated reply, so an over-long answer degrades instead
of being lost.

It runs through Patchright + bundled Chromium with the CAPTCHA-solver
extension (GOOGLE_AI_MODE_* in .env); no AI Overview backend/relay on
:15551 is involved.

NO FABRICATION: every name written back must appear verbatim in the model's
own answer, and each answer is parsed per-company. Anything the answer does
not support is left blank rather than guessed.

Usage:
    python run_leadership_aimode_batch.py --dry-run
    python run_leadership_aimode_batch.py --fresh --limit 20   # one batch
    python run_leadership_aimode_batch.py --fresh              # all, new file
    python run_leadership_aimode_batch.py --max-names 3        # sparse rows only
"""
import argparse
import json
import re
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

from vendor_intel.placeholders.load_keys import apply_env_overrides
from vendor_intel.pipeline.leadership_lookup import (
    ROLES,
    Checkpoint,
    LeadershipRow,
    missing_roles,
    write_excel,
)

DEFAULT_CHECKPOINT = ROOT / "output" / "leadership_lookup" / "leadership_checkpoint.json"
DEFAULT_OUTPUT = ROOT / "output" / "leadership_lookup" / "leadership_contacts.xlsx"

BATCH_PROMPT = """For each company listed below, give me the full name and \
email of its MD, CEO, Country Head, CTO, AI Head, Marketing Head and \
Founder.

Give the email wherever a real, publicly known address for that person \
exists; leave the email empty when it does not, but still give the name.

CRITICAL ACCURACY RULES:
- Never invent or guess a person. If you do not know who holds a role at a \
company, leave it empty. An empty field is CORRECT; a wrong name is a \
serious error.
- Do not carry a person from one company to another. Each company's people \
must be that company's own.
- Many companies genuinely have no MD, Country Head or AI Head. Leave those \
empty rather than assigning someone.
- Only give an email you know is that individual's real published address. \
Never construct one from a name-and-domain pattern, and never use a generic \
inbox such as info@ or support@.

Reply as a JSON array, one object per company, in the same order, with no \
markdown and no commentary:
[{{"company":"<exact name as given>","MD":{{"name":"","email":""}},\
"CEO":{{"name":"","email":""}},"COUNTRY HEAD":{{"name":"","email":""}},\
"CTO":{{"name":"","email":""}},"AI HEAD":{{"name":"","email":""}},\
"MARKETING HEAD":{{"name":"","email":""}},"Founder":{{"name":"","email":""}}}}]

Companies:
{companies}"""


DEFAULT_COMPANY_CSV = r"C:\Users\aman.kumar\Downloads\TO RESEARCH PARTNER - Sheet1.csv"


def _load_company_list(source_path: str | None) -> list[str]:
    """Company names for a --fresh pass, de-duplicated, input order kept.

    Accepts .csv or .xlsx. Column choice: a column literally named
    "company" wins; otherwise the first column that actually contains
    values. That fallback matters because a hand-exported sheet often has
    an empty leading column (pandas names it "Unnamed: 0"), and picking
    column 0 blindly would yield an empty list.
    """
    import pandas as pd

    path = Path(source_path or DEFAULT_COMPANY_CSV)
    if path.suffix.lower() in (".xlsx", ".xls", ".xlsm"):
        df = pd.read_excel(path)
    else:
        df = pd.read_csv(path)

    column = next((c for c in df.columns if str(c).strip().lower() == "company"), None)
    if column is None:
        column = next(
            (c for c in df.columns if df[c].notna().any()),
            df.columns[0] if len(df.columns) else None,
        )
    if column is None:
        raise SystemExit(f"no usable column found in {path}")

    seen: set[str] = set()
    names: list[str] = []
    for value in df[column].tolist():
        name = str(value).strip()
        if not name or name.lower() == "nan" or name.lower() in seen:
            continue
        seen.add(name.lower())
        names.append(name)
    return names


def norm(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", (text or "").lower()).strip()


def tokens(text: str) -> set[str]:
    return set(re.findall(r"[a-z0-9]{2,}", (text or "").lower()))


def _balanced_arrays(text: str) -> list[str]:
    """Every top-level [...] span, tracking depth and string state."""
    spans: list[str] = []
    depth = 0
    start = -1
    in_string = False
    escape = False
    for i, ch in enumerate(text):
        if in_string:
            if escape:
                escape = False
            elif ch == "\\":
                escape = True
            elif ch == '"':
                in_string = False
            continue
        if ch == '"':
            in_string = True
        elif ch == "[":
            if depth == 0:
                start = i
            depth += 1
        elif ch == "]":
            if depth > 0:
                depth -= 1
                if depth == 0 and start >= 0:
                    spans.append(text[start : i + 1])
    return spans


def parse_batch(answer: str) -> list[dict]:
    """Pull the answer's JSON array out of the AI Mode transcript.

    AI Mode echoes the prompt before answering, and the prompt contains a
    `[{...}]` shape template. A naive first-"[" to last-"]" slice therefore
    spans the echoed template plus the real reply and fails to parse. Scan
    for balanced [...] spans instead and take the last one that yields
    objects carrying a "company" key.
    """
    text = (answer or "").strip()
    if text.startswith("```"):
        text = text.strip("`")
        text = text[4:] if text.lower().startswith("json") else text

    for span in reversed(_balanced_arrays(text)):
        try:
            data = json.loads(span, strict=False)
        except (json.JSONDecodeError, ValueError):
            continue
        if not isinstance(data, list):
            continue
        records = [d for d in data if isinstance(d, dict) and d.get("company")]
        if records:
            return records

    # Nothing parsed as a whole array. AI Mode truncates long answers
    # mid-JSON (measured: fine at 20 companies, cut off at 30), which
    # leaves a valid prefix of complete objects followed by a broken one.
    # Salvage every complete {...} object rather than discarding the batch.
    salvaged: list[dict] = []
    depth = 0
    start = -1
    in_string = False
    escape = False
    for i, ch in enumerate(text):
        if in_string:
            if escape:
                escape = False
            elif ch == "\\":
                escape = True
            elif ch == '"':
                in_string = False
            continue
        if ch == '"':
            in_string = True
        elif ch == "{":
            if depth == 0:
                start = i
            depth += 1
        elif ch == "}":
            if depth > 0:
                depth -= 1
                if depth == 0 and start >= 0:
                    try:
                        obj = json.loads(text[start : i + 1], strict=False)
                    except (json.JSONDecodeError, ValueError):
                        continue
                    company = obj.get("company") if isinstance(obj, dict) else None
                    # Skip the prompt's own echoed shape template.
                    if company and "<" not in str(company):
                        salvaged.append(obj)
    return salvaged


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--checkpoint", default=str(DEFAULT_CHECKPOINT))
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT))
    parser.add_argument(
        "--batch-size",
        type=int,
        default=20,
        help=(
            "Companies per AI Mode query. Measured ceiling: 20 parses "
            "cleanly, 30 gets truncated mid-JSON by AI Mode."
        ),
    )
    parser.add_argument("--limit", type=int, default=None, help="Only this many companies total")
    parser.add_argument(
        "--max-names",
        type=int,
        default=None,
        help="Only companies with fewer than this many names already (e.g. 3)",
    )
    parser.add_argument("--delay", type=float, default=3.0, help="Seconds between batches")
    parser.add_argument(
        "--fresh",
        action="store_true",
        help=(
            "Re-collect EVERY company from scratch instead of only filling "
            "blanks. Writes to its own checkpoint/Excel so existing results "
            "are left untouched, and nothing is carried over — every name in "
            "the output comes from this run's answers."
        ),
    )
    parser.add_argument(
        "--companies",
        "--companies-csv",
        dest="companies",
        default=None,
        help="Company list for --fresh: .csv or .xlsx (default: the original input CSV)",
    )
    parser.add_argument(
        "--retry-empty",
        action="store_true",
        help=(
            "With --fresh, re-query companies already recorded with zero "
            "names instead of skipping them. Existing names are kept."
        ),
    )
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    apply_env_overrides()

    def n_names(row: LeadershipRow) -> int:
        return sum(1 for name, _ in row.values.values() if name)

    if args.fresh:
        # Separate checkpoint/Excel so a fresh pass cannot damage the
        # existing results while it runs.
        ckpt_path = Path(args.checkpoint)
        if ckpt_path.name == Path(DEFAULT_CHECKPOINT).name:
            ckpt_path = ckpt_path.with_name("leadership_checkpoint_aimode.json")
        out_path = Path(args.output)
        if out_path.name == Path(DEFAULT_OUTPUT).name:
            out_path = out_path.with_name("leadership_contacts_aimode.xlsx")
        args.checkpoint, args.output = str(ckpt_path), str(out_path)

        companies = _load_company_list(args.companies)
        checkpoint = Checkpoint(args.checkpoint)
        # Any recorded company counts as done, including one the model had
        # no names for — that is an answer, not a gap to retry. With
        # --retry-empty the zero-name rows are re-queried instead: a batched
        # answer sometimes skips companies it does know, so a second ask
        # recovers some of them.
        if args.retry_empty:
            done = {
                r.company.lower()
                for r in checkpoint.all_rows()
                if any(name for name, _ in r.values.values())
            }
        else:
            done = {r.company.lower() for r in checkpoint.all_rows()}
        targets = [LeadershipRow(company=c) for c in companies if c.lower() not in done]
        for row in targets:
            for role in ROLES:
                row.values.setdefault(role, ("", ""))
        print(f"FRESH pass -> {args.checkpoint}")
        if done:
            print(f"resuming   : {len(done)} companies already collected this pass")
    else:
        checkpoint = Checkpoint(args.checkpoint)
        rows = checkpoint.all_rows()
        targets = [r for r in rows if missing_roles(r)]
        if args.max_names is not None:
            targets = [r for r in targets if n_names(r) < args.max_names]
        targets.sort(key=lambda r: n_names(r))

    if args.limit is not None:
        targets = targets[: args.limit]

    batches = [targets[i : i + args.batch_size] for i in range(0, len(targets), args.batch_size)]
    print(f"targets    : {len(targets)} companies")
    print(f"batches    : {len(batches)} x {args.batch_size}")

    if args.dry_run:
        if batches:
            names = "\n".join(f"{i+1}. {r.company}" for i, r in enumerate(batches[0]))
            print("\n--- first prompt ---")
            print(BATCH_PROMPT.format(companies=names)[:1400], "...")
        return
    if not batches:
        print("nothing to do")
        return

    from vendor_intel.scraping.google_ai_mode import ai_mode_enabled, ask, shutdown

    if not ai_mode_enabled():
        raise SystemExit("Set GOOGLE_AI_MODE_ENABLED=true in .env")

    added = kept = rejected = 0
    try:
        for bi, batch in enumerate(batches, start=1):
            listing = "\n".join(f"{i+1}. {r.company}" for i, r in enumerate(batch))
            print(f"\n[batch {bi}/{len(batches)}] {len(batch)} companies ...", flush=True)
            answer = ask(BATCH_PROMPT.format(companies=listing))
            if not answer:
                print("   no answer (CAPTCHA or timeout) — skipping batch")
                continue

            records = parse_batch(answer)
            if not records:
                print(f"   could not parse a JSON array from {len(answer)} chars — skipping")
                continue

            # Match each record back to the company we asked about.
            wanted = {norm(r.company): r for r in batch}
            answer_tokens = tokens(answer)
            batch_added = 0
            for rec in records:
                row = wanted.get(norm(str(rec.get("company") or "")))
                if row is None:
                    continue
                changed = False
                for role in ROLES:
                    entry = rec.get(role) or {}
                    if not isinstance(entry, dict):
                        continue
                    name = " ".join(str(entry.get("name") or "").split())
                    email = " ".join(str(entry.get("email") or "").split())
                    have_name, have_email = row.values.get(role) or ("", "")
                    if have_name or not name:
                        continue
                    # Grounding: every substantial word of the name must
                    # appear in the model's own answer. Cheap protection
                    # against a name invented during JSON assembly.
                    words = [w for w in re.findall(r"[a-z0-9]{2,}", name.lower())]
                    if words and any(w not in answer_tokens for w in words):
                        rejected += 1
                        continue
                    if email and email.lower() not in answer.lower():
                        email = ""
                    row.values[role] = (name, email or have_email)
                    changed = True
                    batch_added += 1
                # In --fresh mode record every company the answer covered,
                # even one with no names: otherwise a company the model
                # legitimately knows nothing about is never checkpointed and
                # gets re-queried on every resume.
                if changed or args.fresh:
                    # Store just this company's slice of the batch answer.
                    # Keeping 4000 chars of a shared 20k-char reply on every
                    # row bloated the checkpoint and stored mostly the
                    # echoed prompt rather than the company's own data.
                    row.raw_reply = json.dumps(rec, ensure_ascii=False)[:1200]
                    checkpoint.record(row)
            added += batch_added
            print(f"   +{batch_added} names   (rejected so far: {rejected})")
            if bi < len(batches) and args.delay > 0:
                time.sleep(args.delay)
    finally:
        try:
            shutdown()
        except Exception:  # noqa: BLE001
            pass

    final = Checkpoint(args.checkpoint).all_rows()
    total = sum(sum(1 for nm, _ in r.values.values() if nm) for r in final)
    filled = sum(1 for r in final if any(nm for nm, _ in r.values.values()))
    Path(args.output).parent.mkdir(parents=True, exist_ok=True)
    write_excel(final, args.output)
    print(f"\nnames added   : {added}")
    print(f"names rejected: {rejected}  (not grounded in the answer)")
    print(f"total names   : {total}")
    print(f"companies with >=1 name: {filled} of {len(final)}")
    print(f"Excel: {args.output}")


if __name__ == "__main__":
    main()
