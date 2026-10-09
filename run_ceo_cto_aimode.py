#!/usr/bin/env python3
"""Find the CEO and CTO of each company via Google AI Mode, batched.

Narrower than run_leadership_aimode_batch.py (which asks for seven roles):
asking for two roles keeps each company's answer short, so more companies
fit in one reply before AI Mode truncates it, and the model spends its
output on the roles that are actually wanted.

ONLY REAL DATA: a name is written back only when it appears verbatim in the
model's own answer. Names the answer does not support are dropped and
counted, never guessed, so a blank cell means "not found" rather than
"invented". This is the guard that caught an earlier fabrication where a
venture-capital firm's name was returned as a company's CEO.

Runs through Patchright + bundled Chromium with the CAPTCHA-solver
extension (GOOGLE_AI_MODE_* in .env).

Usage:
    python run_ceo_cto_aimode.py --companies "New company list.xlsx" --dry-run
    python run_ceo_cto_aimode.py --companies "New company list.xlsx"
    python run_ceo_cto_aimode.py --companies "New company list.xlsx" --retry-empty
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

from run_leadership_aimode_batch import _load_company_list, norm, parse_batch, tokens
from vendor_intel.placeholders.load_keys import apply_env_overrides

OUT_DIR = ROOT / "output" / "ceo_cto"
DEFAULT_CHECKPOINT = OUT_DIR / "ceo_cto_checkpoint.json"
DEFAULT_OUTPUT = OUT_DIR / "ceo_cto.xlsx"

PROMPT = """Give me the full name of the CEO and the CTO of each company \
listed below.

RULES:
- Only give a name you actually know for that specific company. If you do \
not know, leave it empty. An empty field is CORRECT; a wrong or guessed \
name is a serious error.
- Never use a person from a different company, and never use an investor, \
parent company or product name as a person.
- If one person holds both roles, repeat the name in both fields.
- Also give that person's email only if a real published address is known; \
otherwise leave the email empty.

Reply as a JSON array, one object per company, in the same order, no \
markdown and no commentary:
[{{"company":"<exact name as given>","CEO":{{"name":"","email":""}},\
"CTO":{{"name":"","email":""}}}}]

Companies:
{companies}"""

FIELDS = ["CEO", "CTO"]


def load_checkpoint(path: Path) -> dict:
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8")) or {}
    except (json.JSONDecodeError, OSError):
        return {}


def save_checkpoint(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    payload = json.dumps(data, indent=2, ensure_ascii=False)
    tmp.write_text(payload, encoding="utf-8")
    for attempt in range(5):
        try:
            tmp.replace(path)
            return
        except PermissionError:
            time.sleep(0.2 * (attempt + 1))
    path.write_text(payload, encoding="utf-8")
    tmp.unlink(missing_ok=True)


def write_excel(data: dict, companies: list[str], output: Path) -> None:
    import pandas as pd
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    records = []
    for company in companies:
        entry = data.get(company.lower()) or {}
        record = {"Company": company}
        found = 0
        for field in FIELDS:
            name, email = (entry.get(field) or ["", ""])[:2]
            record[field] = name
            record[f"{field} EMAIL"] = email
            if name:
                found += 1
        record["Found"] = found
        records.append(record)

    df = pd.DataFrame(records)
    output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        sheet = "CEO and CTO"
        df.to_excel(writer, index=False, sheet_name=sheet)
        ws = writer.sheets[sheet]
        for cell in ws[1]:
            cell.fill = PatternFill("solid", fgColor="1F3864")
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        body = Alignment(horizontal="left", vertical="top", wrap_text=False)
        for idx, column in enumerate(df.columns, start=1):
            lengths = sorted(len(str(v)) for v in df[column].tolist() if v not in (None, ""))
            typical = lengths[min(int(len(lengths) * 0.9), len(lengths) - 1)] if lengths else 0
            width = min(max(typical + 2, len(str(column)) + 2, 10), 34)
            ws.column_dimensions[get_column_letter(idx)].width = width
            for cell in ws[get_column_letter(idx)][1:]:
                cell.alignment = body
        ws.freeze_panes = "B2"
        ws.auto_filter.ref = ws.dimensions
        ws.row_dimensions[1].height = 24


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--companies", required=True, help=".xlsx or .csv company list")
    parser.add_argument("--checkpoint", default=str(DEFAULT_CHECKPOINT))
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT))
    parser.add_argument("--batch-size", type=int, default=25)
    parser.add_argument("--limit", type=int, default=None)
    parser.add_argument("--delay", type=float, default=3.0)
    parser.add_argument(
        "--retry-empty",
        action="store_true",
        help="Re-query companies already recorded with no CEO and no CTO",
    )
    parser.add_argument(
        "--seed-from",
        default=None,
        help="Existing leadership checkpoint to import known CEO/CTO from, so those are not re-queried",
    )
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    apply_env_overrides()

    companies = _load_company_list(args.companies)
    ckpt_path = Path(args.checkpoint)
    data = load_checkpoint(ckpt_path)

    if args.seed_from:
        seed = load_checkpoint(Path(args.seed_from))
        imported = 0
        for company in companies:
            key = company.lower()
            if key in data:
                continue
            values = (seed.get(key) or {}).get("values") or {}
            entry = {}
            for field in FIELDS:
                name, email = (values.get(field) or ["", ""])[:2]
                if name:
                    entry[field] = [name, email]
            if entry:
                data[key] = entry
                imported += 1
        if imported:
            save_checkpoint(ckpt_path, data)
            print(f"seeded {imported} companies with known CEO/CTO from {args.seed_from}")

    def has_any(key: str) -> bool:
        entry = data.get(key) or {}
        return any((entry.get(f) or ["", ""])[0] for f in FIELDS)

    if args.retry_empty:
        todo = [c for c in companies if not has_any(c.lower())]
    else:
        todo = [c for c in companies if c.lower() not in data]
    if args.limit is not None:
        todo = todo[: args.limit]

    batches = [todo[i : i + args.batch_size] for i in range(0, len(todo), args.batch_size)]
    print(f"companies : {len(companies)}")
    print(f"to query  : {len(todo)}   ({len(batches)} batches x {args.batch_size})")

    if args.dry_run:
        if batches:
            listing = "\n".join(f"{i+1}. {c}" for i, c in enumerate(batches[0]))
            print("\n--- first prompt ---")
            print(PROMPT.format(companies=listing)[:900], "...")
        return
    if not batches:
        print("nothing to do")
        write_excel(data, companies, Path(args.output))
        return

    from vendor_intel.scraping.google_ai_mode import ai_mode_enabled, ask, shutdown

    if not ai_mode_enabled():
        raise SystemExit("Set GOOGLE_AI_MODE_ENABLED=true in .env")

    added = rejected = 0
    try:
        for bi, batch in enumerate(batches, start=1):
            listing = "\n".join(f"{i+1}. {c}" for i, c in enumerate(batch))
            print(f"\n[batch {bi}/{len(batches)}] {len(batch)} companies ...", flush=True)
            answer = ask(PROMPT.format(companies=listing))
            if not answer:
                print("   no answer — skipping batch")
                continue
            records = parse_batch(answer)
            if not records:
                print(f"   unparseable answer ({len(answer)} chars) — skipping")
                continue

            wanted = {norm(c): c for c in batch}
            answer_tokens = tokens(answer)
            batch_added = 0
            for rec in records:
                company = wanted.get(norm(str(rec.get("company") or "")))
                if company is None:
                    continue
                entry = data.setdefault(company.lower(), {})
                for field in FIELDS:
                    slot = rec.get(field) or {}
                    if not isinstance(slot, dict):
                        continue
                    name = " ".join(str(slot.get("name") or "").split())
                    email = " ".join(str(slot.get("email") or "").split())
                    if not name or (entry.get(field) or ["", ""])[0]:
                        continue
                    # Only accept a name the answer itself contains.
                    words = re.findall(r"[a-z0-9]{2,}", name.lower())
                    if words and any(w not in answer_tokens for w in words):
                        rejected += 1
                        continue
                    if email and email.lower() not in answer.lower():
                        email = ""
                    entry[field] = [name, email]
                    batch_added += 1
                # Record even an empty result so a resume does not re-ask.
                data.setdefault(company.lower(), entry)
            save_checkpoint(ckpt_path, data)
            added += batch_added
            print(f"   +{batch_added} names   (rejected: {rejected})")
            if bi < len(batches) and args.delay > 0:
                time.sleep(args.delay)
    finally:
        try:
            shutdown()
        except Exception:  # noqa: BLE001
            pass

    write_excel(data, companies, Path(args.output))
    ceo = sum(1 for c in companies if ((data.get(c.lower()) or {}).get("CEO") or ["", ""])[0])
    cto = sum(1 for c in companies if ((data.get(c.lower()) or {}).get("CTO") or ["", ""])[0])
    both = sum(
        1
        for c in companies
        if ((data.get(c.lower()) or {}).get("CEO") or ["", ""])[0]
        and ((data.get(c.lower()) or {}).get("CTO") or ["", ""])[0]
    )
    print(f"\nnames added : {added}")
    print(f"rejected    : {rejected} (not grounded in the answer)")
    print(f"CEO found   : {ceo} of {len(companies)}")
    print(f"CTO found   : {cto} of {len(companies)}")
    print(f"both        : {both}")
    print(f"Excel: {args.output}")


if __name__ == "__main__":
    main()
