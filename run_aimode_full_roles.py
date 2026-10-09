#!/usr/bin/env python3
"""Google AI Mode lookup of full leadership for a company list, batched.

Improved over run_leadership_aimode_batch.py in three ways that matter for
obscure Indian SMEs, which is what is left after SalesQL:

  * asks for MD/CEO/COO/CFO/Country Head/CTO/AI Head/Marketing Head/Founder
    rather than a narrow pair, so one query harvests whatever exists;
  * tells the model where to look (company site "Leadership"/"About"
    pages, LinkedIn, ZaubaCorp/Tofler MCA filings, news) instead of
    relying on unaided recall, and asks for a website and source note so
    the answer is checkable;
  * accepts an "as listed" title, so a real title like "Director" or
    "Partner" is captured rather than being forced into a role that does
    not exist at a 20-person firm.

ONLY REAL DATA: every name written back must appear verbatim in the
model's own reply; anything unsupported is dropped and counted.

Usage:
    python run_aimode_full_roles.py --companies-file need.txt --dry-run
    python run_aimode_full_roles.py --companies-file need.txt
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

OUT_DIR = ROOT / "output" / "aimode_roles"
ROLES = [
    "MD", "CEO", "COO", "CFO", "COUNTRY HEAD", "CTO",
    "AI HEAD", "MARKETING HEAD", "Founder",
]

# Keep this prompt SHORT and lead with the output format.
#
# Measured failure: an earlier version opened with a long "where to look"
# briefing (LinkedIn People tab, ZaubaCorp, Crunchbase, ...). AI Mode then
# answered like a research assistant — "Data regarding the leadership teams
# was retrieved and compiled from corporate records..." — and never emitted
# the JSON array, so several batches returned 0 names. The shorter,
# format-first phrasing below was scoring ~82%.
PROMPT = """give me full name and email of md,ceo,coo,cfo,country head,cto,\
ai head,marketing head,founder and company website of these Indian companies. \
check linkedin, the company website and ZaubaCorp/MCA director records.

Reply with ONLY a JSON array, no commentary, one object per company in the \
same order:
[{{"company":"<name as given>","website":"","MD":{{"name":"","email":""}},\
"CEO":{{"name":"","email":""}},"COO":{{"name":"","email":""}},\
"CFO":{{"name":"","email":""}},"COUNTRY HEAD":{{"name":"","email":""}},\
"CTO":{{"name":"","email":""}},"AI HEAD":{{"name":"","email":""}},\
"MARKETING HEAD":{{"name":"","email":""}},"Founder":{{"name":"","email":""}},\
"other":{{"name":"","title":"","email":""}},"source":""}}]

Leave a field empty if you do not know it — an empty field is correct, a \
guessed name is wrong. Never take a person from a different company. If the \
head has another title (Director, Partner, Proprietor, Principal), put them \
in "other" with that title. Give an email only if it is a real published \
address, never info@ or a made-up pattern.

Companies:
{companies}"""

# Second pass. Each line names only the roles still blank for that company,
# so the model spends the answer on what is missing instead of restating
# what is already known.
GAP_PROMPT = """For each company below, find ONLY the people listed after \
the company name, plus the company website. Check LinkedIn, the company's \
own site and ZaubaCorp/MCA director records.

Reply with ONLY a JSON array, no commentary, one object per company in the \
same order:
[{{"company":"<name as given>","website":"","MD":{{"name":"","email":""}},\
"CEO":{{"name":"","email":""}},"COO":{{"name":"","email":""}},\
"CFO":{{"name":"","email":""}},"COUNTRY HEAD":{{"name":"","email":""}},\
"CTO":{{"name":"","email":""}},"AI HEAD":{{"name":"","email":""}},\
"MARKETING HEAD":{{"name":"","email":""}},"Founder":{{"name":"","email":""}},\
"other":{{"name":"","title":"","email":""}},"source":""}}]

Most of these are small private Indian companies. If a role genuinely does \
not exist there, leave it empty — an empty field is correct, a guessed name \
is wrong. Never take a person from a different company. If the company is \
run by a Director, Partner or Proprietor, put that person in "other" with \
their title. Give an email only if it is a real published address.

Companies and the roles still needed:
{companies}"""


def load_names(args) -> list[str]:
    if args.companies_file:
        return [l.strip() for l in Path(args.companies_file).read_text(encoding="utf-8").splitlines() if l.strip()]
    return _load_company_list(args.companies)


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
        other = entry.get("other") or {}
        row["OTHER NAME"] = other.get("name", "")
        row["OTHER TITLE"] = other.get("title", "")
        row["OTHER EMAIL"] = other.get("email", "")
        row["Found"] = sum(1 for r in ROLES if (entry.get(r) or {}).get("name")) + (
            1 if other.get("name") else 0
        )
        row["Source"] = entry.get("source", "")
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
                max(typical + 2, len(str(column)) + 2, 10), 34
            )
            for cell in ws[get_column_letter(idx)][1:]:
                cell.alignment = body
        ws.freeze_panes = "B2"
        ws.auto_filter.ref = ws.dimensions
        ws.row_dimensions[1].height = 24


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--companies", default=None, help=".xlsx/.csv list")
    parser.add_argument("--companies-file", default=None, help="plain text, one company per line")
    parser.add_argument("--checkpoint", default=str(OUT_DIR / "aimode_roles.json"))
    parser.add_argument("--output", default=str(OUT_DIR / "aimode_leadership.xlsx"))
    parser.add_argument(
        "--batch-size",
        type=int,
        default=10,
        help=(
            "Companies per query. 10 by default: a smaller batch leaves the "
            "model room to actually search each company rather than "
            "answering the easy ones and padding the rest."
        ),
    )
    parser.add_argument("--limit", type=int, default=None)
    parser.add_argument("--delay", type=float, default=4.0)
    parser.add_argument("--retry-empty", action="store_true", help="Re-query companies recorded with no names")
    parser.add_argument(
        "--gap-fill",
        action="store_true",
        help=(
            "Second pass: ask only for the roles each company is still "
            "missing. Existing names are never overwritten."
        ),
    )
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    apply_env_overrides()

    companies = load_names(args)
    ckpt = Path(args.checkpoint)
    data = json.loads(ckpt.read_text(encoding="utf-8")) if ckpt.exists() else {}

    def has_any(name: str) -> bool:
        entry = data.get(name.lower()) or {}
        return any((entry.get(r) or {}).get("name") for r in ROLES) or bool(
            (entry.get("other") or {}).get("name")
        )

    def missing_for(name: str) -> list[str]:
        entry = data.get(name.lower()) or {}
        return [r for r in ROLES if not (entry.get(r) or {}).get("name")]

    if args.gap_fill:
        # Anything with a blank role, plus companies missing only a website.
        todo = [
            c for c in companies
            if missing_for(c) or not (data.get(c.lower()) or {}).get("website")
        ]
    elif args.retry_empty:
        todo = [c for c in companies if not has_any(c)]
    else:
        todo = [c for c in companies if c.lower() not in data]
    if args.limit is not None:
        todo = todo[: args.limit]

    batches = [todo[i : i + args.batch_size] for i in range(0, len(todo), args.batch_size)]
    print(f"companies : {len(companies)}   to query: {len(todo)}   batches: {len(batches)} x {args.batch_size}")

    if args.dry_run:
        if batches:
            if args.gap_fill:
                listing = "\n".join(
                    f"{i+1}. {c} — need: {', '.join(missing_for(c) or ['website only'])}"
                    for i, c in enumerate(batches[0])
                )
                template = GAP_PROMPT
            else:
                listing = "\n".join(f"{i+1}. {c}" for i, c in enumerate(batches[0]))
                template = PROMPT
            print("\n--- prompt ---")
            print(template.format(companies=listing)[:1700], "...")
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
            if args.gap_fill:
                lines = []
                for i, c in enumerate(batch):
                    need = missing_for(c) or ["website only"]
                    lines.append(f"{i+1}. {c} — need: {', '.join(need)}")
                listing = "\n".join(lines)
                template = GAP_PROMPT
            else:
                listing = "\n".join(f"{i+1}. {c}" for i, c in enumerate(batch))
                template = PROMPT
            print(f"\n[batch {bi}/{len(batches)}] {len(batch)} companies ...", flush=True)
            # Google intermittently answers "Something went wrong and an AI
            # response wasn't generated" — measured even on a trivial
            # control question, while the very next multi-company query
            # succeeded. Treating that as "no data" silently loses a whole
            # batch, so retry it a few times before giving up.
            answer = ""
            records: list[dict] = []
            for attempt in range(4):
                answer = ask(template.format(companies=listing)) or ""
                if "wasn't generated" in answer or "was not generated" in answer:
                    print(f"   Google declined to answer (try {attempt + 1}/4) — retrying")
                    time.sleep(6 * (attempt + 1))
                    continue
                records = parse_batch(answer)
                # A parse that only recovers the prompt's own echoed
                # template is not an answer either.
                records = [r for r in records if "<" not in str(r.get("company") or "")]
                if records:
                    break
                print(f"   unparseable ({len(answer)} chars, try {attempt + 1}/4) — retrying")
                time.sleep(5)
            if not records:
                print("   giving up on this batch — will retry on resume")
                continue

            wanted = {norm(c): c for c in batch}
            answer_tokens = tokens(answer)
            batch_added = 0
            for rec in records:
                company = wanted.get(norm(str(rec.get("company") or "")))
                if company is None:
                    continue
                entry = {"website": " ".join(str(rec.get("website") or "").split()),
                         "source": " ".join(str(rec.get("source") or "").split())[:200]}
                for role in ROLES + ["other"]:
                    slot = rec.get(role) or {}
                    if not isinstance(slot, dict):
                        continue
                    name = " ".join(str(slot.get("name") or "").split())
                    email = " ".join(str(slot.get("email") or "").split())
                    if name:
                        # Grounding: the name must be in the model's own reply.
                        words = re.findall(r"[a-z0-9]{2,}", name.lower())
                        if words and any(w not in answer_tokens for w in words):
                            rejected += 1
                            name = email = ""
                    if email and email.lower() not in answer.lower():
                        email = ""
                    if name:
                        batch_added += 1
                    slot_out = {"name": name, "email": email}
                    if role == "other":
                        slot_out["title"] = " ".join(str(slot.get("title") or "").split())
                    entry[role] = slot_out
                if args.gap_fill:
                    # Never overwrite a name already found: a gap-fill pass
                    # may only add.
                    prev = data.get(company.lower()) or {}
                    merged = dict(prev)
                    merged["website"] = prev.get("website") or entry.get("website", "")
                    merged["source"] = entry.get("source") or prev.get("source", "")
                    for role in ROLES + ["other"]:
                        old = prev.get(role) or {}
                        new = entry.get(role) or {}
                        if old.get("name"):
                            merged[role] = old
                        elif new.get("name"):
                            merged[role] = new
                        else:
                            merged[role] = old or new
                    data[company.lower()] = merged
                else:
                    data[company.lower()] = entry

            ckpt.parent.mkdir(parents=True, exist_ok=True)
            ckpt.write_text(json.dumps(data, indent=1, ensure_ascii=False), encoding="utf-8")
            added += batch_added
            print(f"   +{batch_added} names   (rejected {rejected})")
            if bi < len(batches) and args.delay:
                time.sleep(args.delay)
    finally:
        try:
            shutdown()
        except Exception:  # noqa: BLE001
            pass

    write_excel(data, companies, Path(args.output))
    withname = sum(1 for c in companies if has_any(c))
    print(f"\nnames added : {added}")
    print(f"rejected    : {rejected} (not grounded in the answer)")
    print(f"companies with >=1 name: {withname} of {len(companies)}")
    print(f"Excel: {args.output}")


if __name__ == "__main__":
    main()
