#!/usr/bin/env python3
"""6th + 7th October list: up to 3 India-based leaders per company.

Designations, in the priority order the brief gives:
    CTO, AI Head  <- these two first
    CEO, Founder, Co-founder, COO, MD

INDIA-BASED ONLY. For a global company this means the India head / India
leadership, not the global person sitting abroad. An Indian arm with no
separate holder of a title should come back blank rather than inheriting
the global name.

UP TO 3 PEOPLE. The prompt asks for three slots rather than one field per
designation, so the model spends its output budget on three real people
instead of leaving five of seven fields empty. Each slot carries its own
designation, so the priority order is expressed as an instruction rather
than as column position.

Same anti-fabrication rule as every other pass in this project: a name is
written back only if it appears verbatim in the model's own reply.
Blanks are correct; invented names are not.
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
from run_retail_pharma import reset_profile  # noqa: E402
from vendor_intel.placeholders.load_keys import apply_env_overrides  # noqa: E402

OUT_DIR = ROOT / "output" / "oct67"
SLOTS = 3

# Accepted designations, normalised. Anything else the model returns is
# kept verbatim in the title column but flagged, so the user can see what
# came back rather than having it silently dropped.
WANTED = ["CTO", "AI Head", "CEO", "Founder", "Co-founder", "COO", "MD"]

PROMPT = """India leadership of each company below, plus website domain.

Give up to 3 people per company, INDIA-BASED ONLY. Priority: CTO and \
AI Head first, then CEO, Founder, Co-founder, COO, MD.

For a global company give the India head / India leadership, NOT the \
global person based abroad.

JSON array only, same order:
[{{"company":"","website":"","people":[{{"name":"","title":"","email":""}}]}}]

Fewer than 3 is fine if only 1-2 India people hold these titles. Empty \
people list if nobody India-based is findable - a guess is wrong.

{companies}"""


def load_names(path: str) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for raw in Path(path).read_text(encoding="utf-8").splitlines():
        name = raw.strip()
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
        people = entry.get("people") or []
        row = {"Company": company, "Website": entry.get("website", "")}
        for i in range(SLOTS):
            person = people[i] if i < len(people) else {}
            row[f"Person {i+1}"] = person.get("name", "")
            row[f"Designation {i+1}"] = person.get("title", "")
            row[f"Email {i+1}"] = person.get("email", "")
        row["Found"] = len([p for p in people if p.get("name")])
        rows.append(row)

    df = pd.DataFrame(rows)
    output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="India Leadership")
        ws = writer.sheets["India Leadership"]
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


def looks_constructed(email: str, name: str, website: str) -> bool:
    """True if the email is just the person's name pasted onto the domain.

    Measured on the smoke test: every single email AI Mode returned was
    exactly firstname.lastname@companydomain. The verbatim grounding check
    cannot catch these -- the model writes the string, so it IS in the
    reply -- but they are generated, not published. The user's standing
    rule is real data or a blank cell, so these are dropped.
    """
    if not email or "@" not in email:
        return False
    local, _, domain = email.lower().partition("@")
    site = (website or "").lower().lstrip("www.")
    if site and not (domain in site or site in domain):
        return False
    parts = [w for w in re.findall(r"[a-z]{2,}", (name or "").lower())]
    if not parts:
        return False
    local_words = [w for w in re.findall(r"[a-z]{2,}", local)]
    return bool(local_words) and all(w in parts for w in local_words)


def rank(title: str) -> int:
    """Priority order from the brief: CTO and AI Head first."""
    low = (title or "").lower()
    for i, w in enumerate(WANTED):
        if w.lower() in low:
            return i
    return len(WANTED)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--companies-file", required=True)
    ap.add_argument("--checkpoint", default=str(OUT_DIR / "oct67.json"))
    ap.add_argument("--output", default=str(OUT_DIR / "oct67.xlsx"))
    ap.add_argument("--batch-size", type=int, default=5)
    ap.add_argument("--limit", type=int, default=None)
    ap.add_argument("--delay", type=float, default=3.0)
    ap.add_argument("--retry-empty", action="store_true")
    ap.add_argument("--auto-reset", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    apply_env_overrides()

    companies = load_names(args.companies_file)
    ckpt = Path(args.checkpoint)
    data = json.loads(ckpt.read_text(encoding="utf-8")) if ckpt.exists() else {}

    def has_any(name: str) -> bool:
        entry = data.get(name.lower()) or {}
        return any(p.get("name") for p in (entry.get("people") or []))

    if args.retry_empty:
        todo = [c for c in companies if not has_any(c)]
    else:
        todo = [c for c in companies if c.lower() not in data]
    if args.limit is not None:
        todo = todo[: args.limit]

    batches = [todo[i: i + args.batch_size] for i in range(0, len(todo), args.batch_size)]
    print(f"companies : {len(companies)}   to query: {len(todo)}   batches: {len(batches)} x {args.batch_size}")

    if args.dry_run:
        if batches:
            listing = "\n".join(f"{i+1}. {c}" for i, c in enumerate(batches[0]))
            p = PROMPT.format(companies=listing)
            print("\n--- prompt ---")
            print(p)
            print(f"\nprompt chars: {len(p)}")
        return
    if not batches:
        write_excel(data, companies, Path(args.output))
        print(f"nothing to do; Excel refreshed: {args.output}")
        return

    from vendor_intel.scraping.google_ai_mode import ai_mode_enabled, ask, shutdown
    if not ai_mode_enabled():
        raise SystemExit("Set GOOGLE_AI_MODE_ENABLED=true in .env")

    added = rejected = 0
    constructed = [0]
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
                print("   giving up on this batch - will retry on resume", flush=True)
                continue

            wanted = {norm(c): c for c in batch}
            answer_tokens = tokens(answer)
            batch_added = 0
            for rec in records:
                company = wanted.get(norm(str(rec.get("company") or "")))
                if company is None:
                    continue
                people = []
                raw_people = rec.get("people")
                if not isinstance(raw_people, list):
                    raw_people = []
                for slot in raw_people:
                    if not isinstance(slot, dict):
                        continue
                    name = " ".join(str(slot.get("name") or "").split())
                    title = " ".join(str(slot.get("title") or "").split())
                    email = " ".join(str(slot.get("email") or "").split())
                    if not name:
                        continue
                    words = re.findall(r"[a-z0-9]{2,}", name.lower())
                    if words and any(w not in answer_tokens for w in words):
                        rejected += 1
                        continue
                    if email and email.lower() not in answer.lower():
                        email = ""
                    site_now = " ".join(str(rec.get("website") or "").split())
                    if email and looks_constructed(email, name, site_now):
                        email = ""
                        constructed[0] += 1
                    people.append({"name": name, "title": title, "email": email})
                # Priority order from the brief, then cap at 3.
                people.sort(key=lambda p: rank(p.get("title", "")))
                people = people[:SLOTS]
                batch_added += len(people)
                site = " ".join(str(rec.get("website") or "").split())
                site = re.sub(r"^(?:https?:)?//", "", site).strip("/")
                data[company.lower()] = {
                    "website": site,
                    "people": people,
                }

            ckpt.parent.mkdir(parents=True, exist_ok=True)
            ckpt.write_text(json.dumps(data, indent=1, ensure_ascii=False), encoding="utf-8")
            added += batch_added
            print(f"   +{batch_added} people   (rejected {rejected})", flush=True)
            if bi < len(batches) and args.delay:
                time.sleep(args.delay)
    finally:
        try:
            shutdown()
        except Exception:  # noqa: BLE001
            pass

    write_excel(data, companies, Path(args.output))
    print(f"\npeople added: {added}   rejected: {rejected}")
    print(f"companies with a name: {sum(1 for c in companies if has_any(c))} of {len(companies)}")
    print(f"Excel: {args.output}")


if __name__ == "__main__":
    main()
