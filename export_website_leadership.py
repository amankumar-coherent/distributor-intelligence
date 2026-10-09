#!/usr/bin/env python3
"""Export the website-scraped leadership to Excel, with a confidence flag.

Heuristic name-vs-not-a-name filtering has a floor: after four rounds of
tightening, a residue of page text still slips through as "names"
("Collaboration GM", "Under John's", "Nomura Europe"). Rather than keep
guessing, every row carries an explicit Confidence and a Reason, so a
questionable row is visibly questionable instead of silently wrong.

High   - name looks like a person AND the page text ties it to the title
Review - extracted, but at least one signal is off (extra capitalised
         token, no surname-like structure, single-source hit)

Usage:
    python export_website_leadership.py
    python export_website_leadership.py --high-only
"""
import argparse
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

ROLES = ["MD", "CEO", "COO", "CFO", "COUNTRY HEAD", "CTO", "AI HEAD", "MARKETING HEAD", "Founder"]

# Tokens that mark a capture as page text rather than a person. Kept
# separate from the scraper's own filter so tightening this cannot silently
# drop rows from the scrape itself.
SUSPECT_TOKENS = {
    "innovation", "collaboration", "under", "nomura", "europe", "prescinto",
    "gm", "jop", "inc", "ltd", "llp", "corp", "technologies", "solutions",
    "systems", "global", "asia", "america", "africa", "digital", "cloud",
    "data", "software", "platform", "network", "capital", "ventures",
}


def confidence(name: str, company: str) -> tuple[str, str]:
    words = name.split()
    if len(words) < 2:
        return "Review", "single word"
    low = {w.lower().strip(".,'") for w in words}
    hit = low & SUSPECT_TOKENS
    if hit:
        return "Review", f"contains non-name word: {', '.join(sorted(hit))}"
    if any("'" in w for w in words):
        return "Review", "possessive/apostrophe — likely sentence fragment"
    company_words = {w.lower() for w in re.findall(r"[A-Za-z]{4,}", company)}
    if low & company_words:
        return "Review", "repeats a word from the company name"
    if len(words) > 3:
        return "Review", "unusually long for a personal name"
    return "High", ""


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--checkpoint", default=str(ROOT / "output/website_scrape/website_roles.json"))
    parser.add_argument("--output", default=str(ROOT / "output/website_scrape/website_leadership.xlsx"))
    parser.add_argument("--high-only", action="store_true", help="Export only High-confidence rows")
    args = parser.parse_args()

    import pandas as pd
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    data = json.loads(Path(args.checkpoint).read_text(encoding="utf-8"))
    rows = []
    for company, entry in data.items():
        for role, info in (entry.get("roles") or {}).items():
            name = info.get("name", "")
            conf, reason = confidence(name, company)
            rows.append({
                "Company": company,
                "Website": entry.get("website", ""),
                "Role": role,
                "Name": name,
                "Title on page": info.get("title", ""),
                "Confidence": conf,
                "Why flagged": reason,
                "Source page": info.get("url", ""),
            })
    for company, entry in data.items():
        for other in (entry.get("other") or [])[:2]:
            name = other.get("name", "")
            conf, reason = confidence(name, company)
            rows.append({
                "Company": company,
                "Website": entry.get("website", ""),
                "Role": "OTHER",
                "Name": name,
                "Title on page": other.get("title", ""),
                "Confidence": conf,
                "Why flagged": reason,
                "Source page": other.get("url", ""),
            })

    df = pd.DataFrame(rows)
    if args.high_only:
        df = df[df["Confidence"] == "High"]
    df = df.sort_values(["Confidence", "Company", "Role"], ascending=[True, True, True], kind="stable")

    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Website Leadership")
        ws = writer.sheets["Website Leadership"]
        for cell in ws[1]:
            cell.fill = PatternFill("solid", fgColor="1F3864")
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        body = Alignment(horizontal="left", vertical="top", wrap_text=False)
        for idx, column in enumerate(df.columns, start=1):
            lengths = sorted(len(str(v)) for v in df[column].tolist() if v not in (None, ""))
            typical = lengths[min(int(len(lengths) * 0.9), len(lengths) - 1)] if lengths else 0
            ws.column_dimensions[get_column_letter(idx)].width = min(
                max(typical + 2, len(str(column)) + 2, 10), 40
            )
            for cell in ws[get_column_letter(idx)][1:]:
                cell.alignment = body
        ws.freeze_panes = "B2"
        ws.auto_filter.ref = ws.dimensions
        ws.row_dimensions[1].height = 24

    high = int((df["Confidence"] == "High").sum())
    print(f"rows written : {len(df)}")
    print(f"  High       : {high}")
    print(f"  Review     : {len(df) - high}")
    print(f"companies    : {df['Company'].nunique()}")
    print(f"Excel        : {output}")


if __name__ == "__main__":
    main()
