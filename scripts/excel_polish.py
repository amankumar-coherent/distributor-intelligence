"""Presentation styling for the distributor workbooks (client deliverable).

Layout of a *_verified_distributors.xlsx sheet: rows 1-3 = market title,
distributor definition, note (merged across); row 4 = spacer; row 5 = header;
data from row 6. This module restyles that sheet WITHOUT changing the data,
apart from three tidy-ups that stop Excel flagging numbers stored as text:
"#" is renumbered 1..N (rows were removed or renamed) and "Founded" years
become numbers.

Look: navy title banner, soft-blue definition band, a one-line summary bar,
navy header with a gold underline, banded rows, light-grey hairline borders,
no gridlines, readable 10 pt text sized to content, company names in bold,
clickable links, frozen header + company column, print-ready landscape with
the header repeated. The main sheets hold verified companies ONLY: companies
flagged for a suspicious ownership claim are kept out and listed (amber) in
a separate flagged_for_review.xlsx.
"""
from __future__ import annotations

import math
import re
from urllib.parse import quote

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

NAVY = "1F3864"
BAND_BLUE = "E8EEF7"
STAT_BLUE = "D6E2F3"
ROW_ALT = "F3F7FC"
GRID = "D0D9E6"
GOLD = "C9A227"
AMBER = "FFF2CC"
LINK = "1F5FBF"

_THIN = Side(style="thin", color=GRID)
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_FONT = "Calibri"

_EMAIL = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
_SHORT_COLS = {"#", "Founded", "Employees", "Online / Offline / Both",
               "Country Code / Region Code"}
_LINK_COLS = {"Website", "Office Website", "LinkedIn", "Email", "Office Email"}


def _url(col: str, value: str) -> str | None:
    """A safe hyperlink target for this cell, or None (never a broken link)."""
    v = value.strip()
    if not v or " " in v or len(v) > 1900:
        return None
    if col in ("Email", "Office Email"):
        return f"mailto:{v}" if _EMAIL.match(v) else None
    if v.lower().startswith(("http://", "https://")):
        return v
    if "." in v and "@" not in v:
        return "https://" + v
    return None


def _lines(text: str, width: float) -> int:
    """Wrapped line count of text in a column of this width (10 pt Calibri
    fits about 1.15 characters per width unit)."""
    per_line = max(4.0, width * 1.15)
    return sum(max(1, math.ceil(len(part) / per_line)) for part in str(text).split("\n"))


def polish_distributor_sheet(ws, header_row: int = 5) -> None:
    """The main market sheet: verified companies only, no colour flags (flagged
    companies are kept out of it and listed in flagged_for_review.xlsx)."""
    first = header_row + 1
    ncol = ws.max_column
    last = get_column_letter(ncol)
    header = [str(ws.cell(row=header_row, column=c).value or "").strip() for c in range(1, ncol + 1)]
    col_of = {name: i + 1 for i, name in enumerate(header)}
    widths = {c: (ws.column_dimensions[get_column_letter(c)].width or 14) for c in range(1, ncol + 1)}
    widths[1] = 6
    for c, w in widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w
    total_width = sum(widths.values())

    # --- rows of data ------------------------------------------------------
    data_rows = [r for r in range(first, ws.max_row + 1)
                 if str(ws.cell(row=r, column=col_of.get("Company", 2)).value or "").strip()]
    countries, acquired = set(), 0
    for n, r in enumerate(data_rows, start=1):
        band = PatternFill("solid", fgColor=ROW_ALT if n % 2 == 0 else "FFFFFF")
        height_lines = 1
        for c in range(1, ncol + 1):
            name = header[c - 1]
            cell = ws.cell(row=r, column=c)
            if name == "#":
                cell.value = n
            elif name == "Founded" and isinstance(cell.value, str) and re.fullmatch(r"\d{4}", cell.value.strip()):
                cell.value = int(cell.value.strip())
            text = "" if cell.value is None else str(cell.value)

            is_company = name == "Company"
            cell.fill = band
            cell.border = _BORDER
            cell.alignment = Alignment(
                horizontal="center" if name in _SHORT_COLS else "left",
                vertical="center", wrap_text=True, indent=0 if name in _SHORT_COLS else 1)
            cell.font = Font(name=_FONT, size=10, bold=is_company, color="1A1A1A")

            if name in _LINK_COLS and text:
                target = _url(name, text)
                if target:
                    cell.hyperlink = target
                    cell.font = Font(name=_FONT, size=10, color=LINK, underline="single")
            if text:
                height_lines = max(height_lines, _lines(text, widths[c]))
        if "Headquarters" in col_of:
            hq = str(ws.cell(row=r, column=col_of["Headquarters"]).value or "")
            if "," in hq:
                countries.add(hq.rsplit(",", 1)[1].strip().lower())
        acquired += "(acquired by" in str(ws.cell(row=r, column=col_of.get("Company", 2)).value or "").lower()
        ws.row_dimensions[r].height = min(130, max(30, height_lines * 13.5 + 6))

    # --- title banner, definition band, summary bar --------------------------
    for r in (1, 2, 3, 4):
        if f"A{r}:{last}{r}" not in {str(m) for m in ws.merged_cells.ranges}:
            ws.merge_cells(f"A{r}:{last}{r}")
    title = ws["A1"]
    title.font = Font(name=_FONT, size=16, bold=True, color="FFFFFF")
    title.fill = PatternFill("solid", fgColor=NAVY)
    title.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 38
    for r in (2, 3):
        cell = ws[f"A{r}"]
        cell.font = Font(name=_FONT, size=10, italic=True, color="33415C")
        cell.fill = PatternFill("solid", fgColor=BAND_BLUE)
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        ws.row_dimensions[r].height = max(30, _lines(cell.value or "", total_width * 0.95) * 14 + 10)
    stat = ws["A4"]
    stat.value = (f"Distributors listed: {len(data_rows)}     |     "
                  f"Headquarter countries: {len(countries)}     |     "
                  f"Acquired by another group (named in the company name): {acquired}")
    stat.font = Font(name=_FONT, size=10, bold=True, color=NAVY)
    stat.fill = PatternFill("solid", fgColor=STAT_BLUE)
    stat.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[4].height = 24

    # --- header -----------------------------------------------------------------
    gold = Side(style="medium", color=GOLD)
    for c in range(1, ncol + 1):
        cell = ws.cell(row=header_row, column=c)
        cell.font = Font(name=_FONT, size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(left=_THIN, right=_THIN, top=_THIN, bottom=gold)
    ws.row_dimensions[header_row].height = 34

    # --- sheet-level ---------------------------------------------------------------
    ws.freeze_panes = f"C{first}"
    ws.auto_filter.ref = f"A{header_row}:{last}{max(first, first + len(data_rows) - 1)}"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.sheet_properties.tabColor = NAVY
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{header_row}:{header_row}"
    ws.print_options.horizontalCentered = True


def polish_flag_sheet(ws) -> None:
    """flagged_for_review.xlsx: Market | Company | Website | Headquarters |
    AI Mode said | Why flagged. Company cell on an amber tint."""
    widths = {"Market": 36, "Company": 44, "Website": 30, "Headquarters": 30,
              "AI Mode said": 36, "Why flagged": 95}
    header = [str(ws.cell(row=1, column=c).value or "") for c in range(1, ws.max_column + 1)]
    for c, name in enumerate(header, start=1):
        ws.column_dimensions[get_column_letter(c)].width = widths.get(name, 20)
        cell = ws.cell(row=1, column=c)
        cell.font = Font(name=_FONT, size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(left=_THIN, right=_THIN, top=_THIN, bottom=Side(style="medium", color=GOLD))
    ws.row_dimensions[1].height = 30
    for r in range(2, ws.max_row + 1):
        lines = 1
        for c, name in enumerate(header, start=1):
            cell = ws.cell(row=r, column=c)
            text = "" if cell.value is None else str(cell.value)
            cell.fill = PatternFill("solid", fgColor=AMBER if name == "Company" else ("FFFFFF" if r % 2 else ROW_ALT))
            cell.border = _BORDER
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
            cell.font = Font(name=_FONT, size=10, bold=name == "Company", color="1A1A1A")
            if name == "Website" and text and _url(name, text):
                cell.hyperlink = _url(name, text)
                cell.font = Font(name=_FONT, size=10, color=LINK, underline="single")
            if text:
                lines = max(lines, _lines(text, widths.get(name, 20)))
        ws.row_dimensions[r].height = min(120, max(24, lines * 13.5 + 6))
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{max(2, ws.max_row)}"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.sheet_properties.tabColor = GOLD


def polish_report_sheet(ws) -> None:
    """status_check_report.xlsx: File | Company | Website | Status | Acquired
    by | Year | Evidence | Action. Neutral banding (no colour coding)."""
    ncol = ws.max_column
    last = get_column_letter(ncol)
    header = [str(ws.cell(row=1, column=c).value or "") for c in range(1, ncol + 1)]
    action_col = header.index("Action") + 1 if "Action" in header else None
    widths = {"File": 34, "Company": 42, "Website": 30, "Status": 11,
              "Acquired by": 32, "Year": 8, "Evidence": 95, "Action": 11}
    for c, name in enumerate(header, start=1):
        ws.column_dimensions[get_column_letter(c)].width = widths.get(name, 16)
        cell = ws.cell(row=1, column=c)
        cell.font = Font(name=_FONT, size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(left=_THIN, right=_THIN, top=_THIN, bottom=Side(style="medium", color=GOLD))
    ws.row_dimensions[1].height = 30
    for r in range(2, ws.max_row + 1):
        action = str(ws.cell(row=r, column=action_col).value or "") if action_col else ""
        # Neutral banding: colour is reserved for flagged-for-review companies.
        fill = PatternFill("solid", fgColor=ROW_ALT if r % 2 == 1 else "FFFFFF")
        lines = 1
        for c, name in enumerate(header, start=1):
            cell = ws.cell(row=r, column=c)
            text = "" if cell.value is None else str(cell.value)
            cell.fill, cell.border = fill, _BORDER
            cell.alignment = Alignment(
                horizontal="center" if name in ("Status", "Year", "Action") else "left",
                vertical="center", wrap_text=True, indent=0 if name in ("Status", "Year", "Action") else 1)
            cell.font = Font(name=_FONT, size=10, bold=name == "Company", color="1A1A1A")
            if name == "Website" and text and _url(name, text):
                cell.hyperlink = _url(name, text)
                cell.font = Font(name=_FONT, size=10, color=LINK, underline="single")
            if text:
                lines = max(lines, _lines(text, widths.get(name, 16)))
        ws.row_dimensions[r].height = min(120, max(24, lines * 13.5 + 6))
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{last}{max(2, ws.max_row)}"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.sheet_properties.tabColor = GOLD
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:1"
