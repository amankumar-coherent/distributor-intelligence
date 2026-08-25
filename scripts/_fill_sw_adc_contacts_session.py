#!/usr/bin/env python3
"""Apply agent web-verified Contact Person / Role / person LinkedIn to:
- Smartwatch FINAL
- ADC slim xlsx
Also applies prior KNOWN + WEB_HITS from ADC scripts.
"""
from __future__ import annotations

import json
import re
import shutil
import sys
from datetime import datetime, timezone
from pathlib import Path

import openpyxl

ROOT = Path(r"D:\vendor-intel")
sys.path.insert(0, str(ROOT / "scripts"))
import fill_adc_contact_role_linkedin as xls  # noqa: E402
from apply_adc_websearch_contacts import WEB_HITS  # noqa: E402

NPD = {
    "",
    "n/a",
    "na",
    "none",
    "-",
    "—",
    "not publicly disclosed",
    "unknown",
    "tbd",
    "null",
    "not publicly available",
    "full name or empty",
    "title or empty",
    "empty",
    "full name",
}


def emp(v: object) -> bool:
    if v is None:
        return True
    s = str(v).strip().lower()
    return s in NPD or s.startswith("not public")


def norm(c: object) -> str:
    return re.sub(r"\s+", " ", str(c or "").strip().lower())


# Agent web-verified smartwatch fills (person LinkedIn preferred)
SW: dict[str, tuple[str, str, str]] = {
    "digisanse": ("Shi Longxiang", "Administrador Único / Founder", ""),
    "distrihub": (
        "Marius Butauskas",
        "Founder",
        "https://www.linkedin.com/in/marius-butauskas-14800698",
    ),
    "duna electronics": (
        "Anett Karaszi",
        "Sales & Purchasing Manager",
        "https://www.linkedin.com/in/anett-karaszi-279653261",
    ),
    "fpt shop": (
        "Nguyen Viet Anh",
        "Executive Vice President / Deputy CEO (FPT Retail – FPT Shop)",
        "https://www.linkedin.com/in/viet-anh-nguyen-3a527945",
    ),
    "hoang ha mobile": (
        "Hoang Huu Huynh",
        "CEO / Tổng Giám đốc",
        "",
    ),
    "rec-line": (
        "Pedro Martínez García",
        "Founder & CEO",
        "https://www.linkedin.com/in/pedro-mart%C3%ADnez-garc%C3%ADa-4384a178",
    ),
    "the gioi di dong": (
        "Vu Dang Linh",
        "CEO (Mobile World Investment Corporation)",
        "",
    ),
    "vivid telecommunication": (
        "Nelson Leong",
        "Managing Director",
        "https://www.linkedin.com/in/nelson-leong-780a13321",
    ),
    "ss solution": (
        "Ko Chun Ming",
        "Director",
        "https://www.linkedin.com/in/ko-chun-ming-daniel-b86913151",
    ),
    "lulu hypermarket": (
        "Yusuff Ali M.A.",
        "Chairman & Managing Director (LuLu Group)",
        "https://www.linkedin.com/in/yusuffalima",
    ),
    "westcon-comstor": (
        "David Grant",
        "Chief Executive Officer",
        "https://www.linkedin.com/in/david-grant-43a17913",
    ),
    "fnac": (
        "Enrique Martinez",
        "Group CEO (Fnac Darty)",
        "https://www.linkedin.com/in/enrique-martinez-7a593618",
    ),
    "notebooksbilliger": (
        "Oliver Hellmold",
        "Vorstandsvorsitzender / CEO",
        "https://www.linkedin.com/in/oliver-hellmold-5b3a9264",
    ),
    "solutions 2 go": (
        "Gabrielle Chevalier",
        "CEO / Co-Founder",
        "https://www.linkedin.com/in/gabrielle-chevalier-857ab1110",
    ),
    "digimap": (
        "Ranty N.",
        "Strategic Leader / Retail Head",
        "https://www.linkedin.com/in/ranty-n-bb06a32b",
    ),
    "costco": (
        "Ron Vachris",
        "President and Chief Executive Officer",
        "",
    ),
    "the good guys commercial": (
        "Nick Wells",
        "Group CEO (JB Hi-Fi / The Good Guys)",
        "",
    ),
    "a.r.m. store sdn bhd": (
        "PH Sim",
        "Director (A.R.M. Century / related)",
        "https://www.linkedin.com/in/ph-sim-75905570",
    ),
}

# Prefer exact company match keys as stored (lowercase)
SW_EXTRA_ALIASES = {
    "a.r.m. store sdn bhd": "a.r.m. store sdn bhd",
}


# Agent web-verified ADC fills this session
ADC_EXTRA: dict[str, tuple[str, str, str]] = {
    "Funakoshi": (
        "Tetsuya Ikeda",
        "President and CEO",
        "https://www.linkedin.com/in/ikeda-tetsuya-a855bb9b",
    ),
    "Euroclone S.p.A.": (
        "Franco Aiolfi",
        "Founder / CEO / Presidente",
        "",
    ),
    "Farmexpert": (
        "Marius Tanase",
        "Executive Director",
        "https://www.linkedin.com/in/tanase-marius-38a7359",
    ),
    "Farmacol": (
        "Tomasz Rogóż",
        "Member of the Management Board / Development Director (ex-President)",
        "https://www.linkedin.com/in/tomasz-rog%C3%B3%C5%BC-mba-5a1621163",
    ),
    "Difare Ecuador": (
        "Carlos Cueva Mejía",
        "Executive President (Grupo Difare)",
        "",
    ),
    "Dianova GmbH": (
        "Jürgen A. Frerichs",
        "Geschäftsführer (historical / registry)",
        "",
    ),
    "Anda Inc.": (
        "Jason Hollar",
        "CEO (Cardinal Health)",
        "https://www.linkedin.com/in/jasonmhollar",
    ),
    "Anda Pharmaceuticals": (
        "Jason Hollar",
        "CEO (Cardinal Health)",
        "https://www.linkedin.com/in/jasonmhollar",
    ),
    "AAH Pharmaceuticals": (
        "Mark Dayan",
        "CEO (AAH / McKesson UK context)",
        "",
    ),
    "Alfresa Holdings": ("Taiji Ito", "President", ""),
    "ASD Healthcare": (
        "Lisa Thompson Harrison",
        "SVP & President, Specialty Distribution & Solutions (Cencora)",
        "https://www.linkedin.com/in/lisa-thompson-harrison-ba066565",
    ),
}


def apply_to_sheet(
    path: Path,
    sheet_name: str | None,
    header_row: int,
    updates: dict[str, tuple[str, str, str]],
    col_map: dict[str, str],
    force_name: bool = False,
) -> dict:
    bak = path.with_suffix(path.suffix + f".bak_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
    shutil.copy2(path, bak)
    wb = openpyxl.load_workbook(path)
    ws = wb[sheet_name] if sheet_name else wb.active
    headers = [c.value for c in ws[header_row]]
    ic = {str(h): i + 1 for i, h in enumerate(headers) if h}
    for need in ("Company", "Contact Person", "Role", "LinkedIn"):
        if need not in ic:
            raise SystemExit(f"Missing column {need} in {path.name}")

    applied = 0
    details = []
    for row in range(header_row + 1, ws.max_row + 1):
        co = ws.cell(row, ic["Company"]).value
        if not co:
            continue
        key = norm(co)
        hit = updates.get(key)
        if not hit:
            continue
        p, role, li = hit
        cur_cp = ws.cell(row, ic["Contact Person"]).value
        cur_role = ws.cell(row, ic["Role"]).value
        cur_li = ws.cell(row, ic["LinkedIn"]).value
        changed = False
        # replace placeholders / empty / bad fragments
        bad_name = emp(cur_cp) or str(cur_cp).strip().lower() in {
            "senior vice",
            "leadership team",
            "full name",
        }
        if (bad_name or force_name) and p and not emp(p):
            ws.cell(row, ic["Contact Person"]).value = p
            changed = True
        bad_role = emp(cur_role) or str(cur_role).strip().lower() in {"president", "title or empty"}
        if (bad_role or force_name) and role and not emp(role):
            # only overwrite weak role when name was weak or role empty
            if bad_role or force_name:
                ws.cell(row, ic["Role"]).value = role
                changed = True
        if li and "/in/" in li:
            if emp(cur_li) or "linkedin.com/company" in str(cur_li or "").lower() or force_name:
                ws.cell(row, ic["LinkedIn"]).value = li
                changed = True
        elif li and emp(cur_li):
            ws.cell(row, ic["LinkedIn"]).value = li
            changed = True
        if changed:
            applied += 1
            details.append({"company": co, "person": p, "role": role, "linkedin": li})
    wb.save(path)
    wb.close()
    return {"backup": str(bak), "applied": applied, "details": details}


def build_adc_updates() -> dict[str, tuple[str, str, str]]:
    out: dict[str, tuple[str, str, str]] = {}
    for k, v in xls.KNOWN.items():
        # skip empty person LI only entries with empty person? keep all with person
        if v[0] and not emp(v[0]) and v[0].lower() != "leadership team":
            out[norm(k)] = v
    for k, v in WEB_HITS.items():
        if v[0] and not emp(v[0]) and "leadership" not in v[0].lower():
            out[norm(k)] = v
    for k, v in ADC_EXTRA.items():
        if v[0] and not emp(v[0]):
            out[norm(k)] = v
    return out


def recount(path: Path, header_row: int) -> dict:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    hdr = rows[header_row - 1]
    ic = {str(h): i for i, h in enumerate(hdr) if h}
    n = r = li = 0
    need = 0
    for row in rows[header_row:]:
        if not row[ic["Company"]]:
            continue
        if not emp(row[ic["Contact Person"]]):
            n += 1
        if not emp(row[ic["Role"]]):
            r += 1
        if row[ic["LinkedIn"]] and "linkedin.com/in/" in str(row[ic["LinkedIn"]]).lower():
            li += 1
        if emp(row[ic["Contact Person"]]) or emp(row[ic["Role"]]) or emp(row[ic["LinkedIn"]]):
            need += 1
    total = len(rows) - header_row
    wb.close()
    return {"names": n, "roles": r, "person_li": li, "still_need_any": need, "total": total}


def main() -> None:
    sw_path = ROOT / "output/run4/global_smartwatch_market_global/global_smartwatch_market_FINAL.xlsx"
    adc_path = (
        ROOT
        / "output/web_expand/global_antibody_drug_conjugates_market_global/global_antibody_drug_conjugates_market.xlsx"
    )

    sw_res = apply_to_sheet(
        sw_path,
        "Smartwatch Companies",
        2,
        SW,
        {},
        force_name=False,
    )
    # also fill LinkedIn-only upgrades for rows that already have good names:
    # re-run with force on LinkedIn only by second pass handled inside apply

    adc_res = apply_to_sheet(
        adc_path,
        None,
        2,
        build_adc_updates(),
        {},
        force_name=False,
    )

    summary = {
        "ts": datetime.now(timezone.utc).isoformat(),
        "smartwatch": {**sw_res, "stats": recount(sw_path, 2)},
        "adc": {**adc_res, "stats": recount(adc_path, 2)},
    }
    out = ROOT / "output/run4/global_smartwatch_market_global/contact_fill_session_summary.json"
    out.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    adc_sum = (
        ROOT
        / "output/web_expand/global_antibody_drug_conjugates_market_global/contact_fill_session_summary.json"
    )
    adc_sum.write_text(json.dumps(summary["adc"], indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps(summary, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
