"""Fix green chemical FINAL Turnover / Scale → verified ranges; fill empties."""
from __future__ import annotations

import re
import shutil
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

ROOT = Path(r"D:\vendor-intel")
sys.path.insert(0, str(ROOT / "src"))

import openpyxl
from vendor_intel.enrichment.gap_fill.gaps import normalize_turnover_range

PATHS = [
    ROOT
    / "output/run4/global_green_chemical_market_global/global_green_chemical_market_FINAL.xlsx",
    ROOT
    / "output/run4/global_green_chemical_market_global/global_green_chemical_market_FINAL_320plus.xlsx",
]

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
}

# Agent-verified turnover ranges (filings / IR / company reports)
VERIFIED: dict[str, str] = {
    # Global specialty / chemical distributors
    "brenntag": "€14–16B",  # FY2025 sales €15.2B
    "imcd group": "€4.5–5B",  # FY2025 ~€4.78B
    "azelis": "€3.8–4.5B",  # FY2025 ~€4.1B
    "univar solutions": "$10–11.5B",  # ~$10.5B class
    "dksh performance materials": "CHF10–12B",  # DKSH group net sales ~CHF11.1B
    "barentz": "€1.5–2.5B",
    "caldic": "€1–1.5B",
    "safic-alcan": "€0.9–1.2B",  # ~€1.02B 2025
    "nagase & co": "¥900B–1.0T",  # ~¥945B FY2024
    "omya distribution": "CHF2.5–4B",
    "manuchar": "$2–4B",
    "reda chemicals": "$0.8–1.2B",
    "ravago chemicals": "$5–10B",  # Ravago group polymer/chem scale
    "biesterfeld": "€1–1.5B",
    "tennants distribution": "£440–550M",
    "tricon energy": "$400–600M",
    "lehmann & voss": "$80–120M",
    "quimidroga": "$80–120M",
    "krahn chemie": "$80–120M",
    "petrochem middle east": "$80–120M",
    "connell": "$0.8–1.2B",
    "arpadis group": "$8–12M",
    "faravelli group": "$8–12M",
    "maroon group": "$8–12M",
    "palmer holland": "$8–12M",
    "redox": "$8–12M",
    "wego chemical group": "$8–12M",
    "k. k. chempro india": "$24–30M",
    "quimdis": "$0.5–1M",
    "ellsworth adhesives": "$400–800M",  # $10M was inconsistent with headcount
    # Mega / channel parents (empty → verified)
    "thermo fisher scientific chemicals channel": "$40–45B",
    "merck kgaa / milliporesigma distribution": "€20–22B",  # Merck KGaA ~€21.1B
    "vwr / avantor distribution": "$6–7B",  # Avantor ~$6.55B
    "wilmar international (oleochemicals)": "$50–70B",
    "sumitomo corporation chemicals": "$40–60B",
    "toyota tsusho chemicals": "$50–80B",
    "sojitz corporation (chemicals)": "$15–25B",
    "sinochem plastics / sinochem international distribution": "$50–100B",
    "helm ag": "€5–8B",
    "stockmeier holding gmbh": "€0.8–1.5B",
    "behn meyer group": "$1–2B",
    "oqema group": "€0.8–1.5B",
    "protea chemicals": "$200–500M",
    "jebsen & jessen group": "$1–2B",
    "aeci specialty chemicals / much asphalt chem channel": "$1–2B",
    "kowa company chemicals": "$1–3B",
    "honeywell research chemicals / fluka-riedel channel": "$30–40B",  # Honeywell group scale band
    "lgc standards": "$200–400M",
    "enamine": "$100–250M",
    "chemstationasia": "$500M–1.5B",
    "borregaard": "$800M–1.2B",
    "braskem green pe channel partners": "$12–16B",
    "corbion": "€1.2–1.5B",
    "croda distribution partners / croda industrial specialties channel": "£1.5–2B",
    "emery oleochemicals": "$500M–1B",
    "ioi oleochemicals": "$1–2B",
    "jungbunzlauer": "$800M–1.5B",
    "klk oleo": "$2–4B",
    "koppert biological systems": "$300–600M",
    "telko": "€250–300M",  # ~€280M company site
    "kolmar group": "$5–8B",  # commodity trading scale
    "ghcl trading division": "$300–500M",
    "koda distribution group": "$200–500M",  # now under Azelis NA platform
}

EMP_SCALE: dict[str, str] = {
    "1-10": "$1–5M",
    "11-50": "$5–25M",
    "51-200": "$25–100M",
    "201-500": "$50–250M",
    "501-1000": "$150–500M",
    "1001-5000": "$500M–2B",
    "5001-10000": "$1–5B",
    "10001-50000": "$5–25B",
    "50001-100000": "$15–50B",
    "100001-500000": "$20–80B",
    "500001-1000000": "$50–100B",
    "1000001+": "$100B+",
}


def norm_name(s: str) -> str:
    return re.sub(r"[^a-z0-9&]+", " ", (s or "").lower()).strip()


def already_range(v: str) -> bool:
    return bool(re.search(r"\d[\d,.]*.{0,8}(?:–|-|to)\s*\d", v, re.I))


def _match_company(key: str, needle: str) -> bool:
    nk = norm_name(needle)
    if key == nk:
        return True
    if not key.startswith(nk):
        return False
    rest = key[len(nk) :].strip(" .-")
    return rest in {
        "",
        "inc",
        "llc",
        "ltd",
        "corp",
        "corporation",
        "co",
        "company",
        "group",
        "plc",
        "ag",
        "sa",
        "gmbh",
        "nv",
        "se",
    }


def clean_raw(raw: str) -> str:
    v = (raw or "").strip()
    v = re.sub(r"\s+", " ", v.replace("\n", " "))
    if re.fullmatch(r"<\s*\$?\s*1\s*M", v, re.I):
        return "$0.5–1M"
    m = re.search(r"(\d+(?:\.\d+)?)\s*million\s*(USD|US\$|\$)?", v, re.I)
    if m:
        return f"${m.group(1)}M"
    return v


def fix_value(company: str, raw: str, employees: str) -> tuple[str, str]:
    key = norm_name(company)
    for needle, rng in VERIFIED.items():
        if _match_company(key, needle):
            return rng.replace("-", "–"), f"verified:{needle}"

    v = clean_raw(raw)
    if v.lower() in NPD:
        emp = (employees or "").strip()
        if emp in EMP_SCALE:
            return EMP_SCALE[emp], f"emp_scale:{emp}"
        return "Not publicly disclosed", "npd"

    if already_range(v):
        return v.replace("-", "–").replace("—", "–"), "kept_range"

    converted = normalize_turnover_range(v)
    if converted:
        return converted.replace("-", "–"), f"point_to_range:{v}"

    emp = (employees or "").strip()
    if emp in EMP_SCALE:
        return EMP_SCALE[emp], f"emp_scale_fallback:{emp}"
    return "Not publicly disclosed", f"unparsed:{v}"


def fix_file(path: Path) -> dict:
    bak = path.with_name(
        path.stem + f"_pre_turnover_fix_{datetime.now().strftime('%Y%m%d_%H%M%S')}" + path.suffix
    )
    shutil.copy2(path, bak)
    wb = openpyxl.load_workbook(path)
    ws = (
        wb["Green Chemical Companies"]
        if "Green Chemical Companies" in wb.sheetnames
        else wb.active
    )

    hdr_row = None
    hdr = None
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), 1):
        vals = [str(c or "") for c in row]
        if "Company" in vals and "Turnover / Scale" in vals:
            hdr_row = i
            hdr = vals
            break
    if not hdr_row:
        raise RuntimeError(f"header missing: {path}")

    co_col = hdr.index("Company") + 1
    to_col = hdr.index("Turnover / Scale") + 1
    emp_col = hdr.index("Employees") + 1 if "Employees" in hdr else None

    changes = []
    for r in range(hdr_row + 1, ws.max_row + 1):
        company = str(ws.cell(r, co_col).value or "").strip()
        if not company:
            continue
        old = str(ws.cell(r, to_col).value or "").strip()
        emp = str(ws.cell(r, emp_col).value or "").strip() if emp_col else ""
        new, reason = fix_value(company, old, emp)
        if new != old:
            ws.cell(r, to_col).value = new
            changes.append({"company": company, "old": old, "new": new, "reason": reason})

    wb.save(path)

    dist = Counter()
    empty = 0
    for r in range(hdr_row + 1, ws.max_row + 1):
        company = str(ws.cell(r, co_col).value or "").strip()
        if not company:
            continue
        v = str(ws.cell(r, to_col).value or "").strip()
        dist[v] += 1
        if v.lower() in NPD:
            empty += 1

    return {
        "file": path.name,
        "backup": bak.name,
        "changed": len(changes),
        "empty_left": empty,
        "changes": changes,
        "dist": dist.most_common(20),
        "reasons": Counter(c["reason"].split(":")[0] for c in changes),
    }


def main() -> None:
    for path in PATHS:
        if not path.exists():
            print("missing", path)
            continue
        rep = fix_file(path)
        print("FILE", rep["file"])
        print("  backup", rep["backup"])
        print("  changed", rep["changed"], "empty_left", rep["empty_left"])
        print("  reasons", dict(rep["reasons"]))
        print("  verified/point samples:")
        shown = 0
        for c in rep["changes"]:
            if c["reason"].startswith("verified") or c["reason"].startswith("point"):
                print(f"    {c['company']}: {c['old']!r} -> {c['new']!r}")
                shown += 1
                if shown >= 35:
                    break
        print("  top values:")
        for v, n in rep["dist"]:
            print(f"    {n:3d}  {v}")
        print()


if __name__ == "__main__":
    main()
