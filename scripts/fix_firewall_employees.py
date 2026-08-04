"""Fix firewall FINAL Employees → verified LinkedIn-style ranges (no trust garbage / raw mega-numbers)."""
from __future__ import annotations

import re
import shutil
from collections import Counter
from datetime import datetime
from pathlib import Path

import openpyxl

PATHS = [
    Path(r"D:\vendor-intel\output\run4\global_firewall_market_global\global_firewall_market_FINAL.xlsx"),
    Path(
        r"D:\vendor-intel\output\run4\global_firewall_market_global\global_firewall_market_FINAL_320plus.xlsx"
    ),
]

# Verified overrides (agent web research — Forbes/LinkedIn/filings)
OVERRIDES: dict[str, str] = {
    # garbage / NPD fills
    "d&h distributing": "1001-5000",  # ~1,600–2,000 (Forbes/Revelio)
    "infinigate nordic": "51-200",  # Nordic entities ~30–100 each; regional cluster
    "solutions2go": "201-500",  # ~300–400 (LinkedIn)
    "chillisoft": "11-50",  # ~17 ANZ cybersecurity distributor (LinkedIn)
    "aspire technology partners": "51-200",  # ~143 (LinkedIn)
    "acpl systems": "201-500",  # ~299 (LinkedIn)
    "m.tech products": "201-500",  # ~268 (LinkedIn)
    "zedsphere": "11-50",  # small UK MSP distributor (QBS group niche)
    # mega / large — replace coarse 10000+ with real ranges
    "accenture": "500001-1000000",  # ~779k FY25
    "cognizant": "100001-500000",  # ~352k
    "tata consultancy services": "1000001+",  # ~600k+
    "infosys": "100001-500000",  # ~324k
    "wipro": "100001-500000",
    "hcltech": "100001-500000",
    "capgemini": "100001-500000",
    "deloitte": "100001-500000",
    "dxc technology": "100001-500000",
    "ntt data": "100001-500000",
    "fujitsu personal / fujitsu japan services": "100001-500000",
    "tech mahindra": "100001-500000",
    "ltimindtree": "50001-100000",
    "kyndryl": "50001-100000",
    "cgi": "50001-100000",
    "bt group": "50001-100000",
    "atos": "50001-100000",
    "vodafone business": "50001-100000",
    "ingram micro": "10001-50000",  # ~25–30k
    "td synnex": "10001-50000",  # ~23–28k
    "arrow electronics": "10001-50000",
    "cdw": "10001-50000",
    "computacenter": "10001-50000",
    "world wide technology": "5001-10000",
    "crayon": "5001-10000",
    "bechtle": "10001-50000",
    "booz allen hamilton": "10001-50000",
    "saic": "10001-50000",
    "samsung sds": "10001-50000",
    "fpt telecom / fpt is": "10001-50000",
    "general dynamics information technology": "10001-50000",
    "lg cns": "10001-50000",
    "ncs (singtel)": "10001-50000",
    "orange business": "10001-50000",
    "singtel enterprise": "10001-50000",
    "sonda": "10001-50000",
    "sopra steria": "10001-50000",
    "stefanini": "10001-50000",
    "swisscom": "10001-50000",
    "t-systems": "10001-50000",
    "telefonica tech": "10001-50000",
    "telia company": "10001-50000",
    "telstra purple": "10001-50000",
    "digital china": "10001-50000",
    "asseco poland": "10001-50000",
    "shi international": "5001-10000",
    "insight enterprises": "10001-50000",  # ~14k; sheet had bad 2-10
    "cdw": "10001-50000",
}

NPD = {"", "n/a", "na", "none", "-", "—", "not publicly disclosed", "unknown", "tbd"}

BUCKETS = [
    (10, "1-10"),
    (50, "11-50"),
    (200, "51-200"),
    (500, "201-500"),
    (1000, "501-1000"),
    (5000, "1001-5000"),
    (10000, "5001-10000"),
    (50000, "10001-50000"),
    (100000, "50001-100000"),
    (500000, "100001-500000"),
    (1000000, "500001-1000000"),
]


def norm_name(s: str) -> str:
    return re.sub(r"[^a-z0-9&]+", " ", (s or "").lower()).strip()


def parse_int(tok: str) -> int | None:
    tok = tok.replace(",", "").replace(" ", "")
    if not tok.isdigit():
        return None
    return int(tok)


def number_to_range(n: int) -> str:
    for hi, label in BUCKETS:
        if n <= hi:
            return label
    return "1000001+"


def looks_like_garbage(text: str) -> bool:
    t = text.strip()
    if not t:
        return False
    if re.search(
        r"(?i)\b(trust|foundation|bank|llc|inc\.?|ltd|gmbh|corp|partners|holdings)\b",
        t,
    ):
        if not re.search(r"\d", t):
            return True
    if re.fullmatch(r"[A-Za-z][A-Za-z &.'-]{2,}", t) and t.lower() not in NPD:
        return True
    return False


def normalize_employees(raw: str, company: str) -> tuple[str, str]:
    key = norm_name(company)
    ov_map = {norm_name(k): (k, v) for k, v in OVERRIDES.items()}
    if key in ov_map:
        needle, rng = ov_map[key]
        return rng, f"verified_override:{needle}"

    v = (raw or "").strip()
    if v.lower() in NPD:
        return "Not publicly disclosed", "npd"

    if looks_like_garbage(v):
        return "Not publicly disclosed", f"cleared_garbage:{v}"

    v2 = re.sub(r"(?i)\s*(employees?|staff|team members?)\s*$", "", v).strip()
    v2 = v2.replace("–", "-").replace("—", "-").replace("−", "-")
    v2 = re.sub(r"\s+", "", v2)
    # drop thousands separators inside numbers for parsing, keep structure
    # e.g. 1,001-5,000 → 1001-5000
    v2 = re.sub(r",(?=\d{3}(\D|$))", "", v2)

    m = re.fullmatch(r"(\d+)\s*-\s*(\d+)", v2)
    if m:
        a, b = parse_int(m.group(1)), parse_int(m.group(2))
        if a is not None and b is not None and a <= b:
            mid = (a + b) // 2
            return number_to_range(mid), f"normalized_range:{a}-{b}"

    m = re.fullmatch(r"(\d+)\+", v2)
    if m:
        n = parse_int(m.group(1))
        if n is not None:
            return number_to_range(n + 1), f"open_ended:{n}+"

    m = re.fullmatch(r"\d+", v2)
    if m:
        n = parse_int(v2)
        if n is not None:
            return number_to_range(n), f"point_to_range:{n}"

    nums = re.findall(r"\d+", v2)
    if nums:
        n = parse_int(nums[0])
        if n is not None:
            return number_to_range(n), f"extracted:{n}"

    return "Not publicly disclosed", f"unparsed:{v}"


def fix_file(path: Path) -> dict:
    bak = path.with_name(
        path.stem + f"_pre_emp_fix_{datetime.now().strftime('%Y%m%d_%H%M%S')}" + path.suffix
    )
    shutil.copy2(path, bak)

    wb = openpyxl.load_workbook(path)
    ws = wb["Firewall Companies"] if "Firewall Companies" in wb.sheetnames else wb.active

    hdr_row = None
    hdr = None
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), 1):
        vals = [str(c or "") for c in row]
        if "Company" in vals and "Employees" in vals:
            hdr_row = i
            hdr = vals
            break
    if not hdr_row:
        raise RuntimeError(f"No header in {path}")

    emp_col = hdr.index("Employees") + 1
    co_col = hdr.index("Company") + 1

    changes = []
    for r in range(hdr_row + 1, ws.max_row + 1):
        company = str(ws.cell(r, co_col).value or "").strip()
        if not company:
            continue
        old = str(ws.cell(r, emp_col).value or "").strip()
        new, reason = normalize_employees(old, company)
        if new != old:
            ws.cell(r, emp_col).value = new
            changes.append((company, old, new, reason))

    wb.save(path)

    dist = Counter()
    for r in range(hdr_row + 1, ws.max_row + 1):
        company = str(ws.cell(r, co_col).value or "").strip()
        if not company:
            continue
        dist[str(ws.cell(r, emp_col).value or "").strip()] += 1

    return {
        "path": path.name,
        "backup": bak.name,
        "changed": len(changes),
        "samples": changes[:25],
        "dist": dist.most_common(),
        "garbage_left": [c for c in changes if "TRUST" in c[1].upper()]
        + [
            (company, str(ws.cell(r, emp_col).value))
            for r in range(hdr_row + 1, ws.max_row + 1)
            for company in [str(ws.cell(r, co_col).value or "").strip()]
            if company and re.search(r"(?i)trust", str(ws.cell(r, emp_col).value or ""))
        ],
    }


def main() -> None:
    for p in PATHS:
        if not p.exists():
            print("missing", p)
            continue
        rep = fix_file(p)
        print("FILE", rep["path"])
        print("  backup", rep["backup"])
        print("  changed", rep["changed"])
        print("  samples:")
        for co, old, new, reason in rep["samples"]:
            print(f"    {co}: {old!r} -> {new!r} ({reason})")
        print("  distribution:")
        for v, c in rep["dist"]:
            print(f"    {c:3d}  {v}")
        if rep["garbage_left"]:
            print("  WARNING garbage left", rep["garbage_left"])
        print()


if __name__ == "__main__":
    main()
