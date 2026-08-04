"""Fix firewall FINAL Turnover / Scale → verified ranges; fill empties."""
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
    ROOT / "output/run4/global_firewall_market_global/global_firewall_market_FINAL.xlsx",
    ROOT / "output/run4/global_firewall_market_global/global_firewall_market_FINAL_320plus.xlsx",
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

# Agent-verified turnover ranges (filings / IR / CRN / LinkedIn scale)
VERIFIED: dict[str, str] = {
    # Global distributors
    "ingram micro": "$50–55B",
    "td synnex": "$55–65B",
    "arrow electronics": "$28–32B",
    "d&h distributing": "$5.5–7B",
    "also holding": "€9–12B",
    "westcon-comstor": "$1.8–2.5B",
    "exclusive networks": "€1.2–1.6B",
    "scan source": "$3–4B",
    "scansource": "$3–4B",
    "redington": "$8–12B",
    "intcomex": "$800M–1.5B",
    "dicker data": "A$2–3B",
    "synnex technology international": "$10–15B",
    "macnica": "$5–8B",
    "datatec": "$3.5–4B",
    "digital china": "$5–10B",
    # Resellers / VARs / MSPs
    "cdw": "$20–23B",
    "cdw government": "$7–9B",
    "shi international": "$14–16B",
    "insight enterprises": "$9–11B",
    "world wide technology": "$17–20B",
    "computacenter": "£7–9.5B",
    "softcat": "£1.2–1.5B",
    "connection": "$2–3B",
    "carahsoft technology": "$8–12B",
    "presidio": "$2.5–4B",
    "softchoice": "C$1.5–2.5B",
    "eplus technology": "$1.5–2.5B",
    "zones": "$2–3B",
    "pax8": "$800M–1.5B",
    "crayon": "$600M–1.2B",
    "softwareone": "$1–1.5B",
    "bechtle": "€6–7B",
    "bechtle systemhaus": "€2–4B",
    "atea": "€3.5–4.5B",
    "cancom": "€1.5–2.5B",
    "logicalis": "$1–1.5B",
    "noventiq": "$800M–1.5B",
    "econocom": "€2–3B",
    "scc": "£1.5–2.5B",
    "matrix it": "$1–2B",
    # Global SIs / telcos
    "accenture": "$65–70B",
    "deloitte": "$60–70B",
    "capgemini": "€20–23B",
    "cognizant": "$19–21B",
    "tata consultancy services": "$28–32B",
    "infosys": "$18–20B",
    "wipro": "$10–12B",
    "hcltech": "$12–14B",
    "tech mahindra": "$5.5–7B",
    "ltimindtree": "$4–5B",
    "dxc technology": "$12–15B",
    "ntt data": "$28–35B",
    "kyndryl": "$14–16B",
    "fujitsu personal / fujitsu japan services": "$20–30B",
    "cgi": "C$14–16B",
    "atos": "€9–12B",
    "bt group": "£19–21B",
    "booz allen hamilton": "$10–12B",
    "saic": "$7–8.5B",
    "general dynamics information technology": "$8–10B",
    "samsung sds": "$10–15B",
    "lg cns": "$3–5B",
    "fpt telecom / fpt is": "$2–4B",
    "asseco poland": "€3–4.5B",
    "sopra steria": "€5–6.5B",
    "stefanini": "$1–2B",
    "sonda": "$1–2B",
    "swisscom": "CHF10–12B",
    "telia company": "SEK80–100B",
    "orange business": "€5–8B",
    "vodafone business": "€8–12B",
    "singtel enterprise": "S$3–5B",
    "telstra purple": "A$2–4B",
    "t-systems": "€5–8B",
    "telefonica tech": "€1–2.5B",
    "ncs (singtel)": "$1–2B",
    "otsuka corporation": "$5–8B",
    "ctc (itochu techno-solutions)": "$5–8B",
    "beijing teamsun technology": "$1–3B",
    "pccw solutions": "$1–2B",
    "elm company": "$1–3B",
    "raya holding": "$500M–1B",
    "altron": "$1–1.5B",
    "business connexion (bcx)": "$500M–1.5B",
    "dimension data africa / ntt": "$5–10B",
    "eoh": "$200–500M",
    # Mid / previously wrong point values
    "appriver": "$80–120M",
    "bluestar": "$80–120M",
    "core bts": "$40–60M",
    "logicom": "$80–120M",
    "ma labs": "$80–120M",
    "sherweb": "$80–120M",
    "ras infotech": "$0.8–1.2M",
    "sisl infotech": "$175–220M",
    "cranel": "$8–12M",
    "dyntek services": "$8–12M",
    "netpoleon solutions": "$8–12M",
    "rahi systems": "$40–60M",
    "red river technology": "$8–12M",
    "spectrami": "$8–12M",
    "valuepoint systems": "$8–12M",
    "invitty": "$110–140K",
    "comsource": "$0.5–1M",
    "elcore group": "$0.5–1M",
}

# Employee-band → channel scale estimate when no verified figure
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
    """Exact / legal-suffix match only — never apply parent revenue to country units."""
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
    }


def clean_raw(raw: str) -> str:
    v = (raw or "").strip()
    v = re.sub(r"\s+", " ", v.replace("\n", " "))
    if re.fullmatch(r"<\s*\$?\s*1\s*M", v, re.I):
        return "$0.5–1M"
    m = re.search(r"(\d+(?:\.\d+)?)\s*million", v, re.I)
    if m and "revenue" in v.lower():
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
    ws = wb["Firewall Companies"] if "Firewall Companies" in wb.sheetnames else wb.active

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
        "dist": dist.most_common(25),
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
                if shown >= 30:
                    break
        print("  top values:")
        for v, n in rep["dist"]:
            print(f"    {n:3d}  {v}")
        print()


if __name__ == "__main__":
    main()
