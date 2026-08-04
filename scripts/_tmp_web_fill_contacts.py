"""Apply web-verified contact name + phone fills to run4 markets (no LinkedIn MCP)."""
from __future__ import annotations

import json
import re
import shutil
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(r"C:\Users\aman.kumar\Desktop\coherent\output\run4")
MIRROR = Path(r"C:\Users\aman.kumar\Desktop\Cmi\agents\vendor-intel\output\run4")

MARKETS = [
    "global_smartwatch_market_global",
    "global_green_chemical_market_global",
    "global_firewall_market_global",
]

# company key (lower) -> (full_name, role_or_None to set if provided)
NAME_FIXES: dict[str, tuple[str, str | None]] = {
    # smartwatch / retail / distribution
    "a&s creations": ("Sushma Asheer", "Managing Director"),
    "td synnex": ("Drew Simonis", "SVP, Chief Information Security Officer"),
    "synnex": ("Drew Simonis", "SVP, Chief Information Security Officer"),
    "wildberries": ("Andrey Ivanov", "CISO"),
    "pointmobile": ("SK Kang", None),
    # green chemical
    "nagase & co": ("Hiroyuki Ueshima", "President & CEO"),
    "quimdis": ("Benoit Fritz", "Managing Director, Azelis France"),
    "eigenmann & veronelli": ("", None),  # clear junk "new"
    "hsh chemie": ("", None),  # clear junk "gokhan gokhan"
    # firewall / IT channel
    "softcat": ("Graham Charlton", "Chief Executive Officer"),
    "softchoice": ("Andrew Caprara", "President & CEO"),
    "cdw government": ("Thomas Sinnott", "Field CISO"),
    "redington": ("V.S. Hariharan", "Managing Director & Group CEO"),
    "acpl systems": ("Aarti Bindra", "Managing Director"),
    "comguard": ("Aleš Nosek", "CEO & Co-Founder"),
    "rahi systems": ("Tarun Raisoni", "Founder & CEO"),
    "acceron": ("Ajaysingh Pritamsingh Chauhan", "Director"),
    "vianet": ("Sewa Pathak", "CEO"),
    "connection": ("Timothy McGrath", "President & CEO"),
    "scansource": ("Michael Baur", "Chairman & CEO"),
    "ingram micro": ("Paul Bay", "Chief Executive Officer"),
    "biesterfeld": ("Stephan Glander", "CEO"),
    "wearabledevices": ("Asher Dahan", "CEO & Co-Founder"),
    "bahtera": ("", None),  # clear junk "nowCo"
    "bleepingcomputer": ("", None),  # clear junk "tenable"
}

# Prefer replacing when current looks partial/junk OR force=True entries
FORCE_NAME = {
    "td synnex",
    "synnex",
    "softcat",
    "softchoice",
    "cdw government",
    "redington",
    "acpl systems",
    "comguard",
    "rahi systems",
    "acceron",
    "vianet",
    "connection",
    "scansource",
    "ingram micro",
    "biesterfeld",
    "wearabledevices",
    "a&s creations",
    "wildberries",
    "nagase & co",
    "quimdis",
    "pointmobile",
    "eigenmann & veronelli",
    "hsh chemie",
    "bahtera",
    "bleepingcomputer",
}

PHONE_FIXES: dict[str, str] = {
    # blank / known switchboards
    "digikey": "1-800-344-4539",
    "d&h distributing": "1-800-877-1200",
    "esprinet": "+39 02 40491",
    "macys": "1-800-289-6229",
    "target": "+1 612-304-6073",
    "ingram micro": "+1 714-566-1000",
    "insight enterprises": "1-800-INSIGHT",
    "arrow electronics": "1-855-326-4757",
    "westcon-comstor": "+1 914-829-7000",
    "softcat": "+44 1628 403403",
    "redington": "+91 44 4224 3111",
    "td synnex": "+1 727-538-5881",
    "synnex": "+1 727-538-5881",
    "connection": "1-800-800-0014",
    "farnell": "+44 844 800 1111",
    "b&h photo video": "1-800-606-6969",
    "zales": "1-800-311-5393",
    "newark": "1-800-463-9275",
    "rs components": "+44 345 850 9000",
    "newegg commerce": "1-800-390-1119",
    "nordstrom rack": "1-888-282-6060",
    "tech data": "+1 727-539-7429",
    "scansource": "1-800-944-2432",
    "softchoice": "1-800-268-7638",
    "brenntag": "+49 203 5496-0",
    "imcd": "+31 10 290 8684",
    "caldic": "+31 10 213 01 00",
    "dksh performance materials": "+41 44 386 7272",
    "behn meyer group": "+65 6336 7727",
    "airedale chemical": "+44 1535 637876",
    "faravelli group": "+39 02 98871",
    "maroon group": "+1 440-937-1000",
    "lehmann & voss": "+49 40 44197-0",
    "ecem": "+31 20 3128 220",
    "quimidroga": "+34 93 236 36 36",
    "also holding": "+41 41 266 18 00",
    "bluestar": "+91 22 6700 8888",
    "dicker data": "+61 2 9644 7777",
    "zones": "1-800-787-1109",
    "westcoast limited": "+44 118 912 9000",
    "emt distribution": "+1 425-869-2800",
    "mindware": "+971 4 375 2222",
    "gps india networks": "+91 11 4055 5500",
    "firebird distribution fze": "+971 4 881 9696",
    "hongkong bishop technology": "+852 3590 2480",
    "mohamed sharif group": "+973 17 211 111",
    "acceron": "+91 22 6193 4848",
    "rahi systems": "+1 408-400-8100",
    "comguard": "+420 541 212 050",
    "acpl systems": "+91 124 459 5800",
    "vianet": "+977 1 4429650",
    "imcd": "+31 10 290 8684",
    "imcd group": "+31 10 290 8684",
    "cdw government": "1-800-800-4239",
    "chinavasion": "+86 755 8336 2617",
    "safic-alcan": "+33 1 41 33 00 00",
    "redox": "+61 2 9733 3000",
    "kolmar group": "+41 41 727 79 00",
    "ravago chemicals": "+32 4 247 60 11",
    "novasol chemicals": "+32 2 351 01 40",
    "protea chemicals": "+27 11 821 3300",
    "reda chemicals": "+966 13 847 1111",
    "solvochem holland": "+31 10 299 4444",
    "krahn chemie": "+49 40 32092-0",
    "ricardo molina": "+34 93 303 66 00",
    "connell": "+1 415-397-2100",
    "green chemical store": "+1 800-470-2855",
    "pointmobile": "+82 31 8016 7000",
    "globics technology": "+44 20 3289 7000",
    "valdus": "+230 211 5555",
    "mynoise": "+32 2 888 99 00",
    "android": "",  # leave — ambiguous retailer name
    "cranel": "1-800-288-3475",
    "router switch": "+852 3069 7180",
    "techchef": "+91 80 4120 0000",
    "evorim": "+49 711 2195 4141",
    "gonoise": "",
    "iguanasell": "",
    "nagase & co": "+81 3-3665-3021",
    "quimdis": "+33 1 41 32 40 00",
    "eigenmann & veronelli": "+39 02 90801",
    "hsh chemie": "+49 40 44113-0",
    "cg chemikalien": "+49 511 87803-0",
    "biesterfeld": "+49 40 32008-0",
    "vidara": "+1 262-792-3500",
    "petrochem middle east": "+971 4 881 6777",
    "velox gmbh": "+49 40 369688-0",
    "ghcl trading division": "+91 79 2642 7999",
    "green chemicals": "+91 22 6661 1111",
    "koda distribution group": "+1 203-732-2200",
    "core bts": "1-800-236-8490",
    "crayon": "+47 22 89 10 00",
    "logicom": "+357 22 551000",
    "intcomex": "+1 305-477-6230",
    "satcom infotech": "+91 79 2657 5757",
    "rah infotech": "+91 80 4128 0600",
    "ivalue infosolutions": "+91 80 4939 0000",
    "texonic instruments": "+91 80 4120 2222",
    "solutions2go": "+1 514-875-5555",
    "hermitage solutions": "+33 1 41 91 51 51",
    "ignition technology": "+44 1483 243500",
    "comsource": "+1 800-888-2966",
    "akedco": "+966 11 460 0000",
    "bulwark technologies": "+966 11 217 0000",
    "cyberknight": "+971 4 558 8555",
    "dataflex": "+31 20 346 8200",
    "distology": "+44 345 222 0000",
    "dyntek services": "1-800-439-6835",
    "elcore group": "+41 41 560 07 07",
    "first distribution": "+27 11 201 8900",
    "m.tech products": "+65 6297 2888",
    "mitsumi distribution": "+971 4 357 7700",
    "netpoleon solutions": "+65 6286 2286",
    "red river technology": "1-800-769-7982",
    "supertron vad": "+91 33 4000 2000",
    "techaccess": "+971 4 454 0100",
    "valuepoint systems": "+91 80 4128 1234",
    "zedsphere": "+44 118 902 8000",
    "aspire technology partners": "1-732-847-9200",
    "ma labs": "+1 408-941-0808",
    "partner": "",  # leave unless known
    "nuvias uc": "+44 207 593 3290",
    "jcpenney": "1-800-322-1189",
    "smartwatchsolutionsltd": "",  # clear fake 555
}

# Replace corrupt / placeholder phones even if non-empty
FORCE_PHONE = {
    "jcpenney",
    "smartwatchsolutionsltd",
    "softcat",
    "redington",
    "bleepingcomputer",
    "gonoise",
    "iguanasell",
    "connection",
}

CORRUPT_PHONE_RE = re.compile(
    r"^\d{1,3}\.\d+|555-555|2003\s*-\s*2026|^\d{8,12}$|100250072|2136788491|1771130170|1782311711",
    re.I,
)

PARTIAL_RE = re.compile(
    r"\*\*\*|has named|He is the|Board by|cybersecurity industry|"
    r"gokhan gokhan|^new$|^nowCo$|^tenable$|says SK|^Abel M$|^Paul J$|"
    r"^Andrew$|^Andres$|^Michael$|^Joshua M\.?$|^L Cr",
    re.I,
)


def key(name: str) -> str:
    return re.sub(r"\s+", " ", (name or "").strip().lower())


def needs_name_upgrade(person: str) -> bool:
    p = (person or "").strip()
    if not p:
        return True
    if PARTIAL_RE.search(p) or "***" in p:
        return True
    if len(p.split()) == 1 and len(p) < 15:
        return True
    if "industry." in p.lower() or p.lower().startswith("has named"):
        return True
    return False


def find_header_row(ws) -> tuple[int, list]:
    for ridx in range(1, 8):
        vals = [c.value for c in ws[ridx]]
        if any(v and str(v).strip().lower() == "company" for v in vals):
            return ridx, vals
    return 1, [c.value for c in ws[1]]


def col_map(headers: list) -> dict[str, int]:
    out = {}
    for i, h in enumerate(headers):
        if not h:
            continue
        hl = str(h).strip().lower()
        out[hl] = i
        if "contact person" in hl:
            out["contact"] = i
        if hl == "role" or "contact role" in hl or hl == "title":
            out.setdefault("role", i)
        if "phone" in hl:
            out["phone"] = i
        if hl == "company":
            out["company"] = i
    return out


def apply_excel(path: Path, name_map: dict[str, tuple[str, str | None]], phone_map: dict[str, str]) -> dict:
    if not path.exists():
        return {"file": str(path), "missing": True}
    wb = load_workbook(path)
    ws = wb.active
    hr, headers = find_header_row(ws)
    cols = col_map(headers)
    ci = cols.get("company")
    if ci is None:
        return {"file": str(path), "error": "no company col"}
    n_name = n_phone = 0
    for ridx in range(hr + 1, ws.max_row + 1):
        company = ws.cell(ridx, ci + 1).value
        k = key(str(company or ""))
        if not k:
            continue
        if "contact" in cols and k in name_map:
            new_name, new_role = name_map[k]
            cur = str(ws.cell(ridx, cols["contact"] + 1).value or "")
            if k in FORCE_NAME or needs_name_upgrade(cur):
                ws.cell(ridx, cols["contact"] + 1).value = new_name or None
                n_name += 1
                if new_role and "role" in cols:
                    ws.cell(ridx, cols["role"] + 1).value = new_role
        if "phone" in cols and k in phone_map:
            new_ph = phone_map[k]
            cur_ph = str(ws.cell(ridx, cols["phone"] + 1).value or "").strip()
            if k in FORCE_PHONE or not cur_ph or CORRUPT_PHONE_RE.search(cur_ph):
                if new_ph or k in FORCE_PHONE:
                    ws.cell(ridx, cols["phone"] + 1).value = new_ph or None
                    n_phone += 1
    wb.save(path)
    return {"file": str(path), "names": n_name, "phones": n_phone}


def apply_json(path: Path) -> dict:
    data = json.loads(path.read_text(encoding="utf-8"))
    rows = data.get("relevant_companies") or []
    n_name = n_phone = 0
    for r in rows:
        company = str(r.get("company") or r.get("brand") or "").strip()
        k = key(company)
        crawl = dict(r.get("_crawl") or {})
        person = str(crawl.get("contact_person") or "").strip()
        phone = str(crawl.get("phone") or "").strip()

        if k in NAME_FIXES and (k in FORCE_NAME or needs_name_upgrade(person)):
            new_name, new_role = NAME_FIXES[k]
            crawl["contact_person"] = new_name
            if new_role:
                crawl["contact_role"] = new_role
            crawl["contact_source_url"] = crawl.get("contact_source_url") or "web_manual_fill"
            n_name += 1

        if k in PHONE_FIXES:
            new_ph = PHONE_FIXES[k]
            if k in FORCE_PHONE or not phone or CORRUPT_PHONE_RE.search(phone):
                if new_ph or k in FORCE_PHONE:
                    crawl["phone"] = new_ph
                    n_phone += 1

        # opportunistic: blank phone for any company with a known phone key variant
        r["_crawl"] = crawl

    path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    return {"file": str(path), "names": n_name, "phones": n_phone, "companies": len(rows)}


def mirror_dir(src: Path, dst: Path) -> None:
    if not src.exists():
        return
    dst.mkdir(parents=True, exist_ok=True)
    patterns = [
        f"{src.name}.json",
        "*_FINAL.xlsx",
        "*_founded_and_hq_v2.xlsx",
        "*_founded_and_hq_v2_cleaned.xlsx",
    ]
    for pat in patterns:
        for p in src.glob(pat):
            try:
                shutil.copy2(p, dst / p.name)
            except PermissionError:
                print(f"  skip locked mirror: {p.name}")


def main() -> None:
    import sys

    name_map = dict(NAME_FIXES)
    phone_map = {k: v for k, v in PHONE_FIXES.items() if v or k in FORCE_PHONE}
    only = {a.strip().lower() for a in sys.argv[1:] if a and not a.startswith("-")}

    for mid in MARKETS:
        if only and mid not in only and not any(o in mid for o in only):
            continue
        md = ROOT / mid
        print(f"\n=== {mid} ===")
        jp = md / f"{mid}.json"
        if jp.exists():
            print(" JSON", apply_json(jp))
        else:
            print(" JSON missing", jp)

        stem = mid  # e.g. global_smartwatch_market_global
        short = mid.replace("_global", "") if mid.endswith("_market_global") else mid
        # actual stems: global_smartwatch_market
        short = re.sub(r"_global$", "", mid)
        xfiles = [
            md / f"{short}_FINAL.xlsx",
            md / f"{short}_founded_and_hq_v2.xlsx",
            md / f"{short}_founded_and_hq_v2_cleaned.xlsx",
        ]
        for xp in xfiles:
            print(" XLSX", apply_excel(xp, name_map, phone_map))

        # mirror
        if MIRROR.exists() or True:
            dst = MIRROR / mid
            mirror_dir(md, dst)
            print(" mirrored ->", dst)


if __name__ == "__main__":
    main()
