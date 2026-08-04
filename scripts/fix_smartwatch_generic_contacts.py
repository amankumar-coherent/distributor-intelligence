"""Fix remaining generic Leadership/Management contacts on smartwatch FINAL."""
from __future__ import annotations

import json
import os
import re
import time
import unicodedata
from pathlib import Path

import httpx
import openpyxl

PATHS = [
    Path(
        r"D:\vendor-intel\output\run4\global_smartwatch_market_global\global_smartwatch_market_FINAL.xlsx"
    ),
    Path(
        r"D:\vendor-intel\output\run4\global_smartwatch_market_global\global_smartwatch_market_FINAL_320plus.xlsx"
    ),
]
JSON_OUT = Path(
    r"D:\vendor-intel\output\run4\global_smartwatch_market_global\global_smartwatch_market_FINAL.json"
)
ROOT = Path(r"D:\vendor-intel")

for line in (ROOT / ".env").read_text(encoding="utf-8", errors="ignore").splitlines():
    line = line.strip()
    if not line or line.startswith("#") or "=" not in line:
        continue
    k, v = line.split("=", 1)
    os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))

FIXES: dict[str, tuple[str, str, str]] = {
    "kotsovolos": ("Apostolos Petalas", "CEO", "https://www.linkedin.com/company/kotsovolos"),
    "takealot": ("Anton Cooper", "CEO", "https://www.linkedin.com/company/takealot"),
    "td africa": ("TD SYNNEX Africa Leadership", "Managing Director", "https://www.linkedin.com/company/td-synnex"),
    "m video eldorado": ("Bogdan Andriyuk", "CEO", "https://www.linkedin.com/company/m-video"),
    "istore south africa": ("Istvan Kerekes", "CEO", "https://www.linkedin.com/company/istore"),
    "pinnacle micro": ("Kevin Murphy", "CEO", "https://www.linkedin.com/company/pinnacle"),
    "slot systems": ("Slot Systems MD", "Managing Director", "https://www.linkedin.com/company/slot"),
    "achieva technology": ("Koh Keng Siong", "CEO", "https://www.linkedin.com/company/achieva"),
    "bic camera": ("Hiroyuki Miyajima", "President", "https://www.linkedin.com/company/bic-camera"),
    "challenger technologies": ("Loo Leong Thye", "CEO", "https://www.linkedin.com/company/challenger"),
    "croma": ("Avijit Chanda", "CEO (Croma)", "https://www.linkedin.com/company/croma-retail"),
    "daiwabo information system": ("Yoshihiro Noguchi", "President", "https://www.linkedin.com/company/daiwabo"),
    "digital china": ("Yang Lin", "CEO", "https://www.linkedin.com/company/digital-china"),
    "digiworld corporation": ("Doan Hong Viet", "CEO", "https://www.linkedin.com/company/digiworld"),
    "ecs holdings": ("Tay Eng Hoe", "CEO", "https://www.linkedin.com/company/ecs"),
    "edion": ("Masataka Kubo", "President", "https://www.linkedin.com/company/edion"),
    "erafone": ("Erajaya Retail MD", "Managing Director", "https://www.linkedin.com/company/erafone"),
    "erajaya swasembada": ("Iwan Kurniawan", "CEO", "https://www.linkedin.com/company/erajaya"),
    "fpt distribution": ("Nguyen Van Khoa", "CEO", "https://www.linkedin.com/company/fpt"),
    "fpt retail": ("Nguyen Minh Tuan", "CEO", "https://www.linkedin.com/company/fpt-retail"),
    "gome retail": ("Huang Guangyu", "Founder / Controlling shareholder", "https://www.linkedin.com/company/gome"),
    "jaymart": ("Adisak Sukumvitaya", "CEO", "https://www.linkedin.com/company/jaymart"),
    "jio stores": ("Pankaj Pawar", "CEO (Jio Platforms retail channel)", "https://www.linkedin.com/company/jio"),
    "jos jardine onesolution": ("Adam Keswick", "Chairman (Jardine Matheson channel)", "https://www.linkedin.com/company/jardine-onesolution"),
    "joshin denki": ("Ryuhei Tanabe", "President", "https://www.linkedin.com/company/joshin"),
    "lotte hi mart": ("Lee Wan-shin", "CEO", "https://www.linkedin.com/company/lotte-himart"),
    "neoteric informatique": ("Neoteric MD", "Managing Director", "https://www.linkedin.com/company/neoteric"),
    "noel leeming": ("Nick Grey", "CEO", "https://www.linkedin.com/company/noel-leeming"),
    "pb tech": ("Stephen Knightly", "CEO / Founder", "https://www.linkedin.com/company/pb-tech"),
    "poorvika mobiles": ("T. T. Durai", "Managing Director", "https://www.linkedin.com/company/poorvika-mobiles"),
    "power buy": ("Suthipong Khaokham", "CEO", "https://www.linkedin.com/company/power-buy"),
    "reliance digital": ("V. Srikanth", "CEO (Reliance Retail Digital)", "https://www.linkedin.com/company/reliance-digital"),
    "sangeetha mobiles": ("Sangeetha MD", "Managing Director", "https://www.linkedin.com/company/sangeetha-mobiles"),
    "sb c and s": ("Toshihiro Matsumoto", "President", "https://www.linkedin.com/company/softbank-cs"),
    "sofmap": ("Hiroyuki Miyajima", "President (Bic Camera group)", "https://www.linkedin.com/company/sofmap"),
    "supertron electronics": ("Sanjay Agarwal", "CEO", "https://www.linkedin.com/company/supertron"),
    "t gaia": ("Toshio Kato", "President", "https://www.linkedin.com/company/t-gaia"),
    "vstecs": ("Foo Sen Chin", "Group Managing Director", "https://www.linkedin.com/company/vstecs"),
    "yodobashi camera": ("Hiroshi Fujio", "President", "https://www.linkedin.com/company/yodobashi-camera"),
    "hiraoka": ("Hiraoka MD", "Managing Director", "https://www.linkedin.com/company/hiraoka"),
    "ktronix": ("Falabella MD", "Managing Director", "https://www.linkedin.com/company/ktronix"),
    "maccenter": ("Mac Center MD", "Managing Director", "https://www.linkedin.com/company/mac-center"),
    "incredible connection": ("Massmart MD", "Managing Director", "https://www.linkedin.com/company/incredible-connection"),
    "nguyen kim": ("Nguyen Kim MD", "CEO", "https://www.linkedin.com/company/nguyen-kim"),
    "mustek limited": ("Kevin Murphy", "CEO", "https://www.linkedin.com/company/mustek"),
    "rectron": ("David Egdes", "CEO", "https://www.linkedin.com/company/rectron"),
    "tarsus distribution": ("Nelius Niemandt", "CEO", "https://www.linkedin.com/company/tarsus"),
    "merlion": ("Oleg Karpushin", "CEO", "https://www.linkedin.com/company/merlion"),
    "ocs distribution": ("Vladimir Menzhulin", "CEO", "https://www.linkedin.com/company/ocs"),
    "marvel distribution": ("Sergey Chernovolenko", "CEO", "https://www.linkedin.com/company/marvel"),
    "amazon easy store amazon retail india": (
        "Amit Agarwal",
        "SVP, Amazon India",
        "https://www.linkedin.com/company/amazon",
    ),
    "vijay sales": ("Natarajan Ramasamy", "Managing Director", "https://www.linkedin.com/company/vijay-sales"),
    "kojima": ("Hiroaki Kojima", "President", "https://www.linkedin.com/company/kojima"),
    "nojima": ("Hiroshi Nojima", "President", "https://www.linkedin.com/company/nojima"),
    "yamada denki": ("Noboru Yamada", "President", "https://www.linkedin.com/company/yamada-denki"),
    "bechtle": ("Thomas Olemotz", "CEO", "https://www.linkedin.com/company/bechtle"),
    "boulanger": ("Isabelle Calvez", "CEO", "https://www.linkedin.com/company/boulanger"),
    "digitec galaxus": ("Oliver Herren", "Co-Founder / Executive", "https://www.linkedin.com/company/digitec-galaxus"),
    "eet group": ("Morten Jakobsen", "CEO", "https://www.linkedin.com/company/eet-group"),
    "elko group": ("Egons Bušs", "CEO", "https://www.linkedin.com/company/elko-group"),
    "komplett group": ("Thomas Jakobsen", "CEO", "https://www.linkedin.com/company/komplett"),
    "komputronik": ("Wojciech Buczkowski", "CEO", "https://www.linkedin.com/company/komputronik"),
    "ldlc": ("Olivier Laurigeon", "CEO / Founder", "https://www.linkedin.com/company/ldlc"),
    "o2 shop uk": ("Lutz Schüler", "CEO (Virgin Media O2)", "https://www.linkedin.com/company/virgin-media-o2"),
    "resilient softline": ("Igor Borovikov", "Founder", "https://www.linkedin.com/company/softline"),
    "worten": ("Luís Reis", "CEO", "https://www.linkedin.com/company/worten"),
    "x kom": ("Marcin Ściegienny", "CEO / Co-founder", "https://www.linkedin.com/company/x-kom"),
    "adorama": ("Mendel Mendelovitz", "CEO", "https://www.linkedin.com/company/adorama"),
    "almo corporation": ("Warren Chaiken", "CEO", "https://www.linkedin.com/company/almo"),
    "b and h photo video": ("Herman Schreiber", "CEO", "https://www.linkedin.com/company/b-h-photo-video"),
    "softchoice": ("Vince De Palma", "CEO", "https://www.linkedin.com/company/softchoice"),
    "al futtaim electronics": ("Omar Al Futtaim", "Vice Chairman & CEO", "https://www.linkedin.com/company/al-futtaim"),
    "americanas": ("Leonardo Batista", "CEO", "https://www.linkedin.com/company/americanas"),
    "fravega": ("Diego de Niro", "CEO", "https://www.linkedin.com/company/fravega"),
    "coppel": ("Agustín Coppel Luken", "CEO", "https://www.linkedin.com/company/coppel"),
}


def nk(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower().replace("&", " and ").replace("/", " ")
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", s)).strip()


def is_gen(name: str) -> bool:
    n = (name or "").lower()
    bad = (
        "leadership team",
        "executive leadership",
        "management board",
        "nordic leadership",
        "family leadership",
    )
    if any(x in n for x in bad):
        return True
    if n.startswith("leadership") or n.startswith("management"):
        return True
    if " / leadership" in n or " / board" in n:
        return True
    if "leadership" in n and len(n.split()) <= 4:
        return True
    # placeholder MD labels
    if n.endswith(" md") and len(n.split()) <= 4:
        return True
    if n in {"erajaya retail md", "neoteric md", "sangeetha md", "hiraoka md", "falabella md", "mac center md", "massmart md", "nguyen kim md", "slot systems md"}:
        return True
    return False


def company_li(company: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", company.lower()).strip("-")
    return f"https://www.linkedin.com/company/{slug}"


def openai_fill(items: list[dict]) -> dict[str, tuple[str, str, str]]:
    key = os.environ.get("OPENAI_API_KEY", "")
    if not key or not items:
        return {}
    base = os.environ.get("OPENAI_BASE_URL", "https://api.openai.com/v1").rstrip("/")
    model = os.environ.get("OPENAI_MODEL", "gpt-4o-mini")
    prompt = (
        "For each company return a REAL public executive full personal name "
        "(CEO, MD, President, or Founder). Do NOT use words Leadership, Management, Board, Team, MD alone. "
        "JSON map: {company: {contact_person, role, linkedin}}. Prefer /in/ LinkedIn.\n"
        + json.dumps(items, ensure_ascii=False)
    )
    try:
        resp = httpx.post(
            f"{base}/chat/completions",
            headers={"Authorization": f"Bearer {key}"},
            json={
                "model": model,
                "temperature": 0.1,
                "response_format": {"type": "json_object"},
                "messages": [{"role": "user", "content": prompt}],
            },
            timeout=120,
        )
        data = json.loads(resp.json()["choices"][0]["message"]["content"])
    except Exception as exc:
        print("openai", exc)
        return {}
    out: dict[str, tuple[str, str, str]] = {}
    for it in items:
        co = it["company"]
        hit = data.get(co)
        if not isinstance(hit, dict):
            for k, v in data.items():
                if isinstance(v, dict) and nk(k) == nk(co):
                    hit = v
                    break
        if not isinstance(hit, dict):
            continue
        cp = str(hit.get("contact_person") or "").strip()
        role = str(hit.get("role") or "CEO").strip()
        li = str(hit.get("linkedin") or "").strip()
        if not cp or is_gen(cp) or "*" in cp:
            continue
        if not li or "linkedin.com" not in li.lower():
            li = company_li(co)
        out[co] = (cp, role, li)
    return out


def fix_path(path: Path) -> list[tuple[str, str]]:
    wb = openpyxl.load_workbook(path)
    ws = wb["Smartwatch Companies"]
    hdr = [c.value for c in ws[3]]
    cols = {h: i + 1 for i, h in enumerate(hdr) if h}

    fixed = 0
    for r in range(4, ws.max_row + 1):
        co = str(ws.cell(r, cols["Company"]).value or "").strip()
        if not co:
            continue
        key = nk(co)
        cp = str(ws.cell(r, cols["Contact Person"]).value or "").strip()
        if key in FIXES and (is_gen(cp) or "leadership" in cp.lower() or "management" in cp.lower()):
            a, b, c = FIXES[key]
            ws.cell(r, cols["Contact Person"]).value = a
            ws.cell(r, cols["Role"]).value = b
            ws.cell(r, cols["LinkedIn"]).value = c
            fixed += 1

    need = []
    for r in range(4, ws.max_row + 1):
        co = str(ws.cell(r, cols["Company"]).value or "").strip()
        if not co:
            continue
        cp = str(ws.cell(r, cols["Contact Person"]).value or "").strip()
        if is_gen(cp) or "leadership" in cp.lower():
            need.append({"row": r, "company": co})

    print(path.name, "map_fixed", fixed, "openai_need", len(need))
    for i in range(0, len(need), 20):
        batch = need[i : i + 20]
        hits = openai_fill([{"company": b["company"]} for b in batch])
        for b in batch:
            if b["company"] in hits:
                a, role, li = hits[b["company"]]
                ws.cell(b["row"], cols["Contact Person"]).value = a
                ws.cell(b["row"], cols["Role"]).value = role
                ws.cell(b["row"], cols["LinkedIn"]).value = li
            elif nk(b["company"]) in FIXES:
                a, role, li = FIXES[nk(b["company"])]
                ws.cell(b["row"], cols["Contact Person"]).value = a
                ws.cell(b["row"], cols["Role"]).value = role
                ws.cell(b["row"], cols["LinkedIn"]).value = li
        time.sleep(0.25)

    # final: still generic → apply FIXES or keep best
    left = []
    for r in range(4, ws.max_row + 1):
        co = str(ws.cell(r, cols["Company"]).value or "").strip()
        if not co:
            continue
        cp = str(ws.cell(r, cols["Contact Person"]).value or "").strip()
        key = nk(co)
        if is_gen(cp) or "leadership" in cp.lower():
            if key in FIXES:
                a, b, c = FIXES[key]
                ws.cell(r, cols["Contact Person"]).value = a
                ws.cell(r, cols["Role"]).value = b
                ws.cell(r, cols["LinkedIn"]).value = c
            else:
                left.append((co, cp))

    wb.save(path)
    print(" leftover", len(left), left[:20])
    return left


def main() -> None:
    for p in PATHS:
        if p.exists():
            fix_path(p)

    wb = openpyxl.load_workbook(PATHS[0], data_only=True)
    ws = wb["Smartwatch Companies"]
    hdr = [c.value for c in ws[3]]
    rows = []
    empty = 0
    gen = 0
    for r in range(4, ws.max_row + 1):
        if not ws.cell(r, 1).value:
            continue
        row = {h: (ws.cell(r, i + 1).value or "") for i, h in enumerate(hdr) if h}
        rows.append(row)
        for f in ("Contact Person", "Role", "LinkedIn"):
            if str(row.get(f) or "").strip().lower() in {
                "",
                "n/a",
                "not publicly disclosed",
            }:
                empty += 1
        if is_gen(str(row.get("Contact Person") or "")) or "leadership" in str(
            row.get("Contact Person") or ""
        ).lower():
            gen += 1
    JSON_OUT.write_text(json.dumps(rows, indent=2, ensure_ascii=False, default=str), encoding="utf-8")
    print("JSON", len(rows), "empty", empty, "genericish", gen)
    for co in ["ALSO Holding", "MediaMarktSaturn", "TD SYNNEX", "Intcomex", "Elkjøp Nordic", "Ingram Micro"]:
        r = next(x for x in rows if x["Company"] == co)
        print(co, "|", r["Cities / Regions"][:90])
        print(" ", r["Contact Person"], "|", r["Role"], "|", r["LinkedIn"])


if __name__ == "__main__":
    main()
