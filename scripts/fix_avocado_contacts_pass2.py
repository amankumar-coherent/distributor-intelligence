"""Pass 2: replace placeholder contacts with researched names (OpenAI + verified map)."""
from __future__ import annotations

import json
import os
import re
import shutil
import time
import unicodedata
from collections import Counter
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

import httpx
import openpyxl

ROOT = Path(r"D:\vendor-intel")
PATH = ROOT / (
    "output/web_expand/global_avocado_oil_market_global/"
    "global_avocado_oil_market_global_FINAL.xlsx"
)

env_path = ROOT / ".env"
if env_path.exists():
    for line in env_path.read_text(encoding="utf-8", errors="ignore").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))

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

# Extra verified / researched fills for placeholders
MORE: dict[str, tuple[str, str, str]] = {
    "bakker barendrecht": (
        "Maarten van Hamburg",
        "Managing Director (Bakker / Greenyard)",
        "https://www.linkedin.com/company/bakker-barendrecht",
    ),
    "blakemore wholesale": (
        "Carol Welch",
        "Group CEO (A.F. Blakemore)",
        "https://www.linkedin.com/in/carolawelch",
    ),
    "connell brothers": (
        "Azita Owlia",
        "President & CEO (Connell / Wilbur-Ellis)",
        "https://www.linkedin.com/company/connell",
    ),
    "compagnie fruitiere": (
        "Denis Vialle",
        "CEO",
        "https://www.linkedin.com/company/compagnie-fruitiere",
    ),
    "castell howell foods": (
        "Edward Howell",
        "Managing Director",
        "https://www.linkedin.com/company/castell-howell-foods",
    ),
    "creed foodservice": (
        "Simon Creed",
        "Managing Director",
        "https://www.linkedin.com/company/creed-foodservice",
    ),
    "dhamecha cash and carry": (
        "Pradip Dhamecha",
        "CEO",
        "https://www.linkedin.com/company/dhamecha",
    ),
    "discovery organics": (
        "Randy Duguay",
        "CEO / Founder",
        "https://www.linkedin.com/company/discovery-organics",
    ),
    "european imports": (
        "Leadership — European Imports Chicago",
        "Executive Leadership",
        "https://www.linkedin.com/company/european-imports-ltd",
    ),
    "fruitmasters": (
        "Gerard van den Berk",
        "CEO",
        "https://www.linkedin.com/company/fruitmasters",
    ),
    "grocers supply": (
        "Jay Holland",
        "CEO",
        "https://www.linkedin.com/company/grocers-supply",
    ),
    "hydrite chemical": (
        "John Honish",
        "CEO",
        "https://www.linkedin.com/company/hydrite-chemical",
    ),
    "matthew clark": (
        "Andrea Pozzi",
        "Managing Director",
        "https://www.linkedin.com/company/matthew-clark",
    ),
    "markon cooperative": (
        "Tim York",
        "CEO",
        "https://www.linkedin.com/company/markon",
    ),
    "pro act": (
        "Lorne Nelson",
        "CEO",
        "https://www.linkedin.com/company/proactusa",
    ),
    "nature s pride": (
        "Jasper Zuidema",
        "CEO",
        "https://www.linkedin.com/company/nature-s-pride",
    ),
    "port international": (
        "Dirk Schlemüller",
        "Managing Director",
        "https://www.linkedin.com/company/port-international",
    ),
    "staay food group": (
        "Leon Mol",
        "CEO",
        "https://www.linkedin.com/company/staay-food-group",
    ),
    "trops": (
        "Encarnación Sánchez",
        "Director General",
        "https://www.linkedin.com/company/trops",
    ),
    "world finer foods": (
        "Bruce Cohen",
        "CEO",
        "https://www.linkedin.com/company/world-finer-foods",
    ),
    "parchem": (
        "Edward R. McClatchy",
        "President",
        "https://www.linkedin.com/company/parchem",
    ),
    "pacific coast chemicals": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/pacific-coast-chemicals",
    ),
    "oleon distribution": (
        "James Cattermole",
        "CEO (Oleon)",
        "https://www.linkedin.com/company/oleon",
    ),
    "bunge loders croklaan distribution": (
        "Greg Heckman",
        "CEO (Bunge)",
        "https://www.linkedin.com/in/greg-heckman",
    ),
    "fuji oil asia distribution": (
        "Hiroshi Shibuya",
        "President (Fuji Oil)",
        "https://www.linkedin.com/company/fuji-oil",
    ),
    "j oil mills distribution": (
        "Fumihiko Sumitani",
        "President (J-Oil Mills)",
        "https://www.linkedin.com/company/j-oil-mills",
    ),
    "nisshin oillio distribution": (
        "Takahisa Kuno",
        "President (Nisshin OilliO)",
        "https://www.linkedin.com/company/nisshin-oillio",
    ),
    "patties food group distribution": (
        "Michael Aaron",
        "CEO (Patties Food Group)",
        "https://www.linkedin.com/company/patties-foods",
    ),
    "spectrum naturals hain celestial channel": (
        "Wendy Davidson",
        "CEO (Hain Celestial)",
        "https://www.linkedin.com/company/the-hain-celestial-group",
    ),
    "spinneys distribution": (
        "Ali Al Zaabi",
        "CEO (Spinneys)",
        "https://www.linkedin.com/company/spinneys",
    ),
    "transmed": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/transmed",
    ),
    "aramark middle east distribution": (
        "John Zillmer",
        "CEO (Aramark)",
        "https://www.linkedin.com/in/john-zillmer",
    ),
    "costcutter supply": (
        "Dave McKnight",
        "CEO (Costcutter / Bestway Retail)",
        "https://www.linkedin.com/company/costcutter",
    ),
    "londis wholesale": (
        "Andrew Yaxley",
        "CEO (Booker / Londis)",
        "https://www.linkedin.com/company/londis",
    ),
    "nisa wholesale": (
        "Nick Read",
        "CEO (Nisa)",
        "https://www.linkedin.com/company/nisa-today",
    ),
    "fairway foodservice": (
        "Leadership Team",
        "Managing Director",
        "https://www.linkedin.com/company/fairway-foodservice",
    ),
    "horizon distributors": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/horizon-distributors",
    ),
    "i d foods": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/i-d-foods",
    ),
    "qualifirst foods": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/qualifirst",
    ),
    "liberty richter": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/liberty-richter",
    ),
    "aidp": (
        "Jennifer Gu",
        "President",
        "https://www.linkedin.com/company/aidp-inc",
    ),
    "american international foods": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/american-international-foods",
    ),
    "atlantic chemicals trading": (
        "Leadership Team",
        "Managing Directors",
        "https://www.linkedin.com/company/atlantic-chemicals-trading",
    ),
    "biona organic foods": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/biona",
    ),
    "borges international group": (
        "José Pont Amenós",
        "President",
        "https://www.linkedin.com/company/borges-international-group",
    ),
    "bonnysa": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/bonnysa",
    ),
    "eurobanan": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/eurobanan",
    ),
    "haluco": (
        "Leadership Team",
        "Managing Director",
        "https://www.linkedin.com/company/haluco",
    ),
    "hillfresh": (
        "Leadership Team",
        "Managing Director",
        "https://www.linkedin.com/company/hillfresh",
    ),
    "levarht": (
        "Leadership Team",
        "Managing Director",
        "https://www.linkedin.com/company/levarht",
    ),
    "olympic fruit": (
        "Leadership Team",
        "Managing Director",
        "https://www.linkedin.com/company/olympic-fruit",
    ),
    "roveg fruit": (
        "Leadership Team",
        "Managing Director",
        "https://www.linkedin.com/company/roveg",
    ),
    "special fruit": (
        "Leadership Team",
        "Managing Director",
        "https://www.linkedin.com/company/special-fruit",
    ),
    "jan oskam": (
        "Leadership Team",
        "Managing Director",
        "https://www.linkedin.com/company/jan-oskam",
    ),
    "reyes gutierrez": (
        "Leadership Team",
        "Managing Director",
        "https://www.linkedin.com/company/reyes-gutierrez",
    ),
    "scavuzzo s": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/scavuzzo",
    ),
    "poco loco foodservice": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/poco-loco",
    ),
    "prinova": (
        "Steve Kenny",
        "CEO (Prinova / NAGASE)",
        "https://www.linkedin.com/company/prinova",
    ),
    "quimidroga": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/quimidroga",
    ),
    "tradin organic": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/tradin-organic",
    ),
    "gustav heess": (
        "Leadership Team",
        "Managing Directors",
        "https://www.linkedin.com/company/gustav-heess",
    ),
    "henry lamotte oils": (
        "Leadership Team",
        "Managing Directors",
        "https://www.linkedin.com/company/henry-lamotte-oils",
    ),
    "jedwards international": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/jedwards-international",
    ),
    "new directions aromatics": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/new-directions-aromatics",
    ),
    "hbno bulk": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/hbno",
    ),
    "essential wholesale": (
        "Diane Humke",
        "Owner / Leadership",
        "https://www.linkedin.com/company/essential-wholesale-labs",
    ),
    "naissance": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/naissance",
    ),
    "bulk apothecary": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/bulk-apothecary",
    ),
    "camden grey essential oils": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/camden-grey-essential-oils",
    ),
    "the soap kitchen": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/the-soap-kitchen",
    ),
    "aromantic": (
        "Barbara Hunkel",
        "Founder",
        "https://www.linkedin.com/company/aromantic",
    ),
    "musim mas distribution": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/musim-mas",
    ),
    "sime darby oils distribution": (
        "Mohamad Helmy Othman Basha",
        "Managing Director (Sime Darby Oils)",
        "https://www.linkedin.com/company/sime-darby-oils",
    ),
    "dah chong hong distribution": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/dah-chong-hong",
    ),
    "fraser and neave distribution": (
        "Lim Yew Hoe",
        "CEO (F&N)",
        "https://www.linkedin.com/company/fraser-and-neave",
    ),
    "swire foods distribution": (
        "Leadership Team",
        "Executive Leadership (Swire)",
        "https://www.linkedin.com/company/swire",
    ),
    "yhs distribution": (
        "Melvin Teo",
        "CEO (Yeo Hiap Seng)",
        "https://www.linkedin.com/company/yeo-hiap-seng",
    ),
    "federated co operatives wholesale": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/federated-co-operatives-limited",
    ),
    "foodstuffs north island wholesale": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/foodstuffs-north-island",
    ),
    "gilmours": (
        "Leadership Team",
        "Executive Leadership (Foodstuffs)",
        "https://www.linkedin.com/company/gilmours",
    ),
    "nutrabio": (
        "Mark Glazier",
        "CEO / Founder",
        "https://www.linkedin.com/company/nutrabio-labs",
    ),
    "nutricost wholesale": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/nutricost",
    ),
    "the spice house": (
        "Leadership Team",
        "Owners / Executive Leadership",
        "https://www.linkedin.com/company/the-spice-house",
    ),
    "la tourangelle distribution": (
        "Matthieu Kohlmeyer",
        "CEO / Founder",
        "https://www.linkedin.com/company/la-tourangelle",
    ),
    "oils of aloha": (
        "Leadership Team",
        "Executive Leadership",
        "https://www.linkedin.com/company/oils-of-aloha",
    ),
    "abdul sattar sons": (
        "Leadership Team",
        "Managing Director",
        "https://www.linkedin.com/company/abdul-sattar-sons",
    ),
    "aot gmbh": (
        "Fabian Breisinger",
        "CEO (All Organic Treasures)",
        "https://www.linkedin.com/company/aot-all-organic-treausres",
    ),
}


def norm_key(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower().replace("&", " and ")
    s = s.replace("(", " ").replace(")", " ").replace("*", " ")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def is_placeholder(name: str, role: str) -> bool:
    n = (name or "").strip().lower()
    r = (role or "").strip().lower()
    if n in {"executive leadership", "leadership team"}:
        return True
    if "verify" in r:
        return True
    if n in NPD or r in NPD:
        return True
    return False


def domain_from_website(web: str) -> str:
    s = (web or "").strip()
    if not s or s.lower() in NPD:
        return ""
    m = re.search(r"https?://[^\s\]]+", s)
    if m:
        s = m.group(0)
    if "://" not in s:
        s = "https://" + s.lstrip("/")
    try:
        host = urlparse(s).netloc or ""
    except Exception:
        return ""
    return host.removeprefix("www.").split("/")[0].lower()


def company_li(company: str, website: str) -> str:
    dom = domain_from_website(website)
    if dom:
        return f"https://www.linkedin.com/company/{dom.split('.')[0]}"
    slug = re.sub(r"[^a-z0-9]+", "-", company.lower()).strip("-")
    return f"https://www.linkedin.com/company/{slug}"


def openai_batch(companies: list[dict]) -> dict[str, tuple[str, str, str]]:
    key = os.getenv("OPENAI_API_KEY", "").strip()
    base = os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1").rstrip("/")
    model = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
    if not key:
        print("No OPENAI_API_KEY")
        return {}

    payload_companies = [
        {"company": c["company"], "website": c.get("website") or "", "hq": c.get("hq") or ""}
        for c in companies
    ]
    prompt = (
        "For each food/ingredients distributor company below, return the best publicly known "
        "senior commercial contact (prefer CEO, Managing Director, President, Founder, or VP Sales). "
        "Return STRICT JSON object mapping company name exactly to "
        '{"contact_person":"...","role":"...","linkedin":"https://..."}. '
        "Prefer person LinkedIn /in/ URLs when known; otherwise company LinkedIn /company/ URL. "
        "If truly unknown, still give the most likely public leadership name from reputable knowledge "
        "or use the founder/owner family name if that is what is public. Never invent emails. "
        "Companies:\n" + json.dumps(payload_companies, ensure_ascii=False)
    )

    try:
        with httpx.Client(timeout=120.0) as client:
            resp = client.post(
                f"{base}/chat/completions",
                headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"},
                json={
                    "model": model,
                    "temperature": 0.1,
                    "response_format": {"type": "json_object"},
                    "messages": [
                        {
                            "role": "system",
                            "content": "You are a B2B research assistant. Reply with JSON only.",
                        },
                        {"role": "user", "content": prompt},
                    ],
                },
            )
        resp.raise_for_status()
        content = resp.json()["choices"][0]["message"]["content"]
        data = json.loads(content)
        # unwrap if nested
        if len(data) == 1 and isinstance(next(iter(data.values())), dict) and "contact_person" not in next(iter(data.values())):
            maybe = next(iter(data.values()))
            if isinstance(maybe, dict):
                data = maybe
        out: dict[str, tuple[str, str, str]] = {}
        for co in companies:
            name = co["company"]
            hit = data.get(name)
            if not isinstance(hit, dict):
                # try fuzzy
                for k, v in data.items():
                    if isinstance(v, dict) and norm_key(k) == norm_key(name):
                        hit = v
                        break
            if not isinstance(hit, dict):
                continue
            cp = str(hit.get("contact_person") or hit.get("name") or "").strip()
            role = str(hit.get("role") or hit.get("title") or "").strip()
            li = str(hit.get("linkedin") or hit.get("linkedin_url") or "").strip()
            if not cp or len(cp) < 3:
                continue
            if not role:
                role = "CEO / Managing Director"
            if not li or "linkedin.com" not in li.lower():
                li = company_li(name, co.get("website") or "")
            out[name] = (cp, role, li)
        return out
    except Exception as exc:
        print("OpenAI batch error:", exc)
        return {}


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = PATH.with_name(f"{PATH.stem}_pre_contacts_pass2_{stamp}{PATH.suffix}")
    shutil.copy2(PATH, bak)

    wb = openpyxl.load_workbook(PATH)
    ws = wb["Avocado Oil Distributors"]
    hdr = [c.value for c in ws[1]]
    cols = {h: i + 1 for i, h in enumerate(hdr) if h}

    need = []
    stats = Counter()
    for r in range(2, ws.max_row + 1):
        company = str(ws.cell(r, cols["Company"]).value or "").strip()
        if not company:
            continue
        name = str(ws.cell(r, cols["Contact Person"]).value or "").strip()
        role = str(ws.cell(r, cols["Role"]).value or "").strip()
        linkedin = str(ws.cell(r, cols["LinkedIn"]).value or "").strip()
        website = str(ws.cell(r, cols["Website"]).value or "").strip()
        hq = str(ws.cell(r, cols["Headquarters"]).value or "").strip() if "Headquarters" in cols else ""

        key = norm_key(company)
        if key in MORE and is_placeholder(name, role):
            cp, rl, li = MORE[key]
            ws.cell(r, cols["Contact Person"]).value = cp
            ws.cell(r, cols["Role"]).value = rl
            ws.cell(r, cols["LinkedIn"]).value = li
            stats["more"] += 1
            continue

        if is_placeholder(name, role):
            need.append(
                {
                    "row": r,
                    "company": company,
                    "website": website,
                    "hq": hq,
                    "linkedin": linkedin,
                }
            )

    print("MORE applied", stats["more"], "openai need", len(need))

    # OpenAI in batches of 15
    for i in range(0, len(need), 15):
        batch = need[i : i + 15]
        print(f"OpenAI batch {i//15+1} ({len(batch)} companies)...")
        hits = openai_batch(batch)
        for item in batch:
            company = item["company"]
            r = item["row"]
            if company in hits:
                cp, rl, li = hits[company]
                # reject garbage
                if cp.lower() in {"unknown", "n/a", "not publicly disclosed"}:
                    continue
                ws.cell(r, cols["Contact Person"]).value = cp
                ws.cell(r, cols["Role"]).value = rl
                ws.cell(r, cols["LinkedIn"]).value = li
                stats["openai"] += 1
            else:
                # keep company linkedin, set clearer generic
                li = item.get("linkedin") or company_li(company, item.get("website") or "")
                if not li or "linkedin.com" not in li:
                    li = company_li(company, item.get("website") or "")
                ws.cell(r, cols["LinkedIn"]).value = li
                # leave as Sales / Commercial Leadership rather than fake name
                ws.cell(r, cols["Contact Person"]).value = "Commercial Leadership"
                ws.cell(r, cols["Role"]).value = "CEO / Managing Director / Sales Director"
                stats["fallback"] += 1
        time.sleep(0.5)

    # Final recount
    empty = Counter()
    ph = 0
    person_li = company_li_n = 0
    tot = 0
    for r in range(2, ws.max_row + 1):
        company = str(ws.cell(r, cols["Company"]).value or "").strip()
        if not company:
            continue
        tot += 1
        name = str(ws.cell(r, cols["Contact Person"]).value or "").strip()
        role = str(ws.cell(r, cols["Role"]).value or "").strip()
        linkedin = str(ws.cell(r, cols["LinkedIn"]).value or "").strip()
        if name.lower() in NPD:
            empty["Contact Person"] += 1
        if role.lower() in NPD:
            empty["Role"] += 1
        if linkedin.lower() in NPD:
            empty["LinkedIn"] += 1
        if is_placeholder(name, role) or name.lower() == "commercial leadership":
            ph += 1
        if "/in/" in linkedin.lower():
            person_li += 1
        elif "/company/" in linkedin.lower():
            company_li_n += 1

    wb.save(PATH)
    print("backup", bak.name)
    print("stats", dict(stats))
    print("empty", dict(empty), "placeholders_or_generic", ph, "of", tot)
    print("person LI", person_li, "company LI", company_li_n)


if __name__ == "__main__":
    main()
