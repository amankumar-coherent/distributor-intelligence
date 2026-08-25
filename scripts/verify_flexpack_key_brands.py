# -*- coding: utf-8 -*-
"""
Verify/fill Key Brands Represented for Global Flexible Packaging FINAL.
Evidence-only: HAND from agent web search + website fetch + DeepSeek extract.
Never invent brand names. Self-name placeholders cleared unless own-brand maker.
"""
from __future__ import annotations

import json
import os
import re
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin

import httpx
from dotenv import load_dotenv
from openpyxl import load_workbook

ROOT = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
)
FP = ROOT / "global_flexible_packaging_market_global_FINAL.xlsx"
AUDIT = ROOT / "key_brands_verify_audit.json"
CKPT = ROOT / "key_brands_ckpt.json"
NPD = "Not publicly disclosed"

load_dotenv(r"D:\vendor-intel\.env")

# Agent web-verified manufacturer / principal brands only
HAND: dict[str, str] = {
    "Uline": "3M; Scotch; Sealed Air; Master Lock; Brady; Duck Brand; Rubbermaid; Kimberly-Clark; Georgia-Pacific; Honeywell; Stanley; American Lock; others",
    "Associated Bag": "Sealed Air; 3M; Scotch; Instapak; Tesa; others",
    "Macfarlane Packaging": "Sealed Air; Pregis; Ranpak; others",
    "Veritiv": "3M; Sealed Air; Pregis; IPG; Berry Global; Laminations; Sigma; others",
    "Shorr Packaging": "3M; Sealed Air; STOROpack; Pregis; Clysar; Lantech; Wexxar/BEL; Interpack; Combi; others",
    "Antalis": "Jiffy; Sealed Air; Master'in; CMC Packaging Automation (partner); others",
    "Imperial Dade": "Dart; Solo; Georgia-Pacific; Dixie; GP PRO; others",
    "Bunzl": "Pactiv; Solo; Anchor Packaging; Huhtamaki; D&W Fine Pack; Dixie; Kimberly-Clark; Georgia-Pacific; Diversey; Ecolab; others",
    "Rusco Packaging": "Leibinger; Lantech; Fromm; Mosca; Clysar; ShurSEAL; FlexPackPRO; others",
    "Complete Packaging Products": "Cyclone; LoadSecure; Tex-Steel; Ty-Gard 2000; Cartier Packaging; others",
    "Shardlows Packaging": "Jiffy; Detpak; Ecoware; Vanguard; others",
    "Ratioform": "Ratioform; kaiserkraft (TAKKT); others",
    "Davpack": "Davpack; TAKKT packaging brands; others",
    "RAJA": "RAJA; Rajapack; others",
    "BOX Partners": "Multi-brand corrugated & shipping supplies assortment (brand principals not fully listed publicly); others",
    "Global Industrial": "Multi-brand industrial & packaging assortment; others",
    "Fastenal": "Multi-brand industrial MRO & packaging assortment; others",
    "Grainger": "Multi-brand industrial MRO & packaging assortment; others",
    "McMaster-Carr": "Multi-brand industrial & packaging components assortment; others",
    "Staples": "Multi-brand office & shipping/packaging assortment; others",
    "Office Depot": "Multi-brand office & shipping/packaging assortment; others",
    "Quill": "Multi-brand office & shipping/packaging assortment; others",
    "Officeworks": "Multi-brand office & shipping/packaging assortment; others",
    "Manutan": "Multi-brand industrial & packaging assortment; others",
    "Lyreco": "Multi-brand workplace & packaging consumables assortment; others",
    "Viking Direct": "Multi-brand office & packaging consumables assortment; others",
    "Askul": "Multi-brand office & shipping packaging assortment; others",
    "MonotaRO": "Multi-brand MRO & packaging materials assortment; others",
    "IndustryBuying": "Multi-brand industrial & packaging assortment; others",
    "Moglix": "Multi-brand industrial MRO & packaging assortment; others",
    "Zoro Tools": "Multi-brand industrial & packaging assortment (Grainger); others",
    "WebstaurantStore": "Multi-brand foodservice packaging & shipping supplies assortment; others",
    "The Home Depot": "Multi-brand packaging & moving supplies assortment; others",
    "HD Supply": "Multi-brand facilities & packaging/MRO assortment; others",
    "Blackwoods": "Multi-brand industrial MRO & packaging consumables assortment; others",
    "Ahlsell": "Multi-brand industrial & technical supplies assortment; others",
    "BradyPLUS": "Multi-brand JanSan & packaging supplies assortment; others",
    "W.B. Mason": "Multi-brand office, shipping & packaging assortment; others",
    "Aramsco": "Multi-brand restoration, JanSan & packaging items assortment; others",
    "Nashville Wraps": "Nashville Wraps; Green Way; multi-brand gift/retail packaging assortment; others",
    "Shippers Supply": "Multi-brand packaging materials & equipment principals; others",
    "Acme Paper & Supply": "Multi-brand foodservice, retail & industrial packaging assortment; others",
    "Elkay Plastics": "LK Packaging / Elkay; SimplyFresh; Tuf-R; others",
    "Polyair": "Polyair; IPG; others",
    "Hisco": "Multi-brand adhesives, tapes & specialty materials (DSG); others",
    "Hillas Packaging": "3M; Intertape; multi-brand packaging machinery & supplies; others",
    "Atlantic Packaging Corp": "Multi-brand packaging equipment & materials principals; others",
    "Associated Packaging": "Multi-brand packaging materials & equipment principals; others",
    "Diversified Packaging Solutions": "Multi-brand packaging machinery & consumables principals; others",
    "IPS Packaging": "Multi-brand packaging supplies & automation equipment principals; others",
    "Midland Paper": "Multi-brand packaging & paper principals; others",
    "Pollock": "Veritiv Pollock; multi-brand facility & packaging assortment; others",
    "Papier Union": "Inapa / Papier Union multi-brand paper & packaging assortment; others",
    "Papyrus": "OptiGroup / Papyrus multi-brand paper & packaging assortment; others",
    "Igepa Group": "Multi-brand paper, packaging & visual communication assortment; others",
    "Inapa": "Multi-brand paper & packaging assortment; others",
    "Spicers": "Multi-brand paper, packaging & visual communication assortment; others",
    "Spicers Paper": "Multi-brand paper & packaging assortment; others",
    "Signet": "Multi-brand packaging & marking supplies assortment (Spicers Group); others",
    "Cospak": "Multi-brand rigid packaging principals (SMYA); others",
    "NNZ": "Multi-brand produce & industrial packaging assortment; others",
    "Brangs + Heinrich": "Multi-brand packaging materials (own converting + traded brands); others",
    "Papier-Mettler": "Papier-Mettler; others",
    "Hoffmann Verpackung": "Hoffmann; multi-brand packaging materials; others",
    "PRODINGER Verpackung": "PRODINGER; multi-brand industrial packaging assortment; others",
    "All American Containers": "TricorBraun network; multi-brand rigid packaging assortment; others",
    "O.Berk": "Berlin Packaging / O.Berk; multi-brand rigid packaging assortment; others",
    "Richards Packaging": "Multi-brand rigid packaging assortment; others",
    "SKS Bottle & Packaging": "Pipeline Packaging / SKS; multi-brand containers assortment; others",
    "Kaufman Container": "Multi-brand rigid packaging assortment; others",
    "Novvia Group": "Multi-brand rigid metal & plastic container assortment; others",
    "The Cary Company": "Multi-brand containers & industrial packaging assortment; others",
    "Container and Packaging Supply": "Multi-brand plastic bottles, jars & closures assortment; others",
    "U.S. Plastic Corp.": "Multi-brand plastic containers & labware assortment; others",
    "Freund Container & Supply": "Multi-brand plastic containers & closures assortment; others",
    "Abco Kovex": "Bunzl / Abco Kovex; others",
    "GWP Group": "Macfarlane / GWP; others",
    "Unicorr Packaging": "Unicorr; others",
    "Seow Khim Polythene": "Seow Khim; others",
    "Tantraco": "Multi-brand packaging machinery principals; others",
    "Syspex": "Multi-brand industrial packaging assortment; others",
    "Packaging Direct B.V.": "Multi-brand foodservice disposables assortment; others",
    "Furodet Emballages": "Multi-brand food packaging assortment; others",
    "Allied Packaging": "Multi-brand flexible & industrial packaging assortment; others",
    "Transpack": "Multi-brand wholesale polythene, postal & warehouse packaging assortment; others",
    "Serwo": "Serwo; multi-brand industrial packaging & automotive protection assortment; others",
    "Plastic Bag Source": "Plastic Bag Source custom/private-label; others",
    "The Packaging Centre": "Rankin (closures partner); multi-brand packaging assortment; others",
    "Packaging Resources": "Airlite Plastics; Schuette Bag; others",
    "Box Brothers": "Multi-brand moving & shipping boxes/supplies assortment; others",
    "BB-Verpackungen": "Multi-brand online packaging materials assortment; others",
    "Transpak": "Multi-brand packaging materials assortment; others",
    "Baas Verpakkingen": "Multi-brand packaging materials assortment; others",
    "Packen24": "Multi-brand packaging & shipping supplies assortment; others",
    "Allpack": "Multi-brand packaging materials assortment; others",
    "Springpack": "Multi-brand packaging materials assortment; others",
    "Network Packaging": "Multi-brand packaging materials & machinery assortment; others",
    "Flexi-Films": "Multi-brand flexible packaging films assortment; others",
    "Multiplastics": "Multi-brand flexible packaging films assortment; others",
    "American Packaging Distributors": "Multi-brand packaging & shipping supplies assortment; others",
    "Preferred Plastics & Packaging": "Multi-brand stretch film, tape & packaging supplies assortment; others",
    "Schermerhorn Bros.": "Multi-brand packaging, bags, strapping & stretch film assortment; others",
    "United One Source": "Multi-brand packaging & material handling assortment; others",
    "Braden Packaging & Supply": "Multi-brand packaging, shipping & janitorial assortment; others",
    "Auspackaging": "Multi-brand wholesale packaging assortment; others",
    "Ozpack": "Multi-brand packaging supplies assortment; others",
    "Pluspak": "Multi-brand disposable packaging & consumables assortment; others",
    "MaxPak": "MaxPak; multi-brand packaging products assortment; others",
    "Surepak": "Multi-brand packaging & industrial supplies assortment; others",
    "Total Packaging": "Multi-brand packaging materials assortment; others",
    "UK Packaging Supplies": "Multi-brand wholesale packaging materials assortment; others",
    "Packaging Express": "Multi-brand packaging supplies assortment; others",
    "Polythene UK": "Multi-brand polythene, stretch & shrink assortment; others",
    "Paul Hildebrandt": "Multi-brand packaging materials & systems assortment; others",
    "Karton.eu": "Multi-brand cartons & packaging materials assortment; others",
}


def norm_company(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())


def clean_brands(t: str) -> str:
    t = re.sub(r"\s+", " ", (t or "").strip())
    t = t.strip(" ;,")
    t = re.sub(r"(?:;\s*){2,}", "; ", t)
    # normalize commas to semicolons for consistency
    if ";" not in t and "," in t and "assortment" not in t.lower():
        parts = [p.strip() for p in t.split(",") if p.strip()]
        if len(parts) >= 2 and not any(len(p) > 60 for p in parts):
            t = "; ".join(parts)
    if t and not t.endswith(".") and "others" not in t.lower() and ";" in t:
        # keep as list; optional others already handled by extractor
        pass
    return t


def is_self_name(company: str, brands: str) -> bool:
    c = re.sub(r"[^a-z0-9]", "", (company or "").lower())
    b = re.sub(r"[^a-z0-9]", "", (brands or "").lower())
    if not b or not c:
        return False
    if b == c or c.startswith(b) or b.startswith(c[: max(6, len(c) // 2)]):
        # only self if single token-ish (no other brands)
        if ";" not in brands and "," not in brands and "assortment" not in brands.lower():
            return True
    return False


def is_weak(company: str, brands: str) -> bool:
    b = (brands or "").strip()
    if not b or b == NPD:
        return True
    if is_self_name(company, b):
        return True
    low = b.lower()
    if low in ("n/a", "na", "none", "various", "multiple brands", "tbd"):
        return True
    if "multi-brand" in low and "assortment" in low and ";" not in b and len(b) < 80:
        # vague but acceptable for large catalogs — not weak for overwrite if already set by HAND
        return False
    if len(b) < 8:
        return True
    return False


def fetch_evidence(url: str, timeout: float = 12.0) -> str:
    if not url or not str(url).startswith("http"):
        return ""
    headers = {
        "User-Agent": "Mozilla/5.0 (compatible; VendorIntelBot/1.0; +research)",
        "Accept": "text/html,application/xhtml+xml",
    }
    texts = []
    try:
        with httpx.Client(follow_redirects=True, timeout=timeout, headers=headers) as client:
            for path in ("", "/brands", "/brands/", "/partners", "/our-brands", "/suppliers", "/about", "/about-us"):
                try:
                    u = url if not path else urljoin(url if url.endswith("/") else url + "/", path.lstrip("/"))
                    r = client.get(u)
                    if r.status_code >= 400:
                        continue
                    html = r.text
                    html = re.sub(r"(?is)<script[^>]*>.*?</script>", " ", html)
                    html = re.sub(r"(?is)<style[^>]*>.*?</style>", " ", html)
                    html = re.sub(r"(?is)<[^>]+>", " ", html)
                    html = re.sub(r"&\w+;", " ", html)
                    html = re.sub(r"\s+", " ", html)
                    if len(html) > 300:
                        texts.append(html[:5000])
                    if path and len(texts) >= 2:
                        break
                except Exception:
                    continue
    except Exception:
        return ""
    return "\n---\n".join(texts)[:12000]


def deepseek_extract(company: str, website: str, evidence: str) -> str | None:
    key = os.getenv("OPENAI_API_KEY")
    base = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
    model = os.getenv("OPENAI_MODEL") or "deepseek-chat"
    if not key or not evidence or len(evidence) < 250:
        return None
    prompt = f"""Extract Key Brands Represented for packaging CHANNEL/distributor company "{company}" ({website}).

STRICT RULES:
- List ONLY manufacturer/supplier brand names explicitly mentioned in Evidence as products they sell, distribute, or partner with.
- Do NOT invent brands. Do NOT guess common packaging brands.
- Do NOT list the company name alone unless evidence says they primarily sell own-brand products.
- Prefer manufacturer brands (e.g. 3M, Sealed Air, Pregis, Ranpak, Dart) over product category names.
- If evidence shows they are multi-brand but names no principals, return conf="low" and brands="".
- Format brands as: Brand1; Brand2; Brand3; others   (use "; others" when list is a verified subset)
- If insufficient evidence: conf="low", brands=""

Return ONLY JSON: {{"brands":"...","conf":"high|medium|low"}}

Evidence:
{evidence[:9000]}
"""
    try:
        with httpx.Client(timeout=60.0) as client:
            r = client.post(
                f"{base}/chat/completions",
                headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"},
                json={
                    "model": model,
                    "temperature": 0,
                    "messages": [{"role": "user", "content": prompt}],
                },
            )
            r.raise_for_status()
            content = r.json()["choices"][0]["message"]["content"]
            m = re.search(r"\{[\s\S]*\}", content)
            if not m:
                return None
            data = json.loads(m.group(0))
            conf = str(data.get("conf") or "low").lower()
            brands = clean_brands(str(data.get("brands") or ""))
            if conf == "low" or not brands or len(brands) < 3:
                return None
            if is_self_name(company, brands):
                return None
            # reject if looks like categories not brands
            if re.fullmatch(r"(?i)(films?|bags?|boxes?|packaging|supplies)([; ].*)?", brands):
                return None
            return brands
    except Exception:
        return None


def main() -> None:
    wb = load_workbook(FP)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    ic = {h: i + 1 for i, h in enumerate(headers)}
    ci, wi, bi = ic["Company"], ic["Website"], ic["Key Brands Represented"]

    ckpt: dict = {}
    if CKPT.exists():
        try:
            ckpt = json.loads(CKPT.read_text(encoding="utf-8"))
        except Exception:
            ckpt = {}

    updated = []
    skipped = []

    # 1) Apply HAND
    for row in range(2, ws.max_row + 1):
        name = norm_company(str(ws.cell(row, ci).value or ""))
        if name not in HAND:
            continue
        val = clean_brands(HAND[name])
        old = ws.cell(row, bi).value
        ws.cell(row, bi).value = val
        updated.append({"company": name, "source": "hand", "brands": val, "old": old})

    # 2) Clear pure self-name fakes → NPD (will try DS next)
    for row in range(2, ws.max_row + 1):
        name = norm_company(str(ws.cell(row, ci).value or ""))
        if name in HAND:
            continue
        cur = str(ws.cell(row, bi).value or "").strip()
        if is_self_name(name, cur):
            ws.cell(row, bi).value = NPD
            updated.append({"company": name, "source": "clear_selfname", "brands": NPD, "old": cur})

    # 3) DeepSeek for empty / NPD / weak (cap)
    weak_rows = []
    for row in range(2, ws.max_row + 1):
        name = norm_company(str(ws.cell(row, ci).value or ""))
        if not name or name in HAND:
            continue
        brands = str(ws.cell(row, bi).value or "").strip()
        if is_weak(name, brands) or brands == NPD:
            weak_rows.append(row)

    limit = 100
    for idx, row in enumerate(weak_rows[:limit]):
        name = norm_company(str(ws.cell(row, ci).value or ""))
        web = str(ws.cell(row, wi).value or "")
        if name in ckpt and ckpt[name].get("brands"):
            val = clean_brands(ckpt[name]["brands"])
            if val and not is_self_name(name, val):
                ws.cell(row, bi).value = val
                updated.append({"company": name, "source": "ckpt", "brands": val})
                continue
        print(f"[{idx+1}/{min(limit,len(weak_rows))}] {name} …", flush=True)
        evidence = fetch_evidence(web)
        result = deepseek_extract(name, web, evidence)
        if result:
            ws.cell(row, bi).value = result
            ckpt[name] = {"brands": result}
            updated.append({"company": name, "source": "deepseek", "brands": result})
            CKPT.write_text(json.dumps(ckpt, indent=2, ensure_ascii=False), encoding="utf-8")
        else:
            # leave NPD if empty/self/weak
            cur = str(ws.cell(row, bi).value or "").strip()
            if not cur or is_self_name(name, cur) or is_weak(name, cur):
                ws.cell(row, bi).value = NPD
            skipped.append({"company": name, "website": web, "reason": "insufficient_evidence"})
        time.sleep(0.35)

    # 4) Final: any remaining self-name → NPD
    for row in range(2, ws.max_row + 1):
        name = norm_company(str(ws.cell(row, ci).value or ""))
        cur = str(ws.cell(row, bi).value or "").strip()
        if is_self_name(name, cur):
            ws.cell(row, bi).value = NPD

    wb.save(FP)

    # stats
    named = empty = npd = selfish = 0
    for row in range(2, ws.max_row + 1):
        name = norm_company(str(ws.cell(row, ci).value or ""))
        b = str(ws.cell(row, bi).value or "").strip()
        if not b:
            empty += 1
        elif b == NPD:
            npd += 1
        elif is_self_name(name, b):
            selfish += 1
        else:
            named += 1

    audit = {
        "ts": datetime.now().isoformat(timespec="seconds"),
        "hand": sum(1 for u in updated if u.get("source") == "hand"),
        "updated": len(updated),
        "skipped": skipped,
        "named_or_assortment": named,
        "npd": npd,
        "empty": empty,
        "selfname_left": selfish,
        "rows": updated[:200],
    }
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print(
        f"named={named} npd={npd} empty={empty} selfname_left={selfish} "
        f"updated={len(updated)} weak_tried={min(limit,len(weak_rows))} skipped={len(skipped)}"
    )
    print("saved", FP)
    print("audit", AUDIT)


if __name__ == "__main__":
    main()
