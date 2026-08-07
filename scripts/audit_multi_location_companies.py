# -*- coding: utf-8 -*-
"""Find same-company / multi-location name families in FINAL Excels.

Example: Alliance Healthcare Portugal + Alliance Healthcare Polska → 1 unique parent.
"""
from __future__ import annotations

import json
import re
import unicodedata
from collections import defaultdict
from pathlib import Path

import openpyxl

MARKETS = [
    {
        "name": "Global Firewall Market",
        "path": Path(
            r"D:\vendor-intel\output\run4\global_firewall_market_global\global_firewall_market_FINAL.xlsx"
        ),
    },
    {
        "name": "Global Green Chemical Market",
        "path": Path(
            r"D:\vendor-intel\output\run4\global_green_chemical_market_global\global_green_chemical_market_FINAL.xlsx"
        ),
    },
    {
        "name": "Global Smartwatch Market",
        "path": Path(
            r"D:\vendor-intel\output\run4\global_smartwatch_market_global\global_smartwatch_market_FINAL.xlsx"
        ),
    },
    {
        "name": "Global Smartphone Market",
        "path": Path(
            r"D:\vendor-intel\output\web_expand\global_smartphone_market_global\global_smartphone_market_global_FINAL.xlsx"
        ),
    },
    {
        "name": "Global Avocado Oil Market",
        "path": Path(
            r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global\global_avocado_oil_market_global_FINAL.xlsx"
        ),
    },
    {
        "name": "Global Antibody Drug Conjugate Market",
        "path": Path(
            r"D:\vendor-intel\output\web_expand\global_antibody_drug_conjugates_market_global\global_antibody_drug_conjugates_market_FINAL.xlsx"
        ),
    },
]

COUNTRIES = {
    "portugal",
    "polska",
    "poland",
    "espana",
    "españa",
    "spain",
    "france",
    "greece",
    "ireland",
    "italia",
    "italy",
    "nederland",
    "netherlands",
    "holland",
    "norge",
    "norway",
    "romania",
    "românia",
    "uk",
    "united kingdom",
    "britain",
    "germany",
    "deutschland",
    "usa",
    "us",
    "u.s.",
    "u.s.a.",
    "united states",
    "america",
    "americas",
    "brazil",
    "brasil",
    "mexico",
    "méxico",
    "canada",
    "australia",
    "india",
    "china",
    "japan",
    "korea",
    "south korea",
    "singapore",
    "malaysia",
    "thailand",
    "indonesia",
    "philippines",
    "vietnam",
    "viet nam",
    "taiwan",
    "hong kong",
    "hongkong",
    "uae",
    "dubai",
    "saudi",
    "saudi arabia",
    "egypt",
    "south africa",
    "nigeria",
    "kenya",
    "turkey",
    "türkiye",
    "israel",
    "switzerland",
    "schweiz",
    "suisse",
    "austria",
    "österreich",
    "oesterreich",
    "belgium",
    "belgie",
    "belgique",
    "sweden",
    "sverige",
    "denmark",
    "danmark",
    "finland",
    "suomi",
    "czech",
    "czechia",
    "czech republic",
    "slovakia",
    "hungary",
    "croatia",
    "serbia",
    "bulgaria",
    "ukraine",
    "russia",
    "chile",
    "argentina",
    "colombia",
    "peru",
    "ecuador",
    "new zealand",
    "nz",
    "pakistan",
    "bangladesh",
    "sri lanka",
    "qatar",
    "kuwait",
    "bahrain",
    "oman",
    "morocco",
    "tunisia",
    "algeria",
    "ghana",
    "ethiopia",
    "luxembourg",
    "liechtenstein",
    "iceland",
    "estonia",
    "latvia",
    "lithuania",
    "slovenia",
    "asia",
    "asia pacific",
    "apac",
    "emea",
    "latam",
    "latin america",
    "middle east",
    "nordics",
    "nordic",
    "europe",
    "eu",
    "anz",
    "mena",
    "gcc",
    "benelux",
    "dach",
    "iberia",
    "scandinavia",
    "espagne",
    "hellas",
    "north america",
    "south america",
    "central america",
    "sea",
    "asean",
    "cis",
    "ksa",
    "uruguay",
    "paraguay",
    "bolivia",
    "venezuela",
    "panama",
    "costa rica",
    "guatemala",
    "dominican republic",
    "puerto rico",
    "jamaica",
    "cambodia",
    "myanmar",
    "laos",
    "mongolia",
    "kazakhstan",
    "uzbekistan",
    "azerbaijan",
    "georgia",
    "armenia",
    "belarus",
    "moldova",
    "albania",
    "macedonia",
    "bosnia",
    "montenegro",
    "kosovo",
    "malta",
    "cyprus",
    "monaco",
    "west africa",
    "east africa",
    "north africa",
    "new south wales",
    "nsw",
    "victoria",
    "queensland",
    "western australia",
    "california",
    "texas",
    "florida",
    "new york",
    "illinois",
    "ontario",
    "quebec",
    "british columbia",
    "sao paulo",
    "são paulo",
    "mexico city",
    "abu dhabi",
    "riyadh",
    "jeddah",
    "london",
    "paris",
    "berlin",
    "madrid",
    "milan",
    "rome",
    "amsterdam",
    "brussels",
    "zurich",
    "munich",
    "frankfurt",
    "hamburg",
    "stockholm",
    "oslo",
    "copenhagen",
    "helsinki",
    "warsaw",
    "prague",
    "vienna",
    "lisbon",
    "athens",
    "dublin",
    "bucharest",
    "budapest",
    "sofia",
    "zagreb",
    "belgrade",
    "tokyo",
    "osaka",
    "seoul",
    "beijing",
    "shanghai",
    "shenzhen",
    "guangzhou",
    "taipei",
    "mumbai",
    "delhi",
    "bangalore",
    "bengaluru",
    "chennai",
    "hyderabad",
    "pune",
    "jakarta",
    "manila",
    "bangkok",
    "kuala lumpur",
    "sydney",
    "melbourne",
    "auckland",
    "toronto",
    "vancouver",
    "montreal",
    "chicago",
    "houston",
    "dallas",
    "atlanta",
    "boston",
    "seattle",
    "denver",
    "phoenix",
    "miami",
    "los angeles",
    "san francisco",
    "san diego",
    "washington",
}

LEGAL_SUFFIX_ONLY = {
    "ltd",
    "limited",
    "llc",
    "inc",
    "corp",
    "corporation",
    "co",
    "gmbh",
    "ag",
    "sa",
    "sas",
    "sarl",
    "bv",
    "nv",
    "oy",
    "ab",
    "as",
    "spa",
    "srl",
    "pty",
    "plc",
    "llp",
    "lp",
    "kk",
    "gk",
    "kft",
    "zrt",
    "doo",
    "holding",
    "group",
    "international",
    "intl",
    "global",
    "s.a.",
    "s.p.a.",
    "s.r.l.",
    "a/s",
    "pty ltd",
    "co ltd",
    "co.",
    "inc.",
    "ltd.",
    "llc.",
    "corp.",
    "bvba",
    "d.o.o.",
    "a.s.",
    "sdn",
    "bhd",
    "sdn bhd",
    "pte",
    "pte ltd",
    "mena",
}

CC2 = {
    "uk",
    "us",
    "ua",
    "ae",
    "sa",
    "nz",
    "za",
    "in",
    "cn",
    "jp",
    "kr",
    "sg",
    "my",
    "th",
    "id",
    "ph",
    "vn",
    "tw",
    "hk",
    "tr",
    "il",
    "ch",
    "at",
    "be",
    "se",
    "dk",
    "fi",
    "cz",
    "sk",
    "hu",
    "hr",
    "rs",
    "bg",
    "ru",
    "cl",
    "ar",
    "co",
    "pe",
    "eg",
    "ng",
    "ke",
    "qa",
    "kw",
    "bh",
    "om",
    "gh",
    "et",
    "pl",
    "pt",
    "es",
    "fr",
    "de",
    "it",
    "nl",
    "no",
    "ro",
    "ie",
    "gr",
    "br",
    "mx",
    "ca",
    "au",
}

GENERIC_STEMS = {
    "best",
    "buy",
    "media",
    "data",
    "soft",
    "tech",
    "global",
    "group",
    "network",
    "networks",
    "arrow",
    "also",
    "star",
    "link",
    "logic",
    "micro",
    "health",
    "care",
    "bio",
    "pharma",
    "chem",
    "chemical",
    "green",
    "smart",
    "watch",
    "phone",
    "mobile",
    "oil",
    "avocado",
    "channel",
    "prime",
    "alliance",
    "power",
    "house",
}

def nk(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(c for c in s if not unicodedata.combining(c)).lower()
    s = s.replace("&", " and ")
    # Keep slash/pipe as separators for brand disambiguation
    s = s.replace("/", " / ").replace("|", " | ")
    s = re.sub(r"[^a-z0-9./|+-]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def company_key(s: str) -> str:
    """Normalized key without slash markers (for uniqueness)."""
    return nk(s).replace(" / ", " ").replace(" | ", " ").strip()


COUNTRY_PHRASES = sorted({nk(c) for c in COUNTRIES if c}, key=len, reverse=True)


def pick_slash_brand(key: str) -> str:
    """For 'Local Brand / Global Brand Country', prefer the global brand side."""
    parts = re.split(r"\s+[\/|]\s+", key)
    if len(parts) < 2:
        return key
    # Prefer part that contains a known multi-country brand token
    preferred_tokens = ("alliance healthcare", "ingram micro", "td synnex", "westcon")
    for p in parts:
        for tok in preferred_tokens:
            if tok in p:
                return p
    return max(parts, key=lambda p: (len(p.split()), len(p)))


def strip_geo_suffix(name: str) -> str:
    key = nk(name)
    if not key:
        return ""
    # Normalize All Eights (M) / (S) style country markers before general strip
    key = re.sub(r"\(([a-z]{1,3})\)", r"\1", key).strip()
    key = re.sub(r"\([^)]{0,40}\)$", "", key).strip()
    key = pick_slash_brand(key)

    changed = True
    while changed and key:
        changed = False
        for c in COUNTRY_PHRASES:
            if key == c:
                return ""
            if key.endswith(" " + c):
                key = key[: -(len(c) + 1)].strip()
                changed = True
                break
        if changed:
            continue
        for leg in sorted(LEGAL_SUFFIX_ONLY, key=len, reverse=True):
            legn = nk(leg)
            if key.endswith(" " + legn):
                key = key[: -(len(legn) + 1)].strip()
                changed = True
                break
        if changed:
            continue
        # trailing single-letter / 2-letter entity country marker: "all eights m"
        m = re.search(r"^(.*) ([a-z]{1,2})$", key)
        if m and len(m.group(1).split()) >= 2:
            tail = m.group(2)
            if tail in CC2 or tail in {"m", "s"}:
                key = m.group(1).strip()
                changed = True
    return key


NON_GEO_REMAINDERS = {
    "commercial",
    "pro audio",
    "specialty",
    "specialty solutions",
    "specialty pharmaceutical distribution",
    "specialty distribution",
    "chef store",
    "chefstore",
    "best",
    "pride",
    "produce",
    "group",
    "group ag",
    "group inc",
    "s p a",
    "spa",
}


def remainder_after_stem(company_nk: str, stem: str) -> str:
    if company_nk == stem:
        return ""
    if company_nk.startswith(stem + " "):
        return company_nk[len(stem) + 1 :].strip()
    # slash names already normalized in stem path
    return company_nk.replace(stem, "").strip(" -–—|/()")


def looks_like_geo_or_legal_remainder(rem: str) -> bool:
    if not rem:
        return True  # bare parent row alongside geo branches
    rem = nk(rem)
    if rem in NON_GEO_REMAINDERS:
        return False
    # product/division words
    if any(
        w in rem.split()
        for w in (
            "commercial",
            "specialty",
            "solutions",
            "audio",
            "chef",
            "store",
            "pride",
            "produce",
            "best",
            "pharmacy",
            "walgreens",
        )
    ):
        return False
    for c in COUNTRY_PHRASES:
        if rem == c or rem.startswith(c + " ") or rem.endswith(" " + c) or c in rem.split():
            return True
    toks = rem.split()
    if len(toks) == 1 and toks[0] in CC2:
        return True
    # Malaysia/Singapore entity markers: "m sdn bhd", "s pte ltd"
    if toks and toks[0] in {"m", "s"} and any(
        t in LEGAL_SUFFIX_ONLY or t in {"sdn", "bhd", "pte"} for t in toks[1:]
    ):
        return True
    if all(
        t in LEGAL_SUFFIX_ONLY or t in {"sdn", "bhd", "pte", "oy", "ab", "as"} or nk(t) in LEGAL_SUFFIX_ONLY
        for t in toks
    ):
        return True
    if rem in {"m", "s", "(m)", "(s)"}:
        return True
    return False


def domain_key(website: str) -> str:
    if not website:
        return ""
    w = str(website).strip().lower()
    w = re.sub(r"^https?://", "", w)
    w = w.split("/")[0].split("?")[0]
    if w.startswith("www."):
        w = w[4:]
    parts = w.split(".")
    if len(parts) >= 3 and parts[-2] in {"co", "com", "net", "org", "gov", "ac"}:
        return ".".join(parts[-3:])
    if len(parts) >= 2:
        return ".".join(parts[-2:])
    return w


def find_header(ws):
    for r in range(1, 6):
        vals = [c.value for c in ws[r]]
        labels = [str(v or "").strip() for v in vals]
        if "Company" in labels:
            cols = {lab: i + 1 for i, lab in enumerate(labels) if lab}
            return r, cols
    raise RuntimeError("no company header")


def load_companies(path: Path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = hdr = cols = None
    for name in wb.sheetnames:
        try:
            hdr, cols = find_header(wb[name])
            ws = wb[name]
            break
        except RuntimeError:
            continue
    if ws is None:
        wb.close()
        raise RuntimeError(f"No company sheet in {path}")
    rows = []
    for r in range(hdr + 1, ws.max_row + 1):
        company = ws.cell(r, cols["Company"]).value
        if not company or not str(company).strip():
            continue
        company = str(company).strip()
        hq = (
            str(ws.cell(r, cols["Headquarters"]).value or "").strip()
            if "Headquarters" in cols
            else ""
        )
        region = (
            str(ws.cell(r, cols["Region Segment"]).value or "").strip()
            if "Region Segment" in cols
            else ""
        )
        website = (
            str(ws.cell(r, cols["Website"]).value or "").strip() if "Website" in cols else ""
        )
        rows.append(
            {
                "company": company,
                "hq": hq,
                "region": region,
                "website": website,
                "stem": strip_geo_suffix(company),
                "dom": domain_key(website),
                "nk": company_key(company),
            }
        )
    wb.close()
    return rows


def find_families(rows):
    by_stem = defaultdict(list)
    for row in rows:
        stem = row["stem"]
        if not stem or len(stem) < 3:
            continue
        tokens = stem.split()
        if len(tokens) == 1 and len(stem) < 5:
            continue
        by_stem[stem].append(row)

    families = []
    for stem, members in by_stem.items():
        uniq = {}
        for m in members:
            uniq[m["nk"]] = m
        if len(uniq) < 2:
            continue
        if len(stem.split()) == 1 and stem in GENERIC_STEMS:
            continue

        # Keep only members whose name remainder is geo/legal (Alliance Healthcare pattern)
        geo_members = []
        for m in uniq.values():
            rem = remainder_after_stem(m["nk"], stem)
            # also try after slash normalization
            rem2 = remainder_after_stem(nk(pick_slash_brand(m["nk"])), stem)
            if looks_like_geo_or_legal_remainder(rem) or looks_like_geo_or_legal_remainder(rem2):
                geo_members.append(m)
        # Need at least 2 variants AND at least one clear geo remainder
        if len(geo_members) < 2:
            continue
        rems = [remainder_after_stem(nk(pick_slash_brand(m["nk"])), stem) for m in geo_members]
        clear_geo = 0
        for rem in rems:
            remn = nk(rem)
            if not remn:
                continue
            if any(remn == c or c in remn.split() for c in COUNTRY_PHRASES) or (
                len(remn.split()) == 1 and remn in CC2
            ):
                clear_geo += 1
            elif remn in {"m", "s"}:  # All Eights (M)/(S)
                clear_geo += 1
            elif any(t in LEGAL_SUFFIX_ONLY or t in {"sdn", "bhd", "pte", "oy"} for t in remn.split()):
                # legal-only family (2BScientific GmbH/Ltd, Nordic Biosite AB/AS/Oy) — count
                clear_geo += 1
        if clear_geo < 1 and not any(r for r in rems):
            continue
        # Drop pure legal-suffix-only duplicates if no geo at all (Computer Gross vs S.p.A.)
        # Still keep GmbH/Ltd cross-country legal forms when >=2 different legal endings
        legal_only = all(
            (not r)
            or all(t in LEGAL_SUFFIX_ONLY or t in {"sdn", "bhd", "pte", "oy", "ab", "as"} for t in nk(r).split())
            for r in rems
        )
        has_geo = any(
            any(nk(r) == c or c in nk(r).split() for c in COUNTRY_PHRASES)
            or (len(nk(r).split()) == 1 and nk(r) in CC2)
            or nk(r) in {"m", "s"}
            for r in rems
            if r
        )
        if legal_only and not has_geo:
            # keep only if different country-style legal entities (oy/ab/as/gmbh/ltd)
            endings = {nk(r) for r in rems if r}
            if len(endings) < 2:
                continue

        families.append(
            {
                "parent_stem": stem,
                "variant_count": len(geo_members),
                "variants": sorted((m["company"] for m in geo_members), key=str.lower),
                "hqs": sorted({m["hq"] for m in geo_members if m["hq"]}),
                "regions": sorted({m["region"] for m in geo_members if m["region"]}),
                "domains": sorted({m["dom"] for m in geo_members if m["dom"]}),
                "match_type": "geo_name" if has_geo else "legal_entity_form",
            }
        )

    families.sort(key=lambda f: (-f["variant_count"], f["parent_stem"]))
    return families


def main():
    out_dir = Path(r"D:\vendor-intel\output\run4\_multi_location_company_audit")
    out_dir.mkdir(parents=True, exist_ok=True)

    summary = []
    all_report = {}

    for m in MARKETS:
        rows = load_companies(m["path"])
        fams = find_families(rows)
        rows_in_fams = sum(f["variant_count"] for f in fams)
        summary.append(
            {
                "market": m["name"],
                "total_companies": len(rows),
                "unique_multi_location_parents": len(fams),
                "total_location_variant_rows": rows_in_fams,
                "file": str(m["path"]),
            }
        )
        all_report[m["name"]] = fams
        print("=" * 80)
        print(m["name"])
        print(f"  Total rows: {len(rows)}")
        print(f"  Unique multi-location parent companies: {len(fams)}")
        print(f"  Rows that are location variants: {rows_in_fams}")
        print("  Families:")
        for f in fams:
            print(f"    • {f['parent_stem']}  ({f['variant_count']} variants)")
            for v in f["variants"]:
                print(f"        - {v}")

    print("\n" + "=" * 80)
    print("SUMMARY COUNTS")
    print(f"{'Market':45} {'Total':>6} {'Unique parents':>15} {'Variant rows':>13}")
    for s in summary:
        print(
            f"{s['market']:45} {s['total_companies']:6} "
            f"{s['unique_multi_location_parents']:15} {s['total_location_variant_rows']:13}"
        )

    (out_dir / "multi_location_families.json").write_text(
        json.dumps({"summary": summary, "families": all_report}, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )

    lines = [
        "# Multi-location company families (same name, many locations)\n",
        "Pattern example: Alliance Healthcare Portugal / Alliance Healthcare Polska → 1 unique parent.\n",
        "## Summary\n",
        "| Market | Total companies | Unique multi-location parents | Location-variant rows |",
        "|---|---:|---:|---:|",
    ]
    for s in summary:
        lines.append(
            f"| {s['market']} | {s['total_companies']} | "
            f"{s['unique_multi_location_parents']} | {s['total_location_variant_rows']} |"
        )
    lines.append("")
    for market, fams in all_report.items():
        lines.append(f"## {market} ({len(fams)} unique parents)\n")
        if not fams:
            lines.append("_None found._\n")
            continue
        for f in fams:
            lines.append(f"### {f['parent_stem']} — {f['variant_count']} variants\n")
            for v in f["variants"]:
                lines.append(f"- {v}")
            if f["hqs"]:
                lines.append(f"\n_HQ values:_ {'; '.join(f['hqs'][:12])}")
            lines.append("")

    (out_dir / "multi_location_families.md").write_text("\n".join(lines), encoding="utf-8")
    print(f"\nWrote: {out_dir / 'multi_location_families.md'}")
    print(f"Wrote: {out_dir / 'multi_location_families.json'}")


if __name__ == "__main__":
    main()
