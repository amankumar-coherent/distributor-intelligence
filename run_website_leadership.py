#!/usr/bin/env python3
"""Scrape leadership names straight off each company's own website.

Why this exists: a Google/AI Mode search can miss a name that is plainly
published on the company's site. Measured on APCFSS — apcfss.in has a "Top
Management" page naming Sri Sandeep Krupakar Gundala as CEO, yet the AI
Mode pass returned only the Chairman. The company's own site is both the
most authoritative source and the one the search engine is least likely to
summarise for a small firm.

How it works, per company:
  1. start from the known website (or derive one from the company name)
  2. fetch the homepage and follow likely leadership pages — /about,
     /team, /leadership, /management, /board, /our-team, ... plus any link
     whose text mentions leadership/management/team/board/directors
  3. pull "Name — Title" pairs out of the page text and map each title to
     a role column

ONLY REAL DATA: every name comes from the fetched page text; nothing is
inferred or generated. Pages that cannot be fetched leave the row blank.
"""
import argparse
import json
import re
import sys
import time
from pathlib import Path
from urllib.parse import urljoin, urlparse

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT))

ROLES = [
    "MD", "CEO", "COO", "CFO", "COUNTRY HEAD", "CTO",
    "AI HEAD", "MARKETING HEAD", "Founder",
]

# Most specific first: "Co-founder & CTO" must land in CTO, not Founder.
# Every acronym needs \b on both sides. Without it "cto" matches inside
# "director", so "Managing Director" was classified as CTO as well as MD.
ROLE_PATTERNS = [
    ("AI HEAD", r"(\bchief ai officer\b|\bhead of (ai|artificial intelligence|data science|machine learning)\b)"),
    ("CTO", r"(\bcto\b|\bcio\b|\bchief technology officer\b|\bchief technical officer\b|\bchief information officer\b|\bhead of engineering\b|\bvp of engineering\b)"),
    ("CFO", r"(\bcfo\b|\bchief financial officer\b|\bfinance director\b|\bhead of finance\b)"),
    ("COO", r"(\bcoo\b|\bchief operating officer\b|\bhead of operations\b)"),
    ("MARKETING HEAD", r"(\bcmo\b|\bchief marketing officer\b|\bhead of marketing\b|\bmarketing director\b|\bvp of marketing\b)"),
    ("COUNTRY HEAD", r"(\bcountry head\b|\bcountry manager\b|\bcountry director\b|\bindia head\b|\bhead of india\b)"),
    ("MD", r"(\bmanaging director\b|\bmd\b)"),
    ("CEO", r"(\bceo\b|\bchief executive officer\b)"),
    ("Founder", r"(\bco-?founder\b|\bfounder\b)"),
]

OTHER_TITLES = r"(chairman|chairperson|president|director|partner|proprietor|principal|owner|head)"

LEADERSHIP_PATHS = [
    "", "/about", "/about-us", "/about_us", "/aboutus", "/company",
    "/team", "/our-team", "/ourteam", "/leadership", "/our-leadership",
    "/management", "/top-management", "/board", "/board-of-directors",
    "/directors", "/people", "/who-we-are", "/corporate-profile",
    "/#about", "/management-team", "/executive-team",
]

LINK_HINT = re.compile(
    r"(leadership|management|our team|the team|board|directors|about us|who we are|people)",
    re.I,
)

# Words that are never part of a person's name; used to reject headings and
# navigation text that happen to sit next to a job title.
NOT_A_NAME = re.compile(
    r"\b(the|and|our|your|home|about|contact|services|solutions|company|team|"
    r"board|management|leadership|profile|read|more|view|click|welcome|message|"
    r"privacy|policy|terms|careers|login|search|menu|share|copyright|rights|"
    r"reserved|limited|private|pvt|ltd|llp|inc|corporation|technologies|"
    r"consists|currently|employees|divisions|approximately|"
    # Organisation / place words: an entity is not a person. Added after
    # "Government Organizations" and "Finance Dept. A.P. Joint" were
    # extracted as people from apcfss.in.
    r"government|organization|organisation|department|dept|ministry|secretary|"
    r"finance|govt|govern|institute|university|college|school|hospital|"
    r"foundation|society|trust|council|centre|center|bank|group|holdings|"
    r"industries|enterprises|systems|services|international|india|national)\b",
    re.I,
)

HONORIFIC = r"(?:Sri|Shri|Smt\.?|Mr\.?|Mrs\.?|Ms\.?|Dr\.?|Prof\.?|CA\b)"
# A person's name: 2-4 capitalised words, optionally preceded by an
# honorific and followed by a credential such as IAS/IFS.
NAME = rf"(?:{HONORIFIC}\s+)?[A-Z][a-z]+(?:\s+[A-Z][a-zA-Z.']+){{1,3}}"


def classify(title: str) -> list[str]:
    low = (title or "").lower()
    return [role for role, pat in ROLE_PATTERNS if re.search(pat, low)]


def clean_name(raw: str) -> str:
    name = " ".join((raw or "").split())
    name = re.sub(rf"^{HONORIFIC}\s+", "", name).strip(" ,.-—–")
    # Drop trailing credentials: "Sandeep Krupakar Gundala, IFS" -> name
    name = re.sub(r",?\s*\b(IAS|IFS|IPS|IRS|FCA|CA|PhD|MBA|B\.?Tech|M\.?Tech)\b\.?$", "", name).strip(" ,.")
    # A trailing title fragment gets swept in when the name runs straight
    # into its title: "Sudhi Hegde Co" came from "Sudhi Hegde Co-founder".
    name = re.sub(
        r"\s+(Co|Founder|Chief|Managing|Director|President|Chairman|CEO|CTO|CFO|COO|MD)\.?$",
        "",
        name,
    ).strip(" ,.-")
    return name


def looks_like_name(name: str, company: str = "") -> bool:
    if not name or len(name) < 5 or len(name) > 48:
        return False
    if NOT_A_NAME.search(name):
        return False
    words = name.split()
    if not (2 <= len(words) <= 4):
        return False
    if not all(w[:1].isupper() for w in words if w):
        return False

    # Reject run-together marketing copy: a real name has no internal
    # capital after the first letter ("DashboardsVisualize" does).
    for w in words:
        if re.search(r"[a-z][A-Z]", w):
            return False
        if not re.fullmatch(r"[A-Za-z][A-Za-z.'\-]*", w):
            return False

    # Reject job titles, professions and org words appearing INSIDE the
    # captured name. Measured failures: "Chief Executive Officer" recorded
    # as a CTO's name, "Chartered Accountant" as a founder, "Officer Scott
    # Pollak" and "Sandeep Suresh Managing" where a title bled into the
    # name, and "Omega Seiki Mobility" (the company itself) as a CFO.
    TITLE_WORDS = {
        "chief", "executive", "officer", "managing", "director", "founder",
        "cofounder", "co-founder", "president", "chairman", "chairperson",
        "vice", "head", "lead", "leader", "manager", "partner", "principal",
        "proprietor", "owner", "board", "member", "advisor", "adviser",
        "consultant", "consulting", "accountant", "chartered", "engineer",
        "analyst", "specialist", "expert", "acquisition", "talent", "sales",
        "avp", "svp", "evp", "vp", "ceo", "cto", "cfo", "coo", "cmo", "cio",
        "mobility", "club", "industry", "economic", "financial",
    }
    if any(w.lower().strip(".,") in TITLE_WORDS for w in words):
        return False

    # Reject common-noun phrases that are Title Cased on marketing pages
    # ("Mini Tractors", "Sight Mar", "Best Room Selection").
    COMMON = {
        "mini", "best", "room", "selection", "explore", "sight", "village",
        "dynamic", "visualize", "dashboards", "smart", "quality", "premium",
        "welcome", "discover", "learn", "products", "product", "service",
        "solution", "solutions", "features", "feature", "pricing", "demo",
        "support", "resources", "blog", "news", "events", "gallery", "video",
        "new", "top", "great", "free", "online", "digital", "global", "world",
        "book", "order", "shop", "store", "price", "offer", "sale", "buy",
    }
    if any(w.lower() in COMMON for w in words):
        return False

    # A word from the company's own name inside a person's name means the
    # capture ran across a boundary: "Quinnox Anil Kumar" on quinnox.com.
    if company:
        company_words = {
            w.lower()
            for w in re.findall(r"[A-Za-z]{4,}", company)
            if w.lower() not in {"private", "limited", "technologies", "solutions",
                                 "services", "india", "company", "group", "systems"}
        }
        if any(w.lower() in company_words for w in words):
            return False

    # "People", "Team" etc. trailing a real name ("Rakesh Sarin People").
    if words[-1].lower() in {"people", "team", "group", "staff", "board", "members"}:
        return False
    return True


# Acronyms must be whole words and UPPERCASE. Measured failures when they
# were matched case-insensitively as substrings: "cto" inside other words
# produced "Village Barhana (cto)" and "Mini Tractors (cto)", and "cio"
# produced "Best Room Selection Explore (cio)".
TITLE_ACRONYM = r"\b(?:CEO|CTO|CFO|COO|CMO|CIO|MD)\b"
TITLE_PHRASE = (
    r"(?:chairman|chairperson|president|managing director|"
    r"chief\s+[a-z]+\s+officer|co-?founder|founder|country head|"
    r"director|partner|proprietor|principal)"
)


def extract_by_proximity(text: str, company: str = "") -> list[tuple[str, str]]:
    """Pair each title with the nearest plausible name.

    Regex joiners between a name and its title are endlessly variable —
    measured on apcfss.in, the real sentence reads "Sri Sandeep Krupakar
    Gundala, IFS as the CEO", where ", IFS" and "as the" sit in between,
    and a card layout instead separates them with markup and newlines.
    Locating names and titles independently and pairing on character
    distance handles all of those without a pattern per layout.
    """
    names = [
        (m.start(), m.end(), clean_name(m.group(0)))
        for m in re.finditer(NAME, text)
    ]
    names = [(s, e, n) for s, e, n in names if looks_like_name(n, company)]
    if not names:
        return []

    # Acronyms case-sensitively, phrases case-insensitively.
    title_spans = [(m.start(), m.end(), m.group(0)) for m in re.finditer(TITLE_ACRONYM, text)]
    title_spans += [(m.start(), m.end(), m.group(0)) for m in re.finditer(TITLE_PHRASE, text, re.I)]

    pairs: list[tuple[str, str]] = []
    for tstart, tend, title in sorted(title_spans):
        # Nearest name, preferring one that PRECEDES the title ("Name, IFS
        # as the CEO" and "Name – CEO" are far more common than the reverse
        # on About/Team pages). The window is deliberately tight: a wide
        # one pulled marketing copy in as names.
        best = None
        best_score = 10**6
        for s, e, n in names:
            if e <= tstart:
                gap = tstart - e
            elif s >= tend:
                gap = (s - tend) + 25  # penalise title-before-name
            else:
                continue
            if gap < best_score and gap <= 40:
                best_score, best = gap, n
        if best:
            pairs.append((best, title))
    return pairs


def extract_pairs(text: str, company: str = "") -> list[tuple[str, str]]:
    """(name, title) pairs found in page text, in several common layouts."""
    pairs: list[tuple[str, str]] = list(extract_by_proximity(text, company))
    title_word = rf"(?:{OTHER_TITLES}|ceo|cto|cfo|coo|cmo|cio|managing director|chief [a-z ]{{3,30}}officer|founder|co-?founder|country head)"

    # Credentials and honorific suffixes that sit between a name and its
    # title: "Sri Sandeep Krupakar Gundala, IFS as the CEO".
    cred = r"(?:,?\s*(?:IAS|IFS|IPS|IRS|FCA|CA|PhD|MBA)\.?\s*,?)?"
    joiner = r"(?:\s*[,\-–—|:]\s*|\s+(?:is|as|serves as|was appointed|holds the (?:role|position) of)\s+(?:the\s+)?|\s*\n\s*)"

    patterns = [
        # name [, CRED] <joiner> TITLE — covers "…, IFS as the CEO",
        # "Name – CEO", "Name\nCEO" (card layouts), "Name, Managing Director"
        rf"({NAME}){cred}{joiner}(?:the\s+)?({title_word}[A-Za-z /&]{{0,28}})",
        # "CEO: Sandeep Gundala" / "CEO - Sandeep Gundala"
        rf"({title_word}[A-Za-z /&]{{0,28}})\s*[:\-–—]\s*(?:{HONORIFIC}\s+)?({NAME})",
        # "Board Chaired by Sri Peeyush Kumar" / "led by X" — title first,
        # then a linking phrase, then the person.
        rf"({title_word}[A-Za-z /&]{{0,20}})\s+(?:by|is|:)\s+(?:the\s+)?({NAME})",
    ]
    # Patterns 1 and 2 put the title first; normalise every match to
    # (name, title).
    title_first = {1, 2}
    for idx, pattern in enumerate(patterns):
        # Pattern 0 must stay case-sensitive so the NAME group only matches
        # properly capitalised names; the title-first ones need re.I to
        # catch "ceo"/"Chief Executive Officer" alike.
        flags = re.I if idx in title_first else 0
        for match in re.finditer(pattern, text, flags):
            a, b = match.group(1), match.group(2)
            name, title = (b, a) if idx in title_first else (a, b)
            name = clean_name(name)
            if looks_like_name(name, company) and (
                classify(title) or re.search(OTHER_TITLES, title, re.I)
            ):
                pairs.append((name, " ".join(title.split())))
    return pairs


def fetch(url: str, max_chars: int = 9000) -> str:
    try:
        from vendor_intel.scraping.httpx_fetch import fetch_page_httpx

        res = fetch_page_httpx(url, max_chars=max_chars)
        return (res.text or "") if res.alive else ""
    except Exception:  # noqa: BLE001
        return ""


def fetch_html(url: str) -> str:
    """Raw HTML, so in-page links to leadership sections can be followed."""
    try:
        import httpx

        from vendor_intel.clients.browser_headers import browser_headers

        with httpx.Client(follow_redirects=True, timeout=20.0, verify=False) as c:
            r = c.get(url, headers=browser_headers())
            return r.text if r.status_code < 400 else ""
    except Exception:  # noqa: BLE001
        return ""


def candidate_urls(website: str) -> list[str]:
    base = website.strip()
    if not base:
        return []
    if not base.startswith("http"):
        base = "https://" + base
    parsed = urlparse(base)
    root = f"{parsed.scheme}://{parsed.netloc}"
    urls = [root + p for p in LEADERSHIP_PATHS]

    # Follow links whose anchor text mentions leadership/team/board.
    html = fetch_html(root)
    if html:
        for m in re.finditer(r'href=["\']([^"\']+)["\'][^>]*>([^<]{3,60})<', html, re.I):
            href, label = m.group(1), m.group(2)
            if LINK_HINT.search(label) and not href.startswith(("mailto:", "tel:", "#")):
                full = urljoin(root + "/", href)
                if urlparse(full).netloc == parsed.netloc and full not in urls:
                    urls.append(full)
    seen, out = set(), []
    for u in urls:
        if u not in seen:
            seen.add(u)
            out.append(u)
    return out[:14]


def scrape_company(website: str, max_pages: int = 8, company: str = "") -> dict:
    result: dict = {"pages_read": 0, "roles": {}, "other": [], "pairs": []}
    for url in candidate_urls(website)[:max_pages]:
        text = fetch(url)
        if len(text.strip()) < 80:
            continue
        result["pages_read"] += 1
        for name, title in extract_pairs(text, company):
            result["pairs"].append({"name": name, "title": title, "url": url})
            roles = classify(title)
            if roles:
                for role in roles:
                    result["roles"].setdefault(role, {"name": name, "title": title, "url": url})
            elif re.search(OTHER_TITLES, title, re.I):
                result["other"].append({"name": name, "title": title, "url": url})
        if len(result["roles"]) >= len(ROLES):
            break
    return result


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--source", default=str(ROOT / "output/aimode_roles/aimode_roles.json"),
                        help="AI Mode checkpoint providing company names + websites")
    parser.add_argument("--checkpoint", default=str(ROOT / "output/website_scrape/website_roles.json"))
    parser.add_argument("--output", default=str(ROOT / "output/website_scrape/website_leadership.xlsx"))
    parser.add_argument("--limit", type=int, default=None)
    parser.add_argument("--only-missing", action="store_true",
                        help="Only companies still missing a CEO/MD/Founder in the source")
    parser.add_argument("--test", default=None, help="Scrape a single website and print what is found")
    args = parser.parse_args()

    if args.test:
        res = scrape_company(args.test)
        print(f"pages read: {res['pages_read']}")
        for role, info in res["roles"].items():
            print(f"  {role:16} {info['name']}  ({info['title']})")
        for o in res["other"][:6]:
            print(f"  OTHER            {o['name']}  ({o['title']})")
        return

    source = json.loads(Path(args.source).read_text(encoding="utf-8"))
    ckpt = Path(args.checkpoint)
    data = json.loads(ckpt.read_text(encoding="utf-8")) if ckpt.exists() else {}

    todo = []
    for key, entry in source.items():
        site = (entry or {}).get("website") or ""
        if not site:
            continue
        if key in data:
            continue
        if args.only_missing and any((entry.get(r) or {}).get("name") for r in ("CEO", "MD", "Founder")):
            continue
        todo.append((key, site))
    if args.limit:
        todo = todo[: args.limit]

    print(f"companies with a website to scrape: {len(todo)}")
    found = 0
    for i, (key, site) in enumerate(todo, start=1):
        try:
            res = scrape_company(site, company=key)
        except Exception as exc:  # noqa: BLE001
            res = {"pages_read": 0, "roles": {}, "other": [], "error": str(exc)[:80]}
        data[key] = {"website": site, **res}
        if res.get("roles") or res.get("other"):
            found += 1
        if i % 5 == 0 or i == len(todo):
            ckpt.parent.mkdir(parents=True, exist_ok=True)
            ckpt.write_text(json.dumps(data, indent=1, ensure_ascii=False), encoding="utf-8")
            print(f"  [{i}/{len(todo)}] with data: {found}", flush=True)
        time.sleep(0.4)

    ckpt.parent.mkdir(parents=True, exist_ok=True)
    ckpt.write_text(json.dumps(data, indent=1, ensure_ascii=False), encoding="utf-8")
    print(f"\ncompanies scraped: {len(todo)}   with leadership found: {found}")
    print(f"checkpoint: {ckpt}")


if __name__ == "__main__":
    main()
