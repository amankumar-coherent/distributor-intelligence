"""Scrape brand lists from Google AI Overview for each keyword in the keyword CSV.

Usage:
    python scrape_brands.py --keywords ../nutraceutical_canonical_names.csv \
        --out out/brands.csv --start 0 --limit 10

Requires the google-ai-scraper FastAPI backend running (default http://localhost:15551)
with the Chrome extension connected (check via /health).
"""

import argparse
import csv
import json
import re
import sys
import time
import urllib.parse
import urllib.request
import urllib.error
from pathlib import Path

SERVER_URL = "http://localhost:15551"
QUERY_TEMPLATE = (
    "list of 100 {keyword} brands globally and provide for each brand: "
    "brand name, official website, official address. Give the results as a list only"
)
REQUEST_TIMEOUT = 100
RETRY_COUNT = 2
RETRY_DELAY = 5
CAPTCHA_WAIT_SECONDS = 60
MIN_ACCEPTABLE_BRANDS = 10

CAPTCHA_MARKERS = (
    "unusual traffic",
    "detected unusual traffic",
    "/sorry/index",
    "our systems have detected",
)

NOISE_LINE_PATTERNS = [
    r"^would you like to",
    r"^show (less|all|more)$",
    r"^table_title:",
    r"^\* ",
]


def http_get_json(url: str, timeout: int) -> dict:
    req = urllib.request.Request(url, headers={"User-Agent": "brand-scraper/1.0"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))


def ask(query: str) -> dict:
    q = urllib.parse.urlencode({"q": query})
    url = f"{SERVER_URL}/ask?{q}"
    last_err = None
    for attempt in range(RETRY_COUNT + 1):
        try:
            return http_get_json(url, REQUEST_TIMEOUT)
        except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError) as e:
            last_err = str(e)
            if attempt < RETRY_COUNT:
                time.sleep(RETRY_DELAY)
    return {"error": last_err, "markdown": "", "citations": []}


_BRAND_STOPWORDS = {
    "the", "inc", "llc", "ltd", "co", "corp", "corporation", "company",
    "official", "site", "website", "brands", "brand", "group", "labs",
    "laboratories", "health", "healthcare", "nutrition", "vitamins",
    "supplements", "products", "international", "global", "usa", "india",
}


def _brand_name_tokens(name: str) -> set[str]:
    words = re.findall(r"[a-z0-9]+", name.lower())
    return {w for w in words if len(w) >= 3 and w not in _BRAND_STOPWORDS}


def url_matches_brand(url: str, brand_name: str) -> bool:
    """Reject a hyperlink href that doesn't plausibly belong to the brand.

    AI Mode sometimes renders "Official Website" as a bare link whose href
    is actually an unrelated citation/footnote (a review article, social
    post, data-broker profile, etc.) rather than the brand's own domain.
    Recording that as the website would be worse than recording nothing —
    so require at least one significant brand-name token to appear in the
    URL's host.
    """
    try:
        host = re.sub(r"^https?://", "", url).split("/")[0].lower()
        host = re.sub(r"^www\.", "", host)
    except Exception:
        return False

    tokens = _brand_name_tokens(brand_name)
    if not tokens:
        return True  # nothing meaningful to check against; don't over-reject
    host_compact = re.sub(r"[^a-z0-9]", "", host)
    return any(tok in host_compact for tok in tokens)


def parse_brands(markdown: str) -> list[dict]:
    """Parse the AI Overview markdown into brand rows.

    Observed format per brand block:
        Brand Name
        - Official Website:
        LinkText
        , Official Address: <address>.
    Citation/source-card noise appears after the last real brand entry
    (repeated link titles, "Would you like to...", "Show all", table dumps).
    """
    text = markdown or ""

    # Cut off the trailing citation/source-card dump and follow-up prompt.
    cutoff_markers = [
        r"\nWould you like\b",
        r"\nIf you require\b",
        r"\nIf you need\b",
        r"\nShow less\b",
        r"\nShow all\b",
        r"\bDue to character limits\b",
        r"\bunfeasible due to\b",
    ]
    cut_idx = len(text)
    for pat in cutoff_markers:
        m = re.search(pat, text)
        if m:
            cut_idx = min(cut_idx, m.start())
    body = text[:cut_idx]

    # The literal word "website" is the one stable anchor across every
    # observed format ("– Official Website:", ": Official website at",
    # "- Website:", etc.) — split the body into per-brand blocks on lines
    # containing it, rather than trying to match one rigid separator pattern.
    lines = body.split("\n")
    website_line_idxs = [i for i, ln in enumerate(lines) if re.search(r"\bwebsite\b", ln, re.I)]

    rows = []
    for n, idx in enumerate(website_line_idxs):
        next_idx = website_line_idxs[n + 1] if n + 1 < len(website_line_idxs) else len(lines)
        block_lines = lines[idx:next_idx]

        website_line = lines[idx].strip()
        # Same-line format: "N. Brand | Official Website (url) | Address: ..."
        # or "Brand - Website: ... - Address: ..." — the brand name sits
        # before the website marker on the SAME line, not a prior line.
        same_line_match = re.match(
            r"^\s*(?:\d+[\.\)]\s*)?(.+?)\s*[\|–—:-]\s*(?=.*\bwebsite\b)",
            website_line,
            re.I,
        )
        name = ""
        if same_line_match:
            candidate_name = same_line_match.group(1).strip()
            # Only treat as same-line brand name if it doesn't itself look
            # like a full previous brand block (i.e. doesn't contain a
            # website/address marker already).
            if candidate_name and not re.search(r"\bwebsite\b|\baddress\b", candidate_name, re.I):
                name = candidate_name

        if not name:
            # Multi-line format: brand name is on the nearest non-empty
            # line before the website marker, back to the previous block.
            for j in range(idx - 1, -1, -1):
                candidate = lines[j].strip()
                if candidate:
                    name = candidate
                    break
                if j <= (website_line_idxs[n - 1] if n > 0 else -1):
                    break
        name = re.sub(r"^\s*\d+[\.\)]\s*", "", name)
        name = name.rstrip(":").strip()

        block_text = "\n".join(block_lines)
        # Website label: text between the "website" marker and the
        # "address" marker (case-insensitive), stripping separators and
        # any leading "Official"/punctuation immediately before "address".
        # Bounded separator before "address" — a handful of punctuation/space
        # chars at most, NOT an unbounded [\W_]* (which would greedily eat
        # backward through an unrelated trailing ")" from a URL earlier on
        # the same line, corrupting the website field's parenthesized URL).
        addr_split = re.search(r"[\s|,;:.–—-]{0,6}(?:Official\s+)?address\b\s*:?\s*", block_text, re.I)
        if not addr_split:
            continue
        website_part = block_text[: addr_split.start()]
        website_part = re.sub(r".*?\bwebsite\b(?: at)?\s*:?\s*", "", website_part, flags=re.I | re.DOTALL)
        # A real URL, when available, is more useful than a bare link label
        # (e.g. AI Mode inlines "Official Website (https://...)" since the
        # visible link text is often just the words "Official Website").
        # But that href can point to an unrelated citation instead of the
        # brand's own domain — only trust it if it plausibly matches.
        url_match = re.search(r"\((https?://[^\s()]+)\)", website_part)
        if url_match and url_matches_brand(url_match.group(1), name):
            website = url_match.group(1)
        else:
            label = re.sub(r"\s*\(https?://[^\s()]+\)", "", website_part)
            website = re.sub(r"^[\W_]+|[\W_]+$", "", label.replace("\n", " ")).strip()

        raw_address = block_text[addr_split.end():].strip()

        # The address capture can run past the sentence into a citation-echo
        # line (source card title repeated, e.g. "...USA.\nThomasnet").
        addr_match = re.match(r"(.*?\.)\s*\n", raw_address, re.DOTALL)
        address = addr_match.group(1) if addr_match else raw_address
        address = address.rstrip(".").strip()

        # Drop a trailing echo line with no digits/postal content (e.g. a
        # bare repeated source-card title like "Nutraceuticals World").
        addr_lines = address.split("\n")
        while len(addr_lines) > 1 and not re.search(r"\d", addr_lines[-1]):
            addr_lines.pop()
        address = " ".join(line.strip() for line in addr_lines)
        address = re.sub(r"\s+", " ", address).strip()

        if not name or is_noise(name) or not website or not address:
            continue
        rows.append({"brand_name": name, "website": website, "address": address})

    if not rows:
        rows = parse_pipe_table_brands(body)
    if not rows:
        rows = parse_unlabeled_pipe_line_brands(body)
    return rows


def parse_unlabeled_pipe_line_brands(body: str) -> list[dict]:
    """Fallback for single-line pipe-delimited entries with no "Address"
    label at all — just a trailing unlabeled address after the website
    marker:
        Brand Name | Official Website [(https://...)] | 123 Main St, City, USA
    (AI Mode sometimes omits "Address:" entirely, unlike classic AI
    Overview which nearly always labels it.)
    """
    rows = []
    line_pattern = re.compile(
        r"^\s*(?:\d+[\.\)]\s*)?(?P<name>[^|\n]+?)\s*\|\s*"
        r"(?:Official\s+)?[Ww]ebsite(?:\s*\((?P<url>https?://[^\s()]+)\))?\s*\|\s*"
        r"(?P<address>[^|\n]+)\s*$"
    )
    for raw_line in body.split("\n"):
        m = line_pattern.match(raw_line.strip())
        if not m:
            continue
        name = m.group("name").strip()
        website = m.group("url") or ""
        address = m.group("address").strip().rstrip(".").strip()
        if not name or is_noise(name) or not address or not re.search(r"\d", address):
            continue
        if not website or not url_matches_brand(website, name):
            continue  # no usable website value, or href looks unrelated to the brand
        rows.append({"brand_name": name, "website": website, "address": address})
    return rows


def parse_pipe_table_brands(body: str) -> list[dict]:
    """Fallback for the flattened-markdown-table format with no "Website"/
    "Address" words at all, observed as repeating triplets of lines:
        Brand Name |
        Website Label
        | City, State/Region, Country
    (Turndown's rendering of a markdown table row split across lines.)
    """
    lines = [ln.strip() for ln in body.split("\n")]
    pattern = re.compile(
        r"^(?P<name>.+?)\s*\|$"  # "Brand Name |"
    )
    rows = []
    i = 0
    while i < len(lines) - 2:
        m = pattern.match(lines[i])
        if m and lines[i + 1] and not lines[i + 1].endswith("|") and lines[i + 2].startswith("|"):
            name = m.group("name").strip()
            website = lines[i + 1].strip()
            address = lines[i + 2].lstrip("|").strip()
            if name and not is_noise(name) and website and address:
                rows.append({"brand_name": name, "website": website, "address": address})
            i += 3
        else:
            i += 1
    return rows


def is_captcha(markdown: str) -> bool:
    low = (markdown or "").lower()
    return any(marker in low for marker in CAPTCHA_MARKERS)


def is_noise(line: str) -> bool:
    low = line.strip().lower()
    return any(re.match(pat, low) for pat in NOISE_LINE_PATTERNS)


def load_keywords(path: str) -> list[str]:
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        col = reader.fieldnames[0]
        return [row[col].strip() for row in reader if row[col].strip()]


def load_done_keywords(checkpoint_path: Path) -> set[str]:
    if not checkpoint_path.exists():
        return set()
    return {line.strip() for line in checkpoint_path.read_text(encoding="utf-8").splitlines() if line.strip()}


def load_weak_keywords(weak_path: Path) -> list[str]:
    if not weak_path.exists():
        return []
    kws = []
    for line in weak_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line:
            kws.append(line.split("\t", 1)[0])
    return kws


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--keywords", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--start", type=int, default=0)
    ap.add_argument("--limit", type=int, default=None)
    ap.add_argument("--sleep", type=float, default=2.0, help="seconds between queries")
    ap.add_argument("--retry-weak", action="store_true", help="Re-query only keywords logged in <out>.weak.txt")
    args = ap.parse_args()

    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    checkpoint_path = out_path.with_suffix(".done.txt")
    errors_path = out_path.with_suffix(".errors.log")
    weak_path = out_path.with_suffix(".weak.txt")

    if args.retry_weak:
        subset = load_weak_keywords(weak_path)
        weak_path.write_text("", encoding="utf-8")  # cleared; re-populated with any still-weak results
        print(f"Retrying {len(subset)} weak keyword(s) from {weak_path}")
    else:
        keywords = load_keywords(args.keywords)
        subset = keywords[args.start:]
        if args.limit is not None:
            subset = subset[: args.limit]

    done = load_done_keywords(checkpoint_path)
    write_header = not out_path.exists() or out_path.stat().st_size == 0

    csv_file = open(out_path, "a", newline="", encoding="utf-8")
    writer = csv.DictWriter(csv_file, fieldnames=["keyword", "brand_name", "website", "address"])
    if write_header:
        writer.writeheader()

    checkpoint_file = open(checkpoint_path, "a", encoding="utf-8")
    errors_file = open(errors_path, "a", encoding="utf-8")
    weak_file = open(weak_path, "a", encoding="utf-8")

    total = len(subset)
    for i, kw in enumerate(subset, 1):
        if kw in done:
            print(f"[{i}/{total}] SKIP (done) {kw}")
            continue

        query = QUERY_TEMPLATE.format(keyword=kw)
        print(f"[{i}/{total}] QUERY: {kw}")

        while True:
            result = ask(query)
            markdown = result.get("markdown", "")
            if result.get("error"):
                break
            if is_captcha(markdown):
                print(f"    CAPTCHA detected — solve it in the browser, waiting {CAPTCHA_WAIT_SECONDS}s to retry...")
                time.sleep(CAPTCHA_WAIT_SECONDS)
                continue
            break

        if result.get("error"):
            print(f"    ERROR: {result['error']}")
            errors_file.write(f"{kw}\t{result['error']}\n")
            errors_file.flush()
            continue

        brands = parse_brands(markdown)
        print(f"    -> {len(brands)} brands")

        if len(brands) < MIN_ACCEPTABLE_BRANDS:
            # Weak result: don't write rows or mark done, so a later retry
            # starts clean instead of duplicating a partial row set.
            debug_dir = out_path.parent / "raw_debug"
            debug_dir.mkdir(exist_ok=True)
            safe_kw = re.sub(r"[^\w-]+", "_", kw)[:60]
            (debug_dir / f"{safe_kw}.md").write_text(markdown, encoding="utf-8")
            weak_file.write(f"{kw}\t{len(brands)}\n")
            weak_file.flush()
            print(f"    -> below {MIN_ACCEPTABLE_BRANDS}-brand threshold, logged for retry (not written, not marked done)")
        else:
            for b in brands:
                writer.writerow({"keyword": kw, **b})
            csv_file.flush()

            checkpoint_file.write(kw + "\n")
            checkpoint_file.flush()

        time.sleep(args.sleep)

    csv_file.close()
    checkpoint_file.close()
    errors_file.close()
    weak_file.close()
    print("Done.")


if __name__ == "__main__":
    main()
