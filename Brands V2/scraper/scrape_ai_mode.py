"""Scrape brand lists from Google AI Mode (udm=50) via Playwright, with a
fallback to the AI Overview extension pipeline (scrape_brands.py) when AI
Mode fails, is empty, or is CAPTCHA-blocked.

AI Mode returns much richer results than classic AI Overview (real website
URLs, full addresses, up to ~100 brands in one query) but Chrome's own
omnibox intercepts udm=50 into an unreadable chrome://contextual-tasks page.
Playwright navigating directly to the URL bypasses that intercept entirely
(confirmed: Google's server still returns real, scrapable AI Mode HTML).

Usage:
    python scrape_ai_mode.py --keywords ../nutraceutical_canonical_names.csv \
        --out out/brands_full.csv --start 0 --limit 10

Requires: playwright (pip), `playwright install chromium` already run.
Falls back to scrape_brands.ask()/parse_brands() which requires the
google-ai-scraper backend + extension running (see scrape_brands.py).
"""

import argparse
import csv
import re
import time
from pathlib import Path

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeoutError

import scrape_brands as fallback

QUERY_TEMPLATE = fallback.QUERY_TEMPLATE
MIN_ACCEPTABLE_BRANDS = fallback.MIN_ACCEPTABLE_BRANDS
CAPTCHA_WAIT_SECONDS = fallback.CAPTCHA_WAIT_SECONDS

# Cloned copy of the real Chrome profile (Cookies/Login Data/history), since
# Chrome refuses CDP remote debugging against the literal default profile
# directory ("DevTools remote debugging requires a non-default data
# directory"). Using a clone with real history/cookies avoids the immediate
# CAPTCHA a brand-new, empty automated profile triggers.
PROFILE_DIR = str(Path(__file__).parent / ".pw-chrome-profile")
CHROME_PROFILE_DIRECTORY = "Default"
RESPONSE_TIMEOUT_MS = 90_000
THINKING_POLL_MS = 1500


def build_ai_mode_url(keyword: str) -> str:
    import urllib.parse

    query = QUERY_TEMPLATE.format(keyword=keyword)
    q = urllib.parse.quote_plus(query)
    return f"https://www.google.com/search?q={q}&udm=50&hl=en"


def extract_answer_text(page) -> str:
    """Grab the AI Mode answer as text, with link hrefs inlined.

    AI Mode's DOM structure varies between responses (flat <li><strong>...
    triplets in some, numbered <ol><b>Official Website:</b> lists in
    others), so instead of chasing every DOM shape we walk the rendered
    text a human would see and hand it to the same multi-format text
    parser used for classic AI Overview (scrape_brands.parse_brands),
    which already tolerates "Website:"/"Official website at"/pipe-table/
    etc. variations.

    Critically, the "Official Website" link text often has NO visible URL
    at all (the domain only exists in the <a href>) — plain innerText
    would silently drop it. So instead of reading innerText directly, we
    clone the answer container and rewrite every <a href="https://...">
    into "linktext (https://...)" before extracting text, giving the
    downstream parser a real URL to find even when the link label is just
    the words "Official Website".
    """
    try:
        return page.evaluate(
            """
            () => {
              const main = document.querySelector('main') || document.body;
              // Rewrite links in place (not a detached clone — innerText
              // needs real layout info, and a cloneNode(true) that's not
              // attached to the document falls back to raw textContent,
              // which leaks <script>/<style> source into the output).
              // Restore original link text afterward so the live page is
              // left untouched.
              const links = Array.from(main.querySelectorAll('a[href^="http"]'))
                .filter(a => !(a.getAttribute('href') || '').includes('google.com'));
              const originals = links.map(a => a.textContent);
              for (const a of links) {
                const href = a.getAttribute('href');
                a.textContent = `${a.textContent.trim()} (${href})`;
              }
              const scripts = main.querySelectorAll('script, style');
              const scriptDisplay = Array.from(scripts).map(el => el.style.display);
              scripts.forEach(el => { el.style.display = 'none'; });

              const text = main.innerText || '';

              scripts.forEach((el, i) => { el.style.display = scriptDisplay[i]; });
              links.forEach((a, i) => { a.textContent = originals[i]; });

              return text;
            }
            """
        )
    except Exception:
        return ""


def extract_brands(page, debug_dump_path: str | None = None) -> list[dict]:
    """Pull {brand_name, website, address} from AI Mode's rendered answer text."""
    text = extract_answer_text(page)
    if debug_dump_path:
        Path(debug_dump_path).write_text(text, encoding="utf-8")
    return fallback.parse_brands(text)


def page_is_captcha(page) -> bool:
    # Check the Google CAPTCHA/"unusual traffic" interstitial specifically —
    # its URL path is /sorry/, and its body text (not raw HTML, which can
    # false-positive on script/comment content) contains the block notice.
    try:
        url = page.url or ""
    except Exception:
        url = ""
    if "/sorry/" in url:
        return True
    try:
        text = page.inner_text("body").lower()
    except Exception:
        return False
    return any(m in text for m in fallback.CAPTCHA_MARKERS)


def wait_for_response(page) -> None:
    """Wait for AI Mode to finish generating.

    Google shows a "Thinking..." indicator while generating, but it can take
    a moment to appear after navigation — polling too early sees zero
    "Thinking" matches and would wrongly conclude the response is already
    done. So: wait for the indicator to appear at least once (bounded by a
    short grace period, since very fast/cached answers may skip it
    entirely), then wait for it to disappear.
    """
    deadline = time.monotonic() + RESPONSE_TIMEOUT_MS / 1000
    seen_thinking = False
    while time.monotonic() < deadline:
        try:
            thinking = page.get_by_text(re.compile(r"Thinking", re.I)).count()
        except Exception:
            thinking = 0
        if thinking > 0:
            seen_thinking = True
        elif seen_thinking:
            return
        elif has_answer_content(page):
            # No "Thinking" indicator was ever seen, but real answer
            # content has already rendered (fast/cached response) — done.
            return
        page.wait_for_timeout(THINKING_POLL_MS)


def has_answer_content(page) -> bool:
    """True once the answer area has more than just the echoed query."""
    try:
        main = page.locator("main").first
        text = main.inner_text(timeout=1000)
    except Exception:
        return False
    return len(text.strip()) > 400


def scrape_ai_mode(page, keyword: str) -> tuple[list[dict], str]:
    """Returns (brands, status) where status is one of:
    "ok", "captcha", "empty", "error"
    """
    url = build_ai_mode_url(keyword)
    try:
        page.goto(url, wait_until="domcontentloaded", timeout=30_000)
    except PWTimeoutError:
        return [], "error"
    except Exception as e:
        # Google's AI Mode page client-side-redirects itself right after
        # load (adds a &sei= session param) — Playwright's goto() can raise
        # when that redirect interrupts its own navigation wait. Benign:
        # the page still ends up in the right place, so just continue.
        if "interrupted by another navigation" not in str(e):
            return [], "error"

    # Let the redirect/initial render settle before judging the page.
    page.wait_for_timeout(1500)

    if page_is_captcha(page):
        return [], "captcha"

    wait_for_response(page)

    if page_is_captcha(page):
        return [], "captcha"

    debug_dir = Path(__file__).parent / "out" / "raw_debug_aimode"
    debug_dir.mkdir(parents=True, exist_ok=True)
    safe_kw = re.sub(r"[^\w-]+", "_", keyword)[:60]
    brands = extract_brands(page, debug_dump_path=str(debug_dir / f"{safe_kw}.txt"))
    if not brands:
        try:
            (debug_dir / f"{safe_kw}.html").write_text(page.content(), encoding="utf-8")
            page.screenshot(path=str(debug_dir / f"{safe_kw}.png"), full_page=True)
        except Exception:
            pass
        return [], "empty"
    return brands, "ok"


def run_keyword(page, kw: str) -> tuple[list[dict], str]:
    """Try AI Mode; on captcha/empty/error, fall back to AI Overview pipeline."""
    brands, status = scrape_ai_mode(page, kw)
    if status == "captcha":
        return brands, "captcha"
    if len(brands) >= MIN_ACCEPTABLE_BRANDS:
        return brands, "ai_mode"

    # Fallback: classic AI Overview via the extension/backend pipeline.
    query = fallback.QUERY_TEMPLATE.format(keyword=kw)
    result = fallback.ask(query)
    if result.get("error"):
        return brands, "ai_mode_weak"  # keep whatever AI Mode found, if any
    fb_markdown = result.get("markdown", "")
    if fallback.is_captcha(fb_markdown):
        return brands, "ai_mode_weak"
    fb_brands = fallback.parse_brands(fb_markdown)

    # Use whichever source found more.
    if len(fb_brands) > len(brands):
        return fb_brands, "ai_overview_fallback"
    return brands, "ai_mode_weak"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--keywords", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--start", type=int, default=0)
    ap.add_argument("--limit", type=int, default=None)
    ap.add_argument("--sleep", type=float, default=6.0, help="seconds between keywords")
    ap.add_argument("--headless", action="store_true", help="run browser headless (default: visible, so CAPTCHAs can be solved manually)")
    ap.add_argument("--retry-weak", action="store_true", help="Re-query only keywords logged in <out>.weak.txt")
    args = ap.parse_args()

    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    checkpoint_path = out_path.with_suffix(".done.txt")
    weak_path = out_path.with_suffix(".weak.txt")
    sources_path = out_path.with_suffix(".sources.log")

    if args.retry_weak:
        subset = fallback.load_weak_keywords(weak_path)
        weak_path.write_text("", encoding="utf-8")
        print(f"Retrying {len(subset)} weak keyword(s) from {weak_path}")
    else:
        keywords = fallback.load_keywords(args.keywords)
        subset = keywords[args.start:]
        if args.limit is not None:
            subset = subset[: args.limit]

    done = fallback.load_done_keywords(checkpoint_path)
    write_header = not out_path.exists() or out_path.stat().st_size == 0

    csv_file = open(out_path, "a", newline="", encoding="utf-8")
    writer = csv.DictWriter(csv_file, fieldnames=["keyword", "brand_name", "website", "address"])
    if write_header:
        writer.writeheader()

    checkpoint_file = open(checkpoint_path, "a", encoding="utf-8")
    weak_file = open(weak_path, "a", encoding="utf-8")
    sources_file = open(sources_path, "a", encoding="utf-8")

    with sync_playwright() as pw:
        context = pw.chromium.launch_persistent_context(
            PROFILE_DIR,
            channel="chrome",
            headless=args.headless,
            viewport={"width": 1280, "height": 900},
            args=[f"--profile-directory={CHROME_PROFILE_DIRECTORY}"],
        )
        pages = context.pages
        page = pages[0] if pages else context.new_page()

        total = len(subset)
        for i, kw in enumerate(subset, 1):
            if kw in done:
                print(f"[{i}/{total}] SKIP (done) {kw}")
                continue

            print(f"[{i}/{total}] QUERY: {kw}")

            while True:
                brands, status = run_keyword(page, kw)
                if status == "captcha":
                    print(f"    CAPTCHA detected in AI Mode — solve it in the browser window, waiting {CAPTCHA_WAIT_SECONDS}s to retry...")
                    time.sleep(CAPTCHA_WAIT_SECONDS)
                    continue
                break

            print(f"    -> {len(brands)} brands (source: {status})")
            sources_file.write(f"{kw}\t{status}\t{len(brands)}\n")
            sources_file.flush()

            if len(brands) < MIN_ACCEPTABLE_BRANDS:
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

        context.close()

    csv_file.close()
    checkpoint_file.close()
    weak_file.close()
    sources_file.close()
    print("Done.")


if __name__ == "__main__":
    main()
