"""Probe DNS + HTTP for every website in FINAL; list suffix-named companies."""
from __future__ import annotations

import json
import socket
import ssl
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from urllib.parse import urlparse
from urllib.request import Request, urlopen

from openpyxl import load_workbook

FP = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\global_flexible_packaging_market_global_FINAL.xlsx"
)
OUT = FP.parent / "website_dns_probe.json"


def domain_of(url: str) -> str:
    if not url:
        return ""
    u = url if "://" in url else "https://" + url
    try:
        return urlparse(u).netloc.lower().replace("www.", "")
    except Exception:
        return ""


def probe(name: str, website: str) -> dict:
    d = domain_of(website)
    result = {
        "company": name,
        "website": website,
        "domain": d,
        "dns_ok": False,
        "http_ok": False,
        "final_url": "",
        "status": None,
        "error": "",
    }
    if not d:
        result["error"] = "no_domain"
        return result
    try:
        socket.getaddrinfo(d, 443)
        result["dns_ok"] = True
    except Exception as e:
        result["error"] = f"dns:{e}"
        # also try without subdomain quirks
        try:
            socket.getaddrinfo(d, 80)
            result["dns_ok"] = True
            result["error"] = ""
        except Exception as e2:
            result["error"] = f"dns:{e2}"
            return result

    base = website if website.startswith("http") else "https://" + website
    for url in [base, base.rstrip("/") + "/"]:
        try:
            req = Request(
                url,
                headers={
                    "User-Agent": (
                        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
                    )
                },
                method="GET",
            )
            ctx = ssl.create_default_context()
            with urlopen(req, timeout=12, context=ctx) as resp:
                result["http_ok"] = True
                result["status"] = getattr(resp, "status", None) or resp.getcode()
                result["final_url"] = resp.geturl()
                return result
        except Exception as e:
            result["error"] = f"http:{type(e).__name__}:{e}"[:240]
            continue
    return result


def main():
    wb = load_workbook(FP, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    items = [(str(r[0] or "").strip(), str(r[1] or "").strip()) for r in rows]

    print("suffix / Inc style names:")
    for name, web in items:
        low = name.lower()
        if any(
            low.endswith(s)
            for s in (
                " inc",
                " inc.",
                " llc",
                " ltd",
                " ltd.",
                " corp",
                " corp.",
                " corporation",
                " company",
                " co.",
                " limited",
            )
        ):
            print(f"  {name} | {web}")

    results = []
    with ThreadPoolExecutor(max_workers=20) as ex:
        futs = {ex.submit(probe, n, w): (n, w) for n, w in items}
        done = 0
        for fut in as_completed(futs):
            results.append(fut.result())
            done += 1
            if done % 40 == 0:
                print(f"probed {done}/{len(items)}", flush=True)

    bad_dns = [r for r in results if not r["dns_ok"]]
    bad_http = [r for r in results if r["dns_ok"] and not r["http_ok"]]
    good = [r for r in results if r["http_ok"]]
    print("total", len(results), "good_http", len(good), "dns_fail", len(bad_dns), "http_fail", len(bad_http))
    print("---DNS FAIL---")
    for r in sorted(bad_dns, key=lambda x: x["company"]):
        print(f"  {r['company']} | {r['website']} | {r['error']}")
    print("---HTTP FAIL (DNS ok)---")
    for r in sorted(bad_http, key=lambda x: x["company"]):
        print(f"  {r['company']} | {r['website']} | {r['error']}")

    OUT.write_text(
        json.dumps(
            {
                "total": len(results),
                "good_http": len(good),
                "dns_fail": bad_dns,
                "http_fail": bad_http,
                "all": results,
            },
            indent=2,
        ),
        encoding="utf-8",
    )
    print("wrote", OUT)
    wb.close()


if __name__ == "__main__":
    main()
