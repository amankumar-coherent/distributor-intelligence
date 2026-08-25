"""Analyze website_verify_audit.json and recheck soft/fail cases."""
from __future__ import annotations

import json
import socket
from collections import Counter
from pathlib import Path
from urllib.parse import urlparse

AUDIT = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\website_verify_audit.json"
)


def domain_of(url: str) -> str:
    u = url if "://" in url else "https://" + url
    return urlparse(u).netloc.lower().replace("www.", "")


def dns_ok(host: str) -> bool:
    try:
        socket.getaddrinfo(host, 443)
        return True
    except Exception:
        try:
            socket.getaddrinfo(host, 80)
            return True
        except Exception:
            return False


def main() -> None:
    a = json.loads(AUDIT.read_text(encoding="utf-8"))
    print("summary", a["total"], "ok", a["ok"], "soft", a["soft_name_mismatch_or_unclear"], "fail", a["fail"])
    print("fail reasons", Counter(x.get("reason") for x in a["fails"]))

    dns_dead = []
    http_fail_dns_ok = []
    for x in a["fails"]:
        host = x.get("domain") or domain_of(x.get("website") or "")
        if not host or not dns_ok(host):
            dns_dead.append(x)
        else:
            http_fail_dns_ok.append(x)

    print("\nDNS dead:", len(dns_dead))
    for x in dns_dead:
        print(" -", x["company"], x.get("website"), x.get("reason"), str(x.get("error", ""))[:100])

    print("\nHTTP fail but DNS OK:", len(http_fail_dns_ok))
    for x in http_fail_dns_ok:
        err = str(x.get("error", ""))[:120]
        print(" -", x["company"], "|", x.get("website"), "|", err)

    print("\nSOFT:")
    for x in a["soft"]:
        print(" -", x["company"], "|", x.get("title", "")[:100], "|", x["website"])


if __name__ == "__main__":
    main()
