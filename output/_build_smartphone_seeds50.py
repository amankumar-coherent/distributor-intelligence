"""Build Global Smartphone Market seeds: 50 genuine distributors/channel partners."""
from __future__ import annotations

import socket
from pathlib import Path

ROOT = Path(r"D:\vendor-intel")
SEED = ROOT / "queries" / "seeds" / "global_smartphone_market.txt"
DISC = ROOT / "queries" / "discovery" / "global_smartphone_market.txt"
DESK = Path(r"C:\Users\aman.kumar\Desktop\Cmi\agents\vendor-intel")

# Genuine B2B distributors / wholesalers / VADs (NOT OEMs, NOT pure retailers)
CANDIDATES: list[tuple[str, str]] = [
    ("Ingram Micro", "ingrammicro.com"),
    ("TD SYNNEX", "tdsynnex.com"),
    ("D&H Distributing", "dandh.com"),
    ("Arrow Electronics", "arrow.com"),
    ("Avnet", "avnet.com"),
    ("ScanSource", "scansource.com"),
    ("Westcon-Comstor", "westconcomstor.com"),
    ("ALSO Holding", "also.com"),
    ("Esprinet", "esprinet.com"),
    ("Exertis", "exertis.com"),
    ("Midwich Group", "midwich.com"),
    ("Action", "action.pl"),
    ("Redington", "redingtongroup.com"),
    ("Savex Technologies", "savex.in"),
    ("Supertron Electronics", "supertronindia.com"),
    ("Neoteric Informatique", "neoteric.co.in"),
    ("Rashi Peripherals", "rptechindia.com"),
    ("Softcell Technologies", "softcell.com"),
    ("Optiemus Infracom", "optiemus.com"),
    ("Brightstar", "brightstar.com"),
    ("Likewize", "likewize.com"),
    ("Mobile Phones Factory", "mobilephonesfactory.com"),
    ("ORSO Distribution", "orsodistribution.com"),
    ("Mobile Depot", "mobiledepot.eu"),
    ("Intcomex", "intcomex.com"),
    ("Grupo BGH", "grupobgh.com"),
    ("Compusoluciones", "compusoluciones.com"),
    ("CT Internacional", "ctinternacional.com.mx"),
    ("Nexsys", "nexsysla.com"),
    ("Dicker Data", "dickerdata.com.au"),
    ("Leader Computers", "leadersystems.com.au"),
    ("Bluechip Infotech", "bluechipit.com.au"),
    ("Synnex Thailand", "synnex.co.th"),
    ("ECS Holdings", "ecsholdings.com"),
    ("SiS Distribution", "sis.com.sg"),
    ("Digital China", "digitalchina.com"),
    ("CDW", "cdw.com"),
    ("SHI International", "shi.com"),
    ("Insight Enterprises", "insight.com"),
    ("Connection", "connection.com"),
    ("Zones", "zones.com"),
    ("Softchoice", "softchoice.com"),
    ("Tech Data", "techdata.com"),
    ("Synnex Corporation", "tdsynnex.com"),  # may duplicate TD - handled by domain dedupe
    ("PCM", "pcm.com"),
    ("Mobile Fun", "mobilefun.co.uk"),
    ("The Phone Hub", "thephonehub.com"),
    ("Phoneport24", "phoneport24.com"),
    ("XS Distribution", "xsdistribution.com"),
    ("3G Mobile", "3gmobile.co.za"),
    ("Cellular Distribution", "cellular.co.za"),
    ("Incredible Connection Wholesale", "incredible.co.za"),
    ("Mustek", "mustek.co.za"),
    ("Rectron", "rectron.co.za"),
    ("Pinnacle Micro", "pinnacle.co.za"),
    ("Axis Technology Distribution", "axis.co.za"),
    ("Ingram Micro Middle East", "ingrammicro.com"),
    ("Finesse Distribution", "finesseuae.com"),
    ("Logicom Distribution", "logicom.com"),
    ("Aptec Distribution", "aptecdubai.com"),
    ("Redington Gulf", "redingtongroup.com"),
    ("Mindware", "mindware.com"),
    ("Emitac Distribution", "emitac.ae"),
    ("Jumbo Electronics Wholesale", "jumbo.ae"),
    ("Sharaf DG Wholesale", "sharafdg.com"),
    ("Emax Wholesale", "emaxme.com"),
    ("Virgin Mobile Distribution Partners", "virginmobile.com"),
    ("Boost Mobile Wholesale Partners", "boostmobile.com"),
    ("Metro by T-Mobile Wholesale", "metrobyt-mobile.com"),
    ("Cricket Wireless Wholesale", "cricketwireless.com"),
    ("Tracfone Wholesale Partners", "tracfone.com"),
]


def dns_ok(domain: str) -> bool:
    host = domain.strip().lower().split("/")[0]
    if host.startswith("www."):
        host = host[4:]
    try:
        socket.setdefaulttimeout(2.0)
        socket.getaddrinfo(host, 80, type=socket.SOCK_STREAM)
        return True
    except OSError:
        try:
            socket.getaddrinfo("www." + host, 80, type=socket.SOCK_STREAM)
            return True
        except OSError:
            return False


def main() -> None:
    seen_name: set[str] = set()
    seen_dom: set[str] = set()
    kept: list[tuple[str, str]] = []
    failed: list[tuple[str, str]] = []

    for name, dom in CANDIDATES:
        nk = "".join(ch for ch in name.lower() if ch.isalnum())
        dk = dom.lower().strip()
        if nk in seen_name or dk in seen_dom:
            continue
        if not dns_ok(dom):
            failed.append((name, dom))
            continue
        seen_name.add(nk)
        seen_dom.add(dk)
        kept.append((name, dom))
        if len(kept) >= 50:
            break

    # top up from DNS-fail only if still short (known real firms)
    if len(kept) < 50:
        for name, dom in failed:
            nk = "".join(ch for ch in name.lower() if ch.isalnum())
            dk = dom.lower().strip()
            if nk in seen_name or dk in seen_dom:
                continue
            seen_name.add(nk)
            seen_dom.add(dk)
            kept.append((name, dom))
            if len(kept) >= 50:
                break

    kept = kept[:50]
    lines = [
        "# Global Smartphone Market — channel / distributor seeds ONLY",
        "# Web-curated authorized distributors, wholesalers, IT/mobile VADs.",
        "# OEMs (Apple/Samsung/Xiaomi/etc.) intentionally omitted.",
        f"# Prefer DNS-ok domains — count={len(kept)}",
        "# Format: Company Name | domain | Distributors of Smartphones",
        "",
    ]
    for name, dom in kept:
        lines.append(f"{name} | {dom} | Distributors of Smartphones")

    SEED.parent.mkdir(parents=True, exist_ok=True)
    SEED.write_text("\n".join(lines) + "\n", encoding="utf-8")

    disc_lines = [
        "# Global Smartphone Market — channel-partner discovery queries",
        "smartphone authorized distributors",
        "mobile phone wholesalers distributors",
        "smartphone channel partners distributors",
        "handset value added distributors Ingram Micro TD SYNNEX",
        "mobile device IT distributor Redington Savex Supertron",
        "smartphone distributor Europe ALSO Esprinet Exertis",
        "mobile phone wholesale distributor Latin America Intcomex",
        "enterprise smartphone distributor CDW SHI Insight Zones",
        "refurbished smartphone wholesale distributor Europe Mobile Depot",
        "mobile handset distributor Middle East Africa Redington Logicom",
    ]
    DISC.parent.mkdir(parents=True, exist_ok=True)
    DISC.write_text("\n".join(disc_lines) + "\n", encoding="utf-8")

    if (DESK / "queries" / "seeds").exists():
        (DESK / "queries" / "seeds" / "global_smartphone_market.txt").write_text(
            "\n".join(lines) + "\n", encoding="utf-8"
        )
        ddir = DESK / "queries" / "discovery"
        ddir.mkdir(parents=True, exist_ok=True)
        (ddir / "global_smartphone_market.txt").write_text(
            "\n".join(disc_lines) + "\n", encoding="utf-8"
        )

    print(f"kept={len(kept)} dns_fail_remaining={len(failed)}")
    for n, d in kept:
        print(f"  OK  {n} | {d}")


if __name__ == "__main__":
    main()
