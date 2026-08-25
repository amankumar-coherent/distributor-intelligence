"""High-confidence packaging DISTRIBUTOR candidates for pass 2 (agent-researched)."""
import json
from pathlib import Path
from urllib.parse import urlparse
from openpyxl import load_workbook

ROOT = Path(r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global")
FP = ROOT / "global_flexible_packaging_market_global_FINAL.xlsx"

wb = load_workbook(FP, read_only=True, data_only=True)
ws = wb.active
seen = set()
for r in ws.iter_rows(min_row=2, values_only=True):
    w = (r[1] or "").strip()
    if not w:
        continue
    try:
        seen.add(urlparse(w if "://" in w else "https://" + w).netloc.lower().replace("www.", ""))
    except Exception:
        pass
wb.close()

raw = [
    # US - verified distributors / wholesalers
    ("BOX Partners", "https://www.boxpartners.com", "US"),
    ("BASCO", "https://www.bascousa.com", "US"),
    ("Stamar Packaging", "https://stamarpackaging.com", "US"),
    ("Excel Pac LLC", "https://excelpacllc.com", "US"),
    ("Fillmore Container", "https://www.fillmorecontainer.com", "US"),
    ("MrBoxOnline", "https://www.mrboxonline.com", "US"),
    ("Industrial & Shippers Supply", "https://issipackaging.com", "US"),
    ("Package Design & Supply", "https://packagedesignsupply.com", "US"),
    ("Pride Container", "https://www.pridecontainer.com", "US"),
    ("Hisco", "https://www.hisco.com", "US"),
    ("Specialty Bottle", "https://www.specialtybottle.com", "US"),
    ("Bottles.com", "https://www.bottles.com", "US"),
    ("The Bottle Crew", "https://www.thebottlecrew.com", "US"),
    ("Cactus Containers", "https://www.cactuscontainers.com", "US"),
    ("Box Brothers", "https://www.boxbrothers.com", "US"),
    ("Uboxes", "https://www.uboxes.com", "US"),
    ("Innovation Packing Group", "https://ipgusa.com", "US"),
    ("World Packaging Corp", "https://www.worldpackaging.com", "US"),
    ("Sun Packaging", "https://www.sunpackaging.com", "US"),
    ("Tanner Packaging", "https://www.tannerpackaging.com", "US"),
    ("Restaurant Depot", "https://www.restaurantdepot.com", "US"),
    ("Sysco", "https://www.sysco.com", "US"),
    ("US Foods", "https://www.usfoods.com", "US"),
    ("Gordon Food Service", "https://www.gfs.com", "US"),
    ("Performance Food Group", "https://www.pfgc.com", "US"),
    ("Chef's Warehouse", "https://www.chefswarehouse.com", "US"),
    ("KeHE Distributors", "https://www.kehe.com", "US"),
    ("Dot Foods", "https://www.dotfoods.com", "US"),
    ("UNFI", "https://www.unfi.com", "US"),
    ("Costco Business Center", "https://www.costco.com", "US"),
    ("Sam's Club", "https://www.samsclub.com", "US"),
    ("Medline Industries", "https://www.medline.com", "US"),
    ("Henry Schein", "https://www.henryschein.com", "US"),
    ("McKesson", "https://www.mckesson.com", "US"),
    ("Cardinal Health", "https://www.cardinalhealth.com", "US"),
    ("Owens & Minor", "https://www.owens-minor.com", "US"),
    ("Motion Industries", "https://www.motionindustries.com", "US"),
    ("Applied Industrial Technologies", "https://www.applied.com", "US"),
    ("WESCO", "https://www.wesco.com", "US"),
    ("Graybar", "https://www.graybar.com", "US"),
    ("Kaman Distribution", "https://www.kamandirect.com", "US"),
    ("Wurth USA", "https://www.wurthusa.com", "US"),
    ("Acme Paper & Supply", "https://www.acmepaper.com", "US"),
    ("Carolina Retail Packaging", "https://www.carolinaretail.com", "US"),
    ("Service Paper Company", "https://www.servicepaper.com", "US"),
    ("Central Paper Products", "https://www.centralpaper.com", "US"),
    ("Butler Paper", "https://www.butlerpaper.com", "US"),
    ("Xpedx Legacy Dist", "https://www.xpedx.com", "US"),
    ("Clayton Packaging Solutions", "https://www.claytonpackaging.com", "US"),
    ("Midland Packaging Dist Alt", "https://www.midlandpkg.com", "US"),
    ("Great Northern Corp Packaging Dist", "https://www.greatnortherncorp.com", "US"),
    ("Durable Packaging Corp", "https://www.durablepackaging.com", "US"),
    ("National Packaging Company", "https://www.natpkg.com", "US"),
    ("American Packaging Distributors", "https://www.ampackaging.com", "US"),
    ("Packaging Unlimited Louisville", "https://www.packagingunlimited.com", "US"),
    ("All Packaging Company Denver", "https://www.allpackco.com", "US"),
    ("Lone Star Containers", "https://www.lonestarcontainers.com", "US"),
    ("Southwest Poly Bag", "https://www.southwestpolybag.com", "US"),
    ("Poly Pak America Dist", "https://www.polypakamerica.com", "US"),
    ("Western Plastics Dist", "https://www.westernplastics.com", "US"),
    ("Flexo Transparent Dist", "https://www.flexotransparent.com", "US"),
    ("Berry Global Distributor Partner Skip", "https://www.berryglobal.com", "US"),
    # CA / MX
    ("Shippers Supply", "https://www.shipperssupply.com", "CA"),
    ("Unisource Canada", "https://www.unisource.ca", "CA"),
    ("Ideal Packaging Canada", "https://www.idealpackaging.ca", "CA"),
    ("All Size Packaging Canada", "https://www.allsizepackaging.com", "CA"),
    ("Landmark Paper & Packaging", "https://www.landmarkpaper.com", "CA"),
    ("Atlantic Packaging Products", "https://www.atlantic.ca", "CA"),
    ("EmpaSoluciones", "https://empasoluciones.com.mx", "MX"),
    ("Empaques del Norte", "https://www.empaquesdelnorte.com", "MX"),
    ("Berlin Packaging Canada", "https://www.berlinpackaging.ca", "CA"),
    ("TricorBraun Canada Valid", "https://www.tricorbraun.ca", "CA"),
    # UK / IE / EU merchants
    ("Bag N Box Man Heritage Packaging Panda Alt Skip", "https://www.bagnboxman.co.uk", "GB"),
    ("Neptune Packaging UK", "https://www.neptunepackaging.co.uk", "GB"),
    ("Stockpak", "https://www.stockpak.co.uk", "GB"),
    ("IMG Packing", "https://www.imgpacking.co.uk", "GB"),
    ("Abco Kovex", "https://www.abcokovex.co.uk", "GB"),
    ("Lesters Packaging", "https://www.lesterspackaging.co.uk", "GB"),
    ("Direct Packaging UK", "https://www.directpackaging.co.uk", "GB"),
    ("BM Packaging UK", "https://www.bmpackaging.co.uk", "GB"),
    ("Agenda Boxes", "https://www.agendaboxes.co.uk", "GB"),
    ("Jaymar Packaging", "https://www.jaymarpackaging.co.uk", "GB"),
    ("UK Packaging Supplies", "https://www.ukpackagingsupplies.co.uk", "GB"),
    ("Aropack UK", "https://www.aropack.co.uk", "GB"),
    ("BigDug", "https://www.bigdug.co.uk", "GB"),
    ("Bunzl Catering Supplies", "https://www.bunzlcatering.co.uk", "GB"),
    ("Bidfood", "https://www.bidfood.co.uk", "GB"),
    ("Booker Wholesale", "https://www.booker.co.uk", "GB"),
    ("Bestway Wholesale", "https://www.bestwaywholesale.co.uk", "GB"),
    ("Screwfix", "https://www.screwfix.com", "GB"),
    ("Toolstation", "https://www.toolstation.com", "GB"),
    ("Cromwell Tools", "https://www.cromwell.co.uk", "GB"),
    ("McLernon Packaging", "https://mclernonpackaging.ie", "IE"),
    ("The Packaging Centre Ireland", "https://www.packagingcentre.ie", "IE"),
    ("RAJA Belgium", "https://www.rajapack.be", "BE"),
    ("RAJA Switzerland", "https://www.rajapack.ch", "CH"),
    ("RAJA Hungary", "https://www.rajapack.hu", "HU"),
    ("RAJA Romania", "https://www.rajapack.ro", "RO"),
    ("RAJA Slovakia", "https://www.rajapack.sk", "SK"),
    ("RAJA Sweden", "https://www.rajapack.se", "SE"),
    ("RAJA Norway", "https://www.rajapack.no", "NO"),
    ("Manutan Switzerland", "https://www.manutan.ch", "CH"),
    ("Manutan Portugal", "https://www.manutan.pt", "PT"),
    ("Manutan Hungary", "https://www.manutan.hu", "HU"),
    ("Manutan Austria", "https://www.manutan.at", "AT"),
    ("Manutan Czech Republic", "https://www.manutan.cz", "CZ"),
    ("Manutan Sweden", "https://www.manutan.se", "SE"),
    ("Viking Direct France", "https://www.vikingdirect.fr", "FR"),
    ("Viking Direct Germany", "https://www.vikingdirect.de", "DE"),
    ("Viking Direct Netherlands", "https://www.vikingdirect.nl", "NL"),
    ("Viking Direct Belgium", "https://www.vikingdirect.be", "BE"),
    ("Viking Direct Austria", "https://www.vikingdirect.at", "AT"),
    ("Viking Direct Spain", "https://www.vikingdirect.es", "ES"),
    ("Viking Direct Italy", "https://www.vikingdirect.it", "IT"),
    ("Lyreco", "https://www.lyreco.com", "FR"),
    ("Verpackungsdiscount", "https://www.verpackungsdiscount.de", "DE"),
    ("Kartonversand", "https://www.kartonversand.de", "DE"),
    ("Packaging24", "https://www.packaging24.com", "DE"),
    ("Paketversandshop", "https://www.paketversandshop.de", "DE"),
    ("Transpak Austria", "https://www.transpak.at", "AT"),
    ("Ratioform Austria", "https://www.ratioform.at", "AT"),
    ("Ratioform Switzerland", "https://www.ratioform.ch", "CH"),
    ("Ratioform Netherlands", "https://www.ratioform.nl", "NL"),
    ("Ratioform France", "https://www.ratioform.fr", "FR"),
    ("Papier Union", "https://www.papierunion.de", "DE"),
    ("Igepa Group", "https://www.igepa.de", "DE"),
    ("Papyrus", "https://www.papyrus.com", "SE"),
    ("Inapa", "https://www.inapa.com", "PT"),
    ("Ahlsell Norway", "https://www.ahlsell.no", "NO"),
    ("Ahlsell Finland", "https://www.ahlsell.fi", "FI"),
    ("Ahlsell Denmark", "https://www.ahlsell.dk", "DK"),
    ("Gerdmans", "https://www.gerdmans.com", "SE"),
    ("Eriks", "https://www.eriks.com", "NL"),
    ("Brammer", "https://www.brammer.biz", "GB"),
    ("Bossard", "https://www.bossard.com", "CH"),
    ("Rexel", "https://www.rexel.com", "FR"),
    ("Sonepar", "https://www.sonepar.com", "FR"),
    ("Conrad Electronic", "https://www.conrad.com", "DE"),
    ("Distrelec", "https://www.distrelec.com", "CH"),
    ("TVH", "https://www.tvh.com", "BE"),
    ("Kramp", "https://www.kramp.com", "NL"),
    ("BayWa", "https://www.baywa.de", "DE"),
    # ANZ
    ("Adelaide Packaging Supplies", "https://www.adpack.com.au", "AU"),
    ("Boxworks Australia", "https://www.boxworks.com.au", "AU"),
    ("Bags Plus Australia", "https://www.bagsplus.com.au", "AU"),
    ("Pluspak Australia", "https://www.pluspak.com.au", "AU"),
    ("MaxPak Australia", "https://www.maxpak.com.au", "AU"),
    ("Ozpack", "https://www.ozpack.com.au", "AU"),
    ("Spicers Australia", "https://www.spicers.com.au", "AU"),
    ("Ball & Doggett", "https://www.ballanddoggett.com.au", "AU"),
    ("Winc Australia", "https://www.winc.com.au", "AU"),
    ("Premier Packaging NZ", "https://www.premierpackaging.co.nz", "NZ"),
    ("Kiwi Plastic Packaging", "https://kiwiplasticpackaging.co.nz", "NZ"),
    # Asia / Africa / other
    ("Clearpack", "https://www.clearpack.com", "SG"),
    ("Seow Khim Polythelene", "https://www.skp.com.sg", "SG"),
    ("Tantraco Enterprise", "https://www.tantraco.com.sg", "SG"),
    ("Bizongo", "https://www.bizongo.com", "IN"),
    ("Packaging 360", "https://www.packaging360.in", "IN"),
    ("Bidvest Paperplus", "https://www.bidvestpaperplus.co.za", "ZA"),
    ("Staples.com", "https://www.staples.com", "US"),
    ("Office Depot", "https://www.officedepot.com", "US"),
    ("Home Depot", "https://www.homedepot.com", "US"),
    ("Grainger Canada", "https://www.grainger.ca", "CA"),
    ("Fastenal Canada", "https://www.fastenal.ca", "CA"),
    ("Staples Canada", "https://www.staples.ca", "CA"),
    ("Bunzl Canada", "https://www.bunzl.ca", "CA"),
    ("RAJA Ireland", "https://www.rajapack.ie", "IE"),
    ("Manutan Ireland", "https://www.manutan.ie", "IE"),
    ("Uline Europe", "https://www.uline.eu", "NL"),
    ("Uline UK", "https://www.uline.co.uk", "GB"),
    ("Packhelp Germany", "https://www.packhelp.de", "DE"),
    ("Packhelp France", "https://www.packhelp.fr", "FR"),
    ("Packhelp UK", "https://www.packhelp.co.uk", "GB"),
    ("Emballages Service", "https://www.emballages-service.fr", "FR"),
    ("BVS Verpacken", "https://bvs-verpacken.com", "CH"),
    ("Antalis Germany", "https://www.antalis.de", "DE"),
    ("JPG Belgium", "https://www.jpg.be", "BE"),
    ("RAJA Luxembourg", "https://www.rajapack.lu", "LU"),
    ("Manutan Luxembourg", "https://www.manutan.lu", "LU"),
    ("Manutan Romania", "https://www.manutan.ro", "RO"),
    ("Manutan Slovakia", "https://www.manutan.sk", "SK"),
    ("Manutan Denmark", "https://www.manutan.dk", "DK"),
    ("Manutan Norway", "https://www.manutan.no", "NO"),
    ("Manutan Finland", "https://www.manutan.fi", "FI"),
    ("Ratioform Belgium", "https://www.ratioform.be", "BE"),
    ("Ratioform Spain", "https://www.ratioform.es", "ES"),
    ("Ratioform UK", "https://www.ratioform.co.uk", "GB"),
    ("Ratioform Poland", "https://www.ratioform.pl", "PL"),
    ("Transpak Switzerland", "https://www.transpak.ch", "CH"),
    ("Transpak Netherlands", "https://www.transpak.nl", "NL"),
    ("Complete Packaging Systems Dist", "https://www.completepackaging.com", "US"),
]

cands = []
for name, url, cc in raw:
    if "skip" in name.lower():
        continue
    try:
        d = urlparse(url).netloc.lower().replace("www.", "")
    except Exception:
        continue
    if not d or d in seen:
        continue
    seen.add(d)
    cands.append({"id": f"p2-{len(cands)+1:03d}", "company": name, "website": url, "country": cc})

out = ROOT / "add_distributor_candidates.json"
out.write_text(json.dumps(cands, indent=2), encoding="utf-8")
print("pass2 candidates", len(cands))
for c in cands[:15]:
    print(" ", c["company"], c["website"])
