"""Japan Fashion E-commerce Market -- ONE market, Japan HQ only, ONLINE
FASHION SELLERS (user, 2026-10-29).

Inputs: user's definition, product / price / age / payment / end-user /
sales-channel / region segmentation, 5 exclusions, and "Sample_Japan_Fashion
E-commerce Market_Coherent_Market_Insights.pdf" (Downloads; scope pages 16-18,
company profiles page 19). The report profiles SELLERS, not distributors:
online fashion marketplaces (Rakuten, Mercari, Enigmo / BUYMA, Jade Group /
LOCONDO), multi-brand fashion retail and e-commerce groups (Fast Retailing,
and ST, World, Onward, TSI, Baroque, PAL, Honeys, Shimamura), footwear and
fashion specialty e-commerce (ABC-Mart, Tokyo Base) and fashion brand owners
/ omnichannel players (Cross Plus, Baycrew's, BEAMS, Urban Research, JUN).
So the list is companies that SELL apparel, footwear, accessories or eyewear
to consumers through online channels in Japan -- the "retail" kind
(discovery.DEF_RETAIL, verdict online_seller).

EXCLUDED: offline-only fashion retailers; non-fashion and general e-commerce
sellers with no fashion store or platform; standalone technology, logistics,
fulfilment, payment and marketing service providers; B2B-only wholesalers and
OEM makers; companies headquartered outside Japan.

Searched prefecture by prefecture, then Tokyo / Osaka business districts.
Scope rides in ONE Google URL (budget 7,550 / hard 7,800); measure with
scripts/preflight_prompt.py --batch jpf.
"""
from __future__ import annotations

MARKETS_JPF: tuple[tuple[str, str, str, str], ...] = (
    (
        "jpf_ecommerce",
        "Japan Fashion E-commerce Market",
        "PRODUCT SCOPE (required): FASHION sold ONLINE in Japan - apparel "
        "(men's, women's, kidswear), footwear, bags, wallets, belts, scarves, "
        "hats, watches, jewelry, sunglasses and eyewear - through online "
        "marketplaces, brand-owned websites, mobile apps or social commerce. "
        "The company SELLS this online (verdict online_seller): fashion "
        "marketplaces (Rakuten Fashion, ZOZOTOWN, Mercari, BUYMA, LOCONDO), "
        "fashion retailers and select shops with online stores (UNIQLO, "
        "BEAMS, Urban Research, ABC-Mart), fashion brands selling online. "
        "EXCLUDE offline-only shops, non-fashion or general e-commerce, "
        "technology, logistics, payment or marketing services, B2B-only "
        "wholesalers and OEM makers.",
        "retail",
    ),
)

SIBLINGS_JPF: dict[str, str] = {}

# The report's regions (Kanto, Kansai, Chubu, Kyushu and Okinawa, rest of
# Japan) as all 47 prefectures, then the districts where Tokyo, Osaka, Kobe
# and Nagoya fashion companies cluster.
JPF_REGIONS: tuple[str, ...] = (
    "Tokyo", "Osaka", "Kanagawa", "Aichi", "Fukuoka", "Hyogo", "Kyoto",
    "Saitama", "Chiba", "Hokkaido", "Okayama", "Hiroshima", "Shizuoka",
    "Gifu", "Ibaraki", "Tochigi", "Gunma", "Nagano", "Niigata", "Miyagi",
    "Fukushima", "Yamagata", "Iwate", "Akita", "Aomori", "Ishikawa",
    "Toyama", "Fukui", "Yamanashi", "Mie", "Shiga", "Nara", "Wakayama",
    "Tottori", "Shimane", "Yamaguchi", "Tokushima", "Kagawa", "Ehime",
    "Kochi", "Saga", "Nagasaki", "Kumamoto", "Oita", "Miyazaki",
    "Kagoshima", "Okinawa",
    "Shibuya Tokyo", "Minato Tokyo", "Chuo Tokyo", "Shinjuku Tokyo",
    "Meguro Tokyo", "Setagaya Tokyo", "Chiyoda Tokyo", "Shinagawa Tokyo",
    "Taito Tokyo", "Osaka Minami Senba", "Kobe city", "Nagoya city",
)

DEFINITIONS: dict[str, str] = {
    "jpf_ecommerce": (
        "The Japan Fashion E-Commerce Market covers the online sale of "
        "apparel, footwear, and fashion accessories through online "
        "marketplaces, company- or brand-owned websites, social "
        "commerce/media, and mobile applications in Japan. The market "
        "assessment includes demand by product type, price range, age group, "
        "payment method, end user, sales channel, and geographical region, "
        "alongside pricing, consumer behaviour, supply chain and logistics, "
        "digital shopping experience, and competitive developments. "
        "BY PRODUCT TYPE: Apparel (Men's Wear - Formal, Casual, Sportswear, "
        "Ethnic; Women's Wear - Dresses, Tops, Bottoms, Ethnic, Maternity; "
        "Kidswear - Infants, Toddlers, Teen Segment); Footwear (Sports Shoes, "
        "Casual Footwear, Formal Footwear, Slippers/Slides); Accessories "
        "(Bags and Wallets, Belts and Scarves, Hats and Headwear, Watches and "
        "Jewelry); Sunglasses and Eyewear. BY PRICE RANGE: Premium, "
        "Mid-Range, Economy. BY AGE GROUP: Gen Z (18-24), Millennials "
        "(25-40), Gen X (41-56). BY PAYMENT METHOD: Debit/Credit Cards, "
        "E-Wallets, Bank Transfer, Cash on Delivery, Others (Cryptocurrency, "
        "etc.). BY END USER: Men, Women, Children. BY SALES CHANNEL: Online "
        "Marketplaces, Company / Brand-Owned Websites, Social Commerce / "
        "Media, Mobile Apps. BY REGION: Kanto, Kansai, Chubu, Kyushu and "
        "Okinawa, Rest of Japan. "
        "KEY EXCLUSIONS: offline-only fashion retail sales not captured "
        "through the defined e-commerce channels; non-fashion merchandise "
        "outside the defined apparel, footwear, accessories and eyewear "
        "categories; general e-commerce sales unrelated to fashion products; "
        "standalone technology, logistics and payment services not included "
        "as fashion merchandise sales; other products and services outside "
        "the defined Japan Fashion E-Commerce Market scope. "
        "COMPANY TYPE IN THIS LIST - ONLINE FASHION SELLER: a company "
        "headquartered in Japan that sells apparel, footwear, fashion "
        "accessories or eyewear to consumers through online channels it "
        "operates - online fashion marketplaces, multi-brand fashion retail "
        "and e-commerce groups, footwear and fashion specialty e-commerce, "
        "and fashion brand owners / omnichannel players (the report's four "
        "company groups). EXCLUDED companies: offline-only retailers; "
        "general or non-fashion e-commerce sellers; technology, platform, "
        "logistics, fulfilment, payment and marketing service providers; "
        "B2B-only wholesalers and OEM manufacturers; companies headquartered "
        "outside Japan."
    ),
}

# Verifier KEEP/DROP (verify_distributors_ai_mode reads this; keep short).
VERIFY_RULES: dict[str, str] = {
    "jpf_ecommerce": (
        "KEEP only Japan-based companies that SELL apparel, footwear, bags, "
        "accessories or eyewear to consumers ONLINE (fashion marketplaces, "
        "online stores of fashion retailers, select shops or brands, "
        "fashion apps).\n"
        "DROP offline-only shops, non-fashion or general e-commerce, "
        "technology/logistics/payment/marketing services, B2B-only "
        "wholesalers, OEM makers."
    ),
}

# Product-evidence gate for the final build (market_batches.evidence_for):
# fashion product words in English and Japanese. Test on the verified list
# before finalize.
EVIDENCE: dict[str, str] = {
    "jpf_ecommerce": (
        r"fashion|apparel|cloth|wear\b|garment|dress|shirt|denim|jean|"
        r"shoe|footwear|sneaker|boot|sandal|bag|wallet|accessor|jewel|watch|"
        r"eyewear|glasses|sunglass|kimono|lingerie|underwear|hat\b|"
        r"select shop|ファッション|アパレル|衣料|靴|服"
    ),
}


def by_slug_jpf(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_JPF:
        if entry[0] == slug:
            return entry
    raise KeyError(f"unknown jpf market slug: {slug}")
