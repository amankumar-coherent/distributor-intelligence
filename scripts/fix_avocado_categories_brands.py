"""Verify/fix avocado Core Categories, Specialty Focus, Price Segment, Key Brands."""
from __future__ import annotations

import re
import shutil
import unicodedata
from collections import Counter
from datetime import datetime
from pathlib import Path

import openpyxl

PATH = Path(
    r"D:\vendor-intel\output\web_expand\global_avocado_oil_market_global\global_avocado_oil_market_global_FINAL.xlsx"
)

GENERIC_CORE = {
    "avocado oil; edible oils; specialty oils",
    "avocado oil",
}
GENERIC_FOCUS = {
    "distributors of avocado oil / edible oils",
    "avocado oil distribution / wholesale / retail channel",
    "avocado oil distribution",
}
GENERIC_BRANDS = {
    "private label; regional avocado oil brands",
    "private label",
    "n/a",
    "not publicly disclosed",
    "",
}

# company key → (Core Categories, Specialty Focus, Price Segment, Key Brands)
VERIFIED: dict[str, tuple[str, str, str, str]] = {
    "sysco": (
        "Broadline foodservice; edible oils; avocado oil SKUs",
        "National foodservice distribution of culinary oils to restaurants/hotels",
        "Full range (Value–Premium)",
        "Sysco Imperial; Sysco Classic; brand avocado oils via Sysco Specialty",
    ),
    "sysco canada": (
        "Broadline foodservice; edible oils; avocado oil",
        "Canadian foodservice oil & grocery distribution",
        "Full range (Value–Premium)",
        "Sysco brand oils; regional avocado oil brands",
    ),
    "sysco europe": (
        "Broadline foodservice; edible oils; avocado oil",
        "European Sysco/Brakes foodservice oil channel",
        "Full range (Value–Premium)",
        "Sysco / Brakes private label; brand edible oils",
    ),
    "sysco specialty": (
        "Specialty foodservice; premium oils; avocado oil",
        "Specialty/gourmet oils for chef & fine-dining accounts",
        "Specialty / Premium",
        "Specialty avocado & culinary oil brands via Sysco Specialty",
    ),
    "us foods": (
        "Broadline foodservice; edible oils; avocado oil",
        "US foodservice distribution of cooking & specialty oils",
        "Full range (Value–Premium)",
        "US Foods Chef's Line; Monarch; brand avocado oils",
    ),
    "us foods chef store": (
        "Cash & carry foodservice; edible oils; avocado oil",
        "Warehouse club oils for independent restaurants",
        "Mid-market / Value",
        "CHEF'STORE private label; national oil brands",
    ),
    "performance food group": (
        "Broadline foodservice; edible oils; avocado oil",
        "PFG foodservice oil & grocery distribution",
        "Full range (Value–Premium)",
        "PFG house brands; brand edible/avocado oils",
    ),
    "pfg customized distribution": (
        "Custom foodservice distribution; edible oils",
        "Chain-account oil & grocery fulfillment",
        "Mid-market / Contract",
        "Customer private label; brand oils",
    ),
    "gordon food service": (
        "Broadline foodservice; edible oils; avocado oil",
        "Family-operated GFS foodservice oil supply",
        "Full range (Value–Premium)",
        "GFS Trade East; brand avocado oils",
    ),
    "unfi united natural foods": (
        "Natural/organic grocery wholesale; specialty oils; avocado oil",
        "Natural channel distribution of organic & specialty avocado oils",
        "Specialty / Premium",
        "Field Day; equal Exchange channel brands; natural avocado oil labels",
    ),
    "kehe distributors": (
        "Natural/specialty grocery wholesale; avocado oil brands",
        "Specialty retail placement for premium edible oils",
        "Specialty / Premium",
        "Natural/organic avocado oil brands to independent & chain specialty",
    ),
    "kehe specialty": (
        "Specialty grocery; premium oils",
        "Specialty oil brands to retail",
        "Specialty / Premium",
        "Premium avocado & specialty oil brands",
    ),
    "brenntag": (
        "Specialty chemicals; food ingredients; vegetable oils",
        "Food-grade & industrial oil/ingredient distribution",
        "Industrial / Mid-to-Premium ingredients",
        "Principal supplier brands (oils & oleochemicals); Brenntag food ingredients",
    ),
    "brenntag north america": (
        "Food ingredients; specialty oils; oleochemicals",
        "NA distribution of food-grade oils including avocado derivatives",
        "Industrial / Mid-to-Premium ingredients",
        "Supplier-principal edible oil & ingredient brands",
    ),
    "barentz": (
        "Specialty food & nutrition ingredients; oils",
        "Life-science ingredients incl. specialty lipid systems",
        "Specialty / Premium ingredients",
        "Barentz portfolio principals; specialty oil suppliers",
    ),
    "barentz north america": (
        "Specialty food ingredients; oils",
        "NA specialty lipid & nutrition ingredients",
        "Specialty / Premium ingredients",
        "Principal specialty oil & nutrition brands",
    ),
    "azelis": (
        "Specialty chemicals & food ingredients",
        "Formulation ingredients; food lipids & specialties",
        "Specialty / Premium ingredients",
        "Azelis principal brands (food & personal care oils)",
    ),
    "azelis americas": (
        "Specialty ingredients; food oils",
        "Americas specialty ingredient distribution",
        "Specialty / Premium ingredients",
        "Principal specialty oil brands",
    ),
    "imcd": (
        "Specialty chemicals & food ingredients",
        "Formulation & nutrition ingredients including lipids",
        "Specialty / Premium ingredients",
        "IMCD principal supplier brands",
    ),
    "imcd us": (
        "Specialty food ingredients; oils",
        "US specialty ingredient distribution",
        "Specialty / Premium ingredients",
        "Principal specialty oil brands",
    ),
    "dksh": (
        "Market expansion services; food & ingredients",
        "Asia-focused distribution of F&B ingredients and oils",
        "Mid-market / Specialty",
        "Principal F&B and oil brands across Asia",
    ),
    "univar solutions": (
        "Chemical distribution; food ingredients",
        "Industrial & food-grade chemical/oil distribution",
        "Industrial / Mid-market",
        "Principal chemical & ingredient suppliers",
    ),
    "fyffes": (
        "Fresh produce; bananas; tropical fruit; avocado",
        "Fresh avocado & tropical produce (oil as related category)",
        "Commodity / Mid-market fresh",
        "Fyffes; Sol; related produce brands",
    ),
    "eosta": (
        "Organic fresh produce; avocados",
        "Organic avocado & produce distribution (Nature & More)",
        "Specialty / Organic Premium",
        "Nature & More; organic grower brands",
    ),
    "anecoop": (
        "Fresh produce cooperative; citrus; avocado",
        "Spanish cooperative avocado & produce export",
        "Commodity / Mid-market fresh",
        "Anecoop grower brands",
    ),
    "trader joe s": (
        "Specialty grocery retail; private-label oils",
        "Private-label specialty avocado & culinary oils",
        "Specialty / Mid-to-Premium",
        "Trader Joe's",
    ),
    "kirkland signature": (
        "Warehouse-club private label; edible oils",
        "Costco Kirkland avocado / culinary oils",
        "Value / Mid-market club",
        "Kirkland Signature",
    ),
    "crisco": (
        "Branded edible oils & shortenings",
        "Mainstream branded cooking oils (Smucker)",
        "Mainstream / Mid-market",
        "Crisco",
    ),
    "acme hardesty": (
        "Oleochemicals; castor & vegetable oil derivatives",
        "Bio-based oleochemical distribution (not retail avocado oil)",
        "Industrial / Specialty",
        "Acme-Hardesty specialty oleochemical lines",
    ),
    "batory foods": (
        "Food ingredients; specialty oils; avocado oil",
        "Ingredient distributor for food manufacturers",
        "Mid-market / Specialty ingredients",
        "Principal edible oil & ingredient brands",
    ),
    "jedwards international": (
        "Bulk natural & carrier oils; avocado oil",
        "B2B bulk refined/unrefined avocado & botanical oils",
        "Specialty / Bulk Mid-to-Premium",
        "Jedwards bulk oils (private label capable)",
    ),
    "new directions aromatics": (
        "Bulk natural oils; carrier oils; avocado oil",
        "Wholesale natural & cosmetic-grade avocado oil",
        "Specialty / Mid-to-Premium",
        "New Directions Aromatics",
    ),
    "hbno bulk": (
        "Bulk essential & carrier oils; avocado oil",
        "Wholesale bulk avocado & botanical oils",
        "Specialty / Bulk Mid-market",
        "HBNO Bulk",
    ),
    "essential wholesale": (
        "Cosmetic ingredients; carrier oils; avocado oil",
        "Wholesale cosmetic & natural oil bases",
        "Specialty / Mid-to-Premium",
        "Essential Wholesale & Labs",
    ),
    "gustav heess": (
        "Vegetable oils; specialty edible & cosmetic oils",
        "European specialty vegetable oil trading incl. avocado",
        "Specialty / Premium",
        "Gustav Heess oil portfolio",
    ),
    "henry lamotte oils": (
        "Vegetable oils; specialty oils",
        "German specialty oil merchant (food & technical)",
        "Specialty / Mid-to-Premium",
        "Henry Lamotte Oils",
    ),
    "la tourangelle distribution": (
        "Artisan specialty oils; avocado oil",
        "Premium roasted & specialty culinary oils",
        "Premium / Specialty",
        "La Tourangelle",
    ),
    "oils of aloha": (
        "Hawaiian specialty oils; macadamia; avocado",
        "Island specialty culinary oils",
        "Premium / Specialty",
        "Oils of Aloha",
    ),
    "avohass": (
        "Avocado oil brand; D2C specialty oil",
        "Premium Hass avocado oil brand",
        "Premium",
        "Avohass",
    ),
    "kevala s avocado oil bulk": (
        "Bulk organic avocado oil",
        "B2B bulk organic avocado oil supply",
        "Specialty / Organic Premium",
        "Kevala",
    ),
    "bulk avocado oil": (
        "Bulk avocado oil wholesale",
        "B2B bulk avocado oil sourcing & wholesale",
        "Commodity / Mid-market bulk",
        "Private label / sourced bulk avocado oil",
    ),
    "california avocado commission": (
        "California avocado marketing; oil promotion",
        "Commodity board promoting CA avocado & oil uses",
        "N/A (industry association)",
        "California Avocado (generic category)",
    ),
    "wilmar international distribution": (
        "Edible oils & fats; tropical oils; food ingredients",
        "Global agribusiness oils (avocado as specialty niche)",
        "Commodity / Mid-market to Premium",
        "Wilmar; Arawana; other Wilmar edible oil brands",
    ),
    "olam food ingredients distribution": (
        "Food ingredients; edible oils; specialty lipids",
        "ofi ingredient oils & specialty fats channel",
        "Mid-market / Specialty ingredients",
        "ofi / Olam ingredient brands",
    ),
    "metro ag": (
        "Cash & carry wholesale; edible oils; avocado oil",
        "HORECA cash & carry culinary oils",
        "Value / Mid-market",
        "Metro Chef; Aro; brand edible oils",
    ),
    "makro netherlands": (
        "Cash & carry; edible oils",
        "Makro HORECA oils assortment",
        "Value / Mid-market",
        "Metro/Makro private label; brand oils",
    ),
    "jetro restaurant depot": (
        "Cash & carry foodservice; edible oils",
        "Membership warehouse oils for foodservice",
        "Value / Mid-market",
        "Restaurant Depot / Jetro private label; brand oils",
    ),
    "restaurant depot": (
        "Cash & carry foodservice; edible oils",
        "Membership warehouse culinary oils",
        "Value / Mid-market",
        "Restaurant Depot private label; brand oils",
    ),
    "baldor specialty foods": (
        "Specialty produce & specialty foods; oils",
        "Chef-driven specialty produce & specialty oils",
        "Specialty / Premium",
        "Specialty avocado oil & produce brands for chefs",
    ),
    "chefs warehouse": (
        "Specialty foodservice; premium oils",
        "Gourmet oils for independent fine dining",
        "Premium / Specialty",
        "Specialty avocado & finishing oil brands",
    ),
    "sid wainer and son": (
        "Specialty produce; specialty foods",
        "Specialty produce & gourmet grocery incl. oils",
        "Specialty / Premium",
        "Specialty oil & produce brands",
    ),
    "borges international group": (
        "Edible oils; olive oil; nuts; specialty oils",
        "Branded Mediterranean oils (avocado as adjacency)",
        "Mid-to-Premium branded",
        "Borges; Capricho Andaluz; related Borges oils",
    ),
    "deoleo distribution partners": (
        "Olive oil & specialty edible oils",
        "Global olive-oil major; specialty oil adjacencies",
        "Mid-to-Premium branded",
        "Bertolli; Carapelli; Carbonell",
    ),
    "savola distribution": (
        "Edible oils; ghee; food staples",
        "MENA edible oils distribution",
        "Mainstream / Value–Mid",
        "Afia; Al-Arabiya; other Savola oil brands",
    ),
    "costco mexico wholesale": (
        "Warehouse club; edible oils; avocado oil",
        "Club-pack culinary & avocado oils",
        "Value / Mid-market club",
        "Kirkland Signature; brand oils",
    ),
    "wakefern food corp": (
        "Grocery wholesale cooperative; edible oils",
        "Retailer-owned grocery oils assortment",
        "Mainstream / Mid-market",
        "ShopRite; Wholesome Pantry; brand oils",
    ),
    "c and s wholesale grocers": (
        "Grocery wholesale; edible oils",
        "Wholesale grocery oils to retail banners",
        "Mainstream / Value–Mid",
        "Retailer private label; national oil brands",
    ),
    "bidfood uk": (
        "Broadline foodservice; edible oils; avocado oil",
        "UK Bidfood foodservice culinary oils",
        "Full range (Value–Premium)",
        "Bidfood private label; brand avocado oils",
    ),
    "colabor group inc": (
        "Broadline foodservice & wholesale; edible oils; avocado oil",
        "Canadian foodservice/wholesale culinary oils for restaurants, hotels & institutions",
        "Full range (Value–Premium)",
        "Colabor / Summit private label; brand edible oils",
    ),
    "sam s club mexico": (
        "Warehouse club wholesale; edible oils; avocado oil",
        "Mexico club-pack culinary & avocado oils for members",
        "Value / Mid-market club",
        "Member's Mark; brand edible oils",
    ),
    "brake bros": (
        "Broadline foodservice; edible oils",
        "Brakes/Sysco UK foodservice oils",
        "Full range (Value–Premium)",
        "Brakes private label; brand oils",
    ),
    "brakes": (
        "Broadline foodservice; edible oils",
        "Brakes foodservice culinary oils",
        "Full range (Value–Premium)",
        "Brakes private label; brand oils",
    ),
    "sligro food group": (
        "Foodservice & cash & carry; edible oils",
        "Dutch foodservice/wholesale culinary oils",
        "Mid-market / Full range",
        "Sligro private label; brand oils",
    ),
    "tradin organic": (
        "Organic ingredients; organic oils",
        "Organic avocado & specialty oil trading",
        "Organic / Premium specialty",
        "Tradin Organic supplier network",
    ),
    "prinova": (
        "Nutrition & food ingredients; specialty oils",
        "Ingredient oils for food/nutrition manufacturers",
        "Specialty ingredients Mid-to-Premium",
        "Prinova / NAGASE ingredient principals",
    ),
    "parchem": (
        "Specialty chemicals; fine chemicals",
        "Chemical intermediates (oils as specialty chem)",
        "Industrial / Specialty",
        "Parchem catalog chemicals",
    ),
    "nature s pride": (
        "Fresh avocados; ripe & ready produce",
        "Fresh avocado specialist (oil as related)",
        "Mid-market / Fresh specialty",
        "Nature's Pride avocados",
    ),
    "markon cooperative": (
        "Produce procurement cooperative",
        "Foodservice produce buying (avocado fresh focus)",
        "Foodservice Mid-market",
        "Markon First Crop; ReadySet Serve",
    ),
    "pro act": (
        "Produce procurement cooperative",
        "Foodservice fresh produce procurement",
        "Foodservice Mid-market",
        "Pro*Act member brands / fresh programs",
    ),
}

# Business-type templates: (core, focus, price, brands)
BT_TEMPLATES: dict[str, tuple[str, str, str, str]] = {
    "Broadline Foodservice Distributor": (
        "Broadline foodservice; edible oils; avocado oil",
        "Foodservice distribution of culinary & avocado oils to HoReCa",
        "Full range (Value–Premium)",
        "House/private-label oils; national avocado oil brands",
    ),
    "Foodservice Distributor": (
        "Foodservice; edible oils; avocado oil",
        "Regional foodservice culinary oil supply",
        "Mid-market / Full range",
        "Private label & brand avocado/culinary oils",
    ),
    "Foodservice Wholesale": (
        "Foodservice wholesale; edible oils",
        "Wholesale culinary oils for foodservice operators",
        "Mid-market / Value",
        "Private label & brand edible oils",
    ),
    "Cash & Carry / Wholesale": (
        "Cash & carry wholesale; edible oils; avocado oil",
        "Self-service wholesale oils for retail & HORECA",
        "Value / Mid-market",
        "Cash & carry private label; brand edible oils",
    ),
    "Cash & Carry / Foodservice Wholesale": (
        "Foodservice cash & carry; edible oils",
        "Membership warehouse culinary oils",
        "Value / Mid-market",
        "Warehouse private label; brand oils",
    ),
    "Grocery Wholesale": (
        "Grocery wholesale; edible oils; avocado oil",
        "Wholesale grocery oils to retail stores",
        "Mainstream / Mid-market",
        "Retailer private label; national oil brands",
    ),
    "Grocery Wholesale Cooperative": (
        "Cooperative grocery wholesale; edible oils",
        "Retailer-owned wholesale oils assortment",
        "Mainstream / Mid-market",
        "Member banner private label; brand oils",
    ),
    "Grocery Wholesale / Retail": (
        "Grocery wholesale & retail; edible oils",
        "Integrated wholesale-retail oils assortment",
        "Mainstream / Mid-market",
        "Banner private label; brand oils",
    ),
    "Natural / Specialty Grocery Wholesale": (
        "Natural/specialty wholesale; organic oils; avocado oil",
        "Natural channel specialty & organic avocado oils",
        "Specialty / Premium",
        "Natural/organic avocado oil brands",
    ),
    "Specialty Food Distributor": (
        "Specialty foods; specialty oils; avocado oil",
        "Specialty & gourmet edible oil distribution",
        "Specialty / Premium",
        "Specialty & import avocado oil brands",
    ),
    "Specialty Food Importer / Distributor": (
        "Imported specialty foods; specialty oils",
        "Import specialty/gourmet oils including avocado",
        "Specialty / Premium",
        "Imported specialty oil brands",
    ),
    "Specialty Foodservice Distributor": (
        "Specialty foodservice; premium oils",
        "Chef/gourmet oil assortment for foodservice",
        "Premium / Specialty",
        "Premium avocado & finishing oil brands",
    ),
    "Specialty Ingredients Distributor": (
        "Food & nutrition ingredients; specialty oils",
        "Ingredient-grade oils & lipids for manufacturers",
        "Specialty ingredients Mid-to-Premium",
        "Principal specialty oil & ingredient suppliers",
    ),
    "Specialty Chemical / Ingredients Distributor": (
        "Specialty chemicals; food ingredients; oils",
        "Chemical & food lipid ingredient distribution",
        "Industrial / Specialty",
        "Principal chemical & oil ingredient brands",
    ),
    "Specialty Chemical Distributor": (
        "Specialty chemicals; fine chemicals",
        "Chemical distribution (oils as specialty chem adjacency)",
        "Industrial / Specialty",
        "Principal specialty chemical suppliers",
    ),
    "Oleochemical / Ingredients Distributor": (
        "Oleochemicals; vegetable oil derivatives",
        "Bio-based oleochemical & oil derivative supply",
        "Industrial / Specialty",
        "Oleochemical principal lines",
    ),
    "Chemical Distributor": (
        "Chemical distribution; process oils",
        "Industrial chemical & oil product distribution",
        "Industrial / Mid-market",
        "Principal chemical suppliers",
    ),
    "Chemical / Ingredients Trader": (
        "Chemical & ingredient trading; oils",
        "Trading of chemical intermediates & oil products",
        "Industrial / Mid-market",
        "Traded principal oil/chem brands",
    ),
    "Fresh Produce Distributor": (
        "Fresh produce; avocados; related specialty oils",
        "Fresh avocado & produce (culinary oil as adjacency)",
        "Commodity / Mid-market fresh",
        "Produce house brands; grower avocado labels",
    ),
    "Fresh Produce Cooperative / Exporter": (
        "Fresh produce cooperative; avocado export",
        "Grower-owned avocado & produce export",
        "Commodity / Mid-market fresh",
        "Cooperative grower brands",
    ),
    "Fresh Produce Grower / Distributor": (
        "Avocado growing & packing; fresh produce",
        "Grower-packer avocado supply (oil as value-add)",
        "Mid-market / Fresh specialty",
        "Estate / grower avocado brands",
    ),
    "Fresh Produce Cooperative": (
        "Fresh produce cooperative; avocados",
        "Cooperative fresh avocado marketing",
        "Commodity / Mid-market fresh",
        "Cooperative produce brands",
    ),
    "Specialty Produce / Foodservice Distributor": (
        "Specialty produce; foodservice; specialty oils",
        "Chef produce + specialty culinary oils",
        "Specialty / Premium",
        "Specialty produce & oil brands",
    ),
    "Organic Produce Distributor": (
        "Organic produce; organic avocados",
        "Organic avocado & produce distribution",
        "Organic / Premium",
        "Organic grower brands",
    ),
    "Bulk Avocado Oil Supplier": (
        "Bulk avocado oil; private-label oil packing",
        "B2B bulk avocado oil supply",
        "Commodity–Specialty bulk",
        "Private-label / sourced bulk avocado oil",
    ),
    "Bulk Oils / Ingredients Supplier": (
        "Bulk vegetable oils; avocado oil; ingredients",
        "Bulk oil wholesale for food/cosmetic makers",
        "Mid-market / Bulk specialty",
        "Bulk private label; sourced oil grades",
    ),
    "Bulk Essential / Carrier Oils Supplier": (
        "Bulk carrier & essential oils; avocado oil",
        "Wholesale carrier-grade avocado oil",
        "Specialty / Mid-to-Premium bulk",
        "House bulk oil lines",
    ),
    "Natural / Cosmetic Ingredients Wholesale": (
        "Cosmetic ingredients; carrier oils; avocado oil",
        "Natural cosmetic & DIY oil ingredients",
        "Specialty / Mid-to-Premium",
        "House natural oil lines; cosmetic ingredient brands",
    ),
    "Cosmetic Ingredients Supplier": (
        "Cosmetic raw materials; carrier oils",
        "Cosmetic-grade avocado & botanical oils",
        "Specialty / Mid-to-Premium",
        "Cosmetic ingredient / oil lines",
    ),
    "Soap / Cosmetic Ingredients Supplier": (
        "Soapmaking & cosmetic ingredients; oils",
        "Craft & cosmetic avocado/carrier oils",
        "Specialty / Mid-market",
        "Craft oil & soapmaking brands",
    ),
    "Soap Making Supplies Wholesale": (
        "Soapmaking supplies; carrier oils",
        "Wholesale soap oils including avocado",
        "Specialty / Mid-market",
        "Soapmaking oil lines",
    ),
    "D2C / Specialty Oils Brand": (
        "Branded specialty avocado / culinary oils",
        "Direct brand specialty avocado oil",
        "Premium / Specialty",
        "Own brand avocado oil",
    ),
    "Specialty Oils Brand / Distributor": (
        "Branded specialty oils; avocado oil",
        "Brand-led specialty culinary oils",
        "Premium / Specialty",
        "Own specialty oil brand portfolio",
    ),
    "Specialty Oils Producer / Distributor": (
        "Specialty oil production & distribution",
        "Producer-distributor of specialty culinary oils",
        "Premium / Specialty",
        "Own estate / brand oils",
    ),
    "Vegetable Oils / Ingredients Distributor": (
        "Vegetable oils; specialty edible oils; avocado oil",
        "Merchant distribution of vegetable & specialty oils",
        "Specialty / Mid-to-Premium",
        "House & principal vegetable oil brands",
    ),
    "Vegetable Oils Distributor": (
        "Vegetable oils; edible oils",
        "Edible vegetable oil wholesale",
        "Mid-market / Specialty",
        "Vegetable oil portfolio brands",
    ),
    "Edible Oils Brand Channel": (
        "Branded edible oils",
        "Branded cooking-oil retail channel",
        "Mainstream / Mid-market",
        "Brand edible oils",
    ),
    "Edible Oils / Foods Producer & Distributor": (
        "Edible oils production & brand distribution",
        "Integrated edible-oil producer-distributor",
        "Mid-to-Premium branded",
        "Own edible oil brands",
    ),
    "Private Label Brand Channel": (
        "Private-label edible oils; avocado oil",
        "Retailer private-label oil program",
        "Value / Mid-market",
        "Retailer private-label oil brands",
    ),
    "FMCG Distributor": (
        "FMCG; packaged foods; edible oils",
        "FMCG distribution including packaged oils",
        "Mainstream / Mid-market",
        "Principal FMCG & oil brands",
    ),
    "FMCG / Food Distributor": (
        "FMCG & food distribution; edible oils",
        "Regional FMCG/food oils assortment",
        "Mainstream / Mid-market",
        "Principal FMCG food & oil brands",
    ),
    "Market Expansion / Ingredients Distributor": (
        "Market expansion; F&B ingredients; oils",
        "Asia market-expansion distribution of ingredients/oils",
        "Mid-market / Specialty",
        "Principal F&B ingredient & oil brands",
    ),
    "Agribusiness / Oils Distributor": (
        "Agribusiness edible oils & fats",
        "Large-scale edible oils trading & distribution",
        "Commodity / Mid-market",
        "Group edible oil brands",
    ),
    "Oils & Fats Distributor": (
        "Oils & fats; edible oils",
        "Oils & fats wholesale distribution",
        "Commodity / Mid-market",
        "Group oils & fats brands",
    ),
    "Oils Distributor": (
        "Edible oils distribution",
        "Edible oil wholesale channel",
        "Mid-market",
        "Edible oil brand portfolio",
    ),
    "Food Ingredients Distributor": (
        "Food ingredients; specialty oils",
        "Ingredient oils for food manufacturers",
        "Mid-market / Specialty ingredients",
        "Principal food ingredient brands",
    ),
    "Food Ingredients / Agri Distributor": (
        "Agri food ingredients; oils",
        "Agribusiness ingredient & oil distribution",
        "Commodity / Mid-market",
        "Group ingredient & oil brands",
    ),
    "Trading House Food Distribution": (
        "Trading-house food & oils distribution",
        "Sogo shosha-style food/oils trading",
        "Commodity / Mid-to-Premium",
        "Trading-house food & oil principals",
    ),
    "Trading House": (
        "General trading; food & chemicals",
        "Diversified trading including oils/ingredients",
        "Commodity / Industrial",
        "Trading-house principal lines",
    ),
    "Retail / Warehouse Club Channel": (
        "Warehouse club / retail; edible oils; avocado oil",
        "Club & mass retail oil assortment",
        "Value / Mid-market",
        "Club private label; national oil brands",
    ),
    "Warehouse Club Wholesale": (
        "Warehouse club wholesale; edible oils",
        "Club-pack culinary oils",
        "Value / Mid-market club",
        "Club private label; brand oils",
    ),
    "Retail Distribution": (
        "Retail distribution; edible oils",
        "Retail banner oil assortment",
        "Mainstream / Mid-market",
        "Banner private label; national oils",
    ),
    "Hypermarket Distribution": (
        "Hypermarket; edible oils; avocado oil",
        "Hypermarket packaged oils assortment",
        "Value / Mid-market",
        "Banner & national oil brands",
    ),
    "Grocery Retail Distribution": (
        "Grocery retail; edible oils",
        "Grocery retail oil assortment",
        "Mainstream / Mid-market",
        "Banner private label; brand oils",
    ),
    "Specialty Grocery Retailer": (
        "Specialty grocery; private-label oils",
        "Specialty retail culinary oils",
        "Specialty / Mid-to-Premium",
        "Private-label specialty oils",
    ),
    "Food & Ingredients Distributor": (
        "Food & ingredients; edible oils; avocado oil",
        "General food/ingredient distribution including avocado oil",
        "Mid-market / Specialty",
        "Private label & regional avocado oil brands",
    ),
    "Wholesale Distributor": (
        "Wholesale foods; edible oils",
        "General wholesale edible oils channel",
        "Mid-market",
        "Private label & brand edible oils",
    ),
    "Food Distributor": (
        "Food distribution; edible oils",
        "General food distribution oils assortment",
        "Mid-market",
        "Private label & brand oils",
    ),
    "QSR / Foodservice Distributor": (
        "QSR foodservice; edible oils",
        "Chain restaurant oil & shortening supply",
        "Contract / Mid-market",
        "Chain specs; approved oil brands",
    ),
    "Beverage / Foodservice Distribution": (
        "Beverage & foodservice distribution",
        "Foodservice adjacency (oils via beverage/FS network)",
        "Mid-market",
        "Network private label & brand SKUs",
    ),
    "Foodservice Redistributor": (
        "Foodservice redistributor; edible oils",
        "Redistribution of oils to regional FS distributors",
        "Mid-market / Full range",
        "Redistributed national oil brands",
    ),
    "Dairy / Foodservice Distributor": (
        "Dairy & foodservice; edible oils adjacency",
        "Dairy foodservice channel with oil SKUs",
        "Mid-market",
        "Dairy brand oils & culinary oils",
    ),
    "Frozen Food Distributor": (
        "Frozen foods; grocery; edible oils adjacency",
        "Frozen/foodservice channel with oil assortment",
        "Mid-market",
        "Frozen brand & oil SKUs",
    ),
    "Olive Oil Brand Distribution": (
        "Olive oil & specialty edible oils",
        "Branded olive-oil channel; specialty oil adjacency",
        "Mid-to-Premium branded",
        "Olive oil brand portfolio",
    ),
    "Commodity Marketing Board": (
        "Avocado commodity marketing & education",
        "Industry promotion of avocado (incl. oil usage)",
        "N/A (association)",
        "Category promotion (not brand retail)",
    ),
    "Social Food Distribution": (
        "Social / public food distribution",
        "Public-program food staples (oils as staple SKU)",
        "Value / Social program",
        "Program staple brands",
    ),
    "Produce Procurement Cooperative": (
        "Foodservice produce procurement; avocados",
        "Cooperative buying of fresh avocado for foodservice",
        "Foodservice Mid-market",
        "Cooperative produce programs",
    ),
    "Organic Ingredients Trader": (
        "Organic ingredients; organic oils",
        "Organic avocado & specialty oil trading",
        "Organic / Premium",
        "Organic supplier network brands",
    ),
    "Nutraceutical Ingredients Distributor": (
        "Nutraceutical ingredients; specialty oils",
        "Supplement-grade oils & ingredients",
        "Specialty / Premium ingredients",
        "Nutra ingredient principals",
    ),
    "Sports Nutrition Brand / Wholesale": (
        "Sports nutrition; specialty oils adjacency",
        "Sports nutrition wholesale (oils as related SKU)",
        "Specialty / Mid-to-Premium",
        "Own sports nutrition brands",
    ),
    "Supplements Brand / Wholesale": (
        "Dietary supplements wholesale",
        "Supplement brand wholesale channel",
        "Specialty / Mid-market",
        "Own supplement brands",
    ),
    "Laboratory / Fine Chemical Distributor": (
        "Lab & fine chemicals",
        "Lab chemical supply (oils as fine-chem SKU)",
        "Industrial / Specialty",
        "Lab chemical catalog brands",
    ),
    "Natural Oils / Ingredients Wholesale": (
        "Natural oils wholesale; avocado oil",
        "Wholesale natural & carrier avocado oil",
        "Specialty / Mid-to-Premium",
        "House natural oil lines",
    ),
    "Essential Oils / Ingredients Wholesale": (
        "Essential & carrier oils; avocado oil",
        "Wholesale aromatherapy & carrier oils",
        "Specialty / Mid-market",
        "House essential/carrier oil lines",
    ),
    "Essential Oils / Craft Ingredients": (
        "Craft & aromatherapy oils",
        "Craft-market avocado & carrier oils",
        "Specialty / Mid-market",
        "Craft oil brand lines",
    ),
    "Natural Ingredients Wholesale": (
        "Natural raw ingredients; oils",
        "Natural ingredient oils for makers",
        "Specialty / Mid-market",
        "Natural ingredient lines",
    ),
    "Natural Ingredients / Oils Distributor": (
        "Natural oils & ingredients distribution",
        "Natural oil wholesale distribution",
        "Specialty / Mid-to-Premium",
        "Natural oil brand portfolio",
    ),
    "Natural Ingredients Retail / Wholesale": (
        "Natural ingredients retail/wholesale; oils",
        "Retail-wholesale natural avocado oil",
        "Specialty / Mid-market",
        "House natural oil brands",
    ),
    "Natural Ingredients Supplier": (
        "Natural ingredient supply; oils",
        "Supply of natural oils to makers/brands",
        "Specialty / Mid-market",
        "Natural ingredient lines",
    ),
    "Cosmetic Ingredients Manufacturer / Wholesale": (
        "Cosmetic bases & oils manufacturing/wholesale",
        "Manufactured cosmetic oils including avocado",
        "Specialty / Mid-to-Premium",
        "Own cosmetic oil/base lines",
    ),
    "Foodservice Supplies Distributor": (
        "Foodservice disposables & supplies; oils adjacency",
        "FS supplies with limited edible oil SKUs",
        "Mid-market",
        "Supplies brands; limited oil brands",
    ),
    "Foodservice / Facility Services Distribution": (
        "Foodservice & facility services; oils",
        "Contract FS distribution including culinary oils",
        "Contract / Mid-market",
        "Contract-approved oil brands",
    ),
    "Foodservice / Cash & Carry": (
        "Foodservice cash & carry; edible oils",
        "FS/cash-carry culinary oils",
        "Value / Mid-market",
        "Private label & brand oils",
    ),
    "Drinks / Foodservice Distributor": (
        "Drinks & foodservice; edible oils adjacency",
        "Drinks FS network with grocery/oil SKUs",
        "Mid-market",
        "Drinks brands; culinary oil SKUs",
    ),
    "Convenience Wholesale": (
        "Convenience wholesale; edible oils",
        "C-store wholesale oils assortment",
        "Value / Mid-market",
        "C-store private label; brand oils",
    ),
    "Grocery / Convenience Wholesale": (
        "Grocery & convenience wholesale; oils",
        "Grocery/c-store wholesale oils",
        "Value / Mid-market",
        "Private label & brand oils",
    ),
}


def norm_key(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower().replace("&", " and ")
    s = s.replace("(", " ").replace(")", " ").replace("/", " ")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def is_generic(val: str, generic_set: set[str]) -> bool:
    return (val or "").strip().lower() in generic_set


WEAK_CORE_EXACT = {
    "edible oils & fats",
    "edible oils",
    "fresh fruits and vegetables",
    "specialty chemicals",
    "organic food products",
    "food distribution, wholesale, foodservice",
    "warehouse club wholesale; edible oils",
}
WEAK_FOCUS_EXACT = {
    "avocado oil",
    "premium avocado oil",
    "cold-pressed avocado oil",
    "club-pack culinary oils",
}
WEAK_PRICE = {
    "specialty / mid-to-premium",
    "mid-range",
    "mid-range to premium",
    "mid-market",
}


def weak_core(val: str) -> bool:
    v = (val or "").strip()
    if not v:
        return True
    low = v.lower()
    if is_generic(v, GENERIC_CORE) or low in WEAK_CORE_EXACT:
        return True
    if len(v) < 20 and "avocado" in low and ";" not in v:
        return True
    return False


def weak_focus(val: str) -> bool:
    v = (val or "").strip()
    if not v:
        return True
    low = v.lower()
    if is_generic(v, GENERIC_FOCUS) or low in WEAK_FOCUS_EXACT:
        return True
    if len(v) < 28:
        return True
    if low.startswith("offers a variety"):
        return True
    return False


TYPE_DEFAULT_BRANDS = {
    "private label & regional avocado oil brands",
    "regional & private-label avocado / edible oil brands",
    "house/private-label oils; national avocado oil brands",
    "private label & brand avocado/culinary oils",
    "produce house brands; grower avocado labels",
    "club private label; brand oils",
    "cash & carry private label; brand edible oils",
    "retailer private label; national oil brands",
    "principal specialty oil & ingredient suppliers",
    "member banner private label; brand oils",
}


def should_replace_brands(val: str) -> bool:
    v = (val or "").strip()
    if not v or is_generic(v, GENERIC_BRANDS):
        return True
    return v.lower() in TYPE_DEFAULT_BRANDS


def bidfood_patch(company: str) -> tuple[str, str, str, str]:
    """Region-specific Bidfood values (do not copy UK text to all markets)."""
    key = norm_key(company)
    if key == "bidfood uk":
        return VERIFIED["bidfood uk"]
    region = re.sub(r"^(bidfood|bidvest)\s+", "", key).strip() or "international"
    special = {
        "poland farutex": "Poland (Farutex)",
        "czech": "Czech Republic",
        "new zealand": "New Zealand",
        "south africa": "South Africa",
        "netherlands": "Netherlands",
        "singapore": "Singapore",
        "australia": "Australia",
        "belgium": "Belgium",
        "chile": "Chile",
        "portugal": "Portugal",
        "spain": "Spain",
        "turkey": "Turkey",
        "asia": "Asia",
        "baltics": "Baltics",
    }
    region_label = special.get(region, region.title())
    return (
        "Broadline foodservice; edible oils; avocado oil",
        f"Bidfood {region_label} foodservice culinary oils",
        "Full range (Value–Premium)",
        "Bidfood private label; brand avocado oils",
    )


def infer_from_name(company: str, btype: str) -> tuple[str, str, str, str] | None:
    n = company.lower()
    if "bidfood" in n or "bidvest" in n:
        return bidfood_patch(company)
    if n.startswith("dksh"):
        return BT_TEMPLATES["Market Expansion / Ingredients Distributor"]
    if (
        "avocado oil" in n
        or n.startswith("bulk avocado")
        or "refined avocado" in n
        or "organic avocado" in n
        or "wholesale avocado" in n
        or "tropical avocado" in n
    ):
        return BT_TEMPLATES["Bulk Avocado Oil Supplier"]
    if any(x in n for x in ("brenntag", "barentz", "azelis", "imcd")):
        return BT_TEMPLATES["Specialty Ingredients Distributor"]
    return None


# Aliases for Business Type strings missing exact BT_TEMPLATES keys
BT_ALIASES: dict[str, str] = {
    "Foodservice / Wholesale Distributor": "Foodservice Distributor",
    "Foodservice / Wholesale": "Foodservice Wholesale",
    "Foodservice Distributor / Cash & Carry": "Cash & Carry / Foodservice Wholesale",
    "Foodservice / Drinks Distributor": "Foodservice Distributor",
    "Specialty Food Ingredients Distributor": "Specialty Ingredients Distributor",
    "Food Ingredients Manufacturer / Distributor": "Food Ingredients Distributor",
    "Food Ingredients Broker / Distributor": "Food Ingredients Distributor",
    "Oils / Food Ingredients Distributor": "Food Ingredients Distributor",
    "Oils & Fats Ingredients Distributor": "Food Ingredients Distributor",
    "Fresh Produce Producer / Distributor": "Fresh Produce Grower / Distributor",
    "Fresh Produce Importer / Distributor": "Fresh Produce Distributor",
    "Specialty Produce Distributor": "Fresh Produce Distributor",
    "Specialty Oils Distributor": "Specialty Food Distributor",
    "Specialty Spices Retail / Wholesale": "Specialty Food Distributor",
    "Natural / Organic Distributor": "Natural / Specialty Grocery Wholesale",
    "Natural Oils Brand Channel": "D2C / Specialty Oils Brand",
    "Oleochemical Distributor": "Specialty Chemical Distributor",
    "Chemical / Ingredients Distributor": "Chemical Distributor",
    "Chemical Distributor / Manufacturer": "Chemical Distributor",
    "Palm / Oils Producer / Distributor": "Bulk Oils / Ingredients Supplier",
    "Bulk Natural Oils Distributor": "Bulk Oils / Ingredients Supplier",
    "Edible Oils / Food Distributor": "FMCG / Food Distributor",
    "Retail / Wholesale": "Grocery Wholesale / Retail",
    "Warehouse Club": "Warehouse Club Wholesale",
    "Cash & Carry / Warehouse Club": "Cash & Carry / Foodservice Wholesale",
}


def resolve(company: str, btype: str, core: str, focus: str, price: str, brands: str) -> dict[str, str]:
    key = norm_key(company)
    patch: dict[str, str] = {}

    verified = VERIFIED.get(key)
    if not verified and ("bidfood" in key or "bidvest" in key):
        verified = bidfood_patch(company)
    if not verified and key.startswith("dksh"):
        verified = (
            "Market expansion; F&B ingredients; edible oils",
            f"DKSH local market expansion — oils & F&B ingredients ({company})",
            "Mid-market / Specialty",
            "Principal F&B and oil brands",
        )
    if not verified and key.startswith("sysco"):
        verified = VERIFIED["sysco"]
    if not verified and key.startswith("transgourmet"):
        verified = (
            "Foodservice wholesale; edible oils; avocado oil",
            "Transgourmet/Coop foodservice culinary oils",
            "Mid-market / Full range",
            "Transgourmet private label; brand oils",
        )

    if verified:
        vc, vf, vp, vb = verified
        patch["Core Categories"] = vc
        patch["Specialty Focus"] = vf
        patch["Price Segment"] = vp
        if should_replace_brands(brands):
            patch["Key Brands Represented"] = vb
        else:
            patch["Key Brands Represented"] = brands
        return patch

    named = infer_from_name(company, btype)
    mapped_bt = BT_ALIASES.get(btype, btype)
    template = named or BT_TEMPLATES.get(mapped_bt) or BT_TEMPLATES.get(btype)

    if template:
        tc, tf, tp, tb = template
        if weak_core(core):
            patch["Core Categories"] = tc
        if weak_focus(focus) or ("uk bidfood" in focus.lower() and "uk" not in company.lower()):
            patch["Specialty Focus"] = tf
        if not price.strip() or price.strip().lower() in WEAK_PRICE:
            patch["Price Segment"] = tp
        if is_generic(brands, GENERIC_BRANDS) or not brands.strip():
            patch["Key Brands Represented"] = tb
        return patch

    # fallback generic but slightly improved
    if weak_core(core):
        patch["Core Categories"] = "Edible oils; avocado oil; foodservice/retail oils"
    if weak_focus(focus):
        patch["Specialty Focus"] = f"Channel distribution of avocado & edible oils ({btype or 'distributor'})"
    if not price.strip() or price.strip().lower() in WEAK_PRICE:
        patch["Price Segment"] = "Mid-market / Specialty"
    if is_generic(brands, GENERIC_BRANDS) or not brands.strip():
        patch["Key Brands Represented"] = "Regional & private-label avocado / edible oil brands"
    return patch


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = PATH.with_name(f"{PATH.stem}_pre_categories_fix_{stamp}{PATH.suffix}")
    shutil.copy2(PATH, bak)

    wb = openpyxl.load_workbook(PATH)
    ws = wb["Avocado Oil Distributors"]
    hdr = [c.value for c in ws[1]]
    cols = {h: i + 1 for i, h in enumerate(hdr) if h}

    changed = 0
    samples = []
    for r in range(2, ws.max_row + 1):
        company = str(ws.cell(r, cols["Company"]).value or "").strip()
        if not company:
            continue
        btype = str(ws.cell(r, cols["Business Type"]).value or "").strip()
        core = str(ws.cell(r, cols["Core Categories"]).value or "").strip()
        focus = str(ws.cell(r, cols["Specialty Focus"]).value or "").strip()
        price = str(ws.cell(r, cols["Price Segment"]).value or "").strip()
        brands = str(ws.cell(r, cols["Key Brands Represented"]).value or "").strip()

        patch = resolve(company, btype, core, focus, price, brands)
        if not patch:
            continue
        for field, val in patch.items():
            old = str(ws.cell(r, cols[field]).value or "").strip()
            if old != val:
                ws.cell(r, cols[field]).value = val
                changed += 1
        if len(samples) < 20:
            samples.append((company, patch))

    # stats
    dist = {f: Counter() for f in ("Core Categories", "Specialty Focus", "Price Segment", "Key Brands Represented")}
    generic_left = Counter()
    tot = 0
    for r in range(2, ws.max_row + 1):
        company = str(ws.cell(r, cols["Company"]).value or "").strip()
        if not company:
            continue
        tot += 1
        for f, gset in (
            ("Core Categories", GENERIC_CORE),
            ("Specialty Focus", GENERIC_FOCUS),
            ("Key Brands Represented", GENERIC_BRANDS),
        ):
            v = str(ws.cell(r, cols[f]).value or "").strip()
            dist[f][v] += 1
            if is_generic(v, gset):
                generic_left[f] += 1
        pv = str(ws.cell(r, cols["Price Segment"]).value or "").strip()
        dist["Price Segment"][pv] += 1

    wb.save(PATH)
    print("backup", bak.name)
    print("cells_changed", changed, "companies", tot)
    print("generic_left", dict(generic_left))
    print("samples:")
    for co, patch in samples:
        print(" ", co, patch)
    for f, c in dist.items():
        print(f"--- {f} top ---")
        for v, n in c.most_common(10):
            print(f"  {n:3d} | {v[:100]}")


if __name__ == "__main__":
    main()
