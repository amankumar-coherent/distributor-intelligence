"""75 ProFood India markets (3 ecosystems x 25), India-only, distributors.

Built from the user's "ProFood India Expo_ CMI Dashboard Topics.xlsx"; the
definitions live as data in queries/profood_india_markets.json (one entry per
market: id, ecosystem, topic, definition, inclusion, exclusion), so a wording
fix there needs no code change.

CLASSIFICATION (same 4-way rule system as fire_markets.py):

  DISTRIBUTOR -- physical goods, stocking / wholesale / C&F distribution:
    every Food Processing (FP) and Food Ingredients (FI) market, plus
    FMT-15 Food Testing Instruments (stocked lab instruments).

  HYBRID -- machinery and automation, sold in India through authorized
  dealers, agents and system integrators as much as stocking distributors:
    the remaining Food Manufacturing Technology (FMT) markets.

  SOFTWARE -- channel partners / resellers of another vendor's platform:
    FMT-18 Food Manufacturing Execution Software, FMT-19 Food Traceability
    Software.

Each entry is (slug, market_name, product_scope, kind), the shape
market_batches.py expects.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

_DATA = Path(__file__).resolve().parent / "queries" / "profood_india_markets.json"

_SOFTWARE_IDS = {"FMT-18", "FMT-19"}
_DISTRIBUTOR_FMT_IDS = {"FMT-15"}

_LEAD = {
    "distributor": "PRODUCT SCOPE (required)",
    "hybrid": "PRODUCT/SOLUTION SCOPE (required)",
    "software": "SOFTWARE SCOPE (required)",
}
_ASK = {
    "distributor": "Name the {what} it distributes.",
    "hybrid": "Name the {what} it distributes, supplies as a dealer, or integrates.",
    "software": "Name the {what} platform it resells or implements.",
}


def _kind(market_id: str) -> str:
    if market_id in _SOFTWARE_IDS:
        return "software"
    if market_id.startswith("FMT-") and market_id not in _DISTRIBUTOR_FMT_IDS:
        return "hybrid"
    return "distributor"


def _short(topic: str) -> str:
    """'India Dairy Products Market' -> 'Dairy Products'."""
    return re.sub(r"^India\s+|\s+Market$", "", topic.strip())


def _slug(topic: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", _short(topic).lower()).strip("_")


def _scope(m: dict, kind: str) -> str:
    what = _short(m["topic"])
    # The "Market sizing is based on ..." sentence is about report revenue,
    # not who qualifies, and its "manufacturer-level revenue" wording pulls
    # a distributor search toward manufacturers.
    definition = re.sub(r"\s*Market sizing[^.]*\.", "", m["definition"].strip())
    return (
        f"{_LEAD[kind]}: {what.upper()} - {definition} "
        f"INCLUDE: {m['inclusion'].strip()} "
        f"{m['exclusion'].strip()} "
        + _ASK[kind].format(what=what.lower())
    )


def _load() -> tuple[tuple[str, str, str, str], ...]:
    data = json.loads(_DATA.read_text(encoding="utf-8"))["markets"]
    out = []
    for m in data:
        kind = _kind(m["id"])
        out.append((_slug(m["topic"]), f"{_short(m['topic'])} Market - India", _scope(m, kind), kind))
    slugs = [e[0] for e in out]
    if len(set(slugs)) != len(slugs):
        raise ValueError("duplicate food market slug")
    return tuple(out)


MARKETS_FOOD: tuple[tuple[str, str, str, str], ...] = _load()


def _definitions() -> dict[str, str]:
    """Client-facing market definition per slug, straight from the workbook's
    definition / inclusion / exclusion columns (row 2 of the final Excel)."""
    data = json.loads(_DATA.read_text(encoding="utf-8"))["markets"]
    return {_slug(m["topic"]): " ".join(x.strip() for x in (
        m["definition"], "Includes: " + m["inclusion"], m["exclusion"]) if x.strip())
        for m in data}


DEFINITIONS: dict[str, str] = _definitions()

# Closest neighbours across the three tabs: a generic search for one drifts
# into the other, so each scope names the market it must NOT become.
SIBLINGS_FOOD = {
    "dairy_products": "dairy ingredients, ice cream and frozen desserts",
    "ice_cream_and_frozen_desserts": "dairy products",
    "dairy_ingredients": "dairy products, food protein ingredients",
    "chocolate_products": "cocoa ingredients, sugar confectionery",
    "cocoa_ingredients": "chocolate products",
    "sugar_confectionery": "chocolate products",
    "savoury_snacks": "frozen potato products",
    "frozen_potato_products": "savoury snacks, processed fruits and vegetables",
    "processed_fruits_and_vegetables": "fruit preparations and purees, vegetable ingredients",
    "fruit_preparations_and_purees": "processed fruits and vegetables",
    "vegetable_ingredients": "processed fruits and vegetables",
    "ready_to_eat_meals": "ready-to-cook foods, foodservice prepared foods",
    "ready_to_cook_foods": "ready-to-eat meals",
    "foodservice_prepared_foods": "ready-to-eat meals",
    "pasta_and_noodles": "breakfast cereals",
    "breakfast_cereals": "pasta and noodles",
    "packaged_rice_products": "packaged flour and atta, pulses and dal products",
    "packaged_flour_and_atta": "packaged rice products",
    "pulses_and_dal_products": "packaged rice products",
    "spices_and_seasoning_products": "savory taste enhancers, sauces and condiments",
    "sauces_and_condiments": "pickles and preserved foods, spices and seasoning products",
    "pickles_and_preserved_foods": "sauces and condiments",
    "edible_oils": "specialty food fats",
    "specialty_food_fats": "edible oils",
    "non_alcoholic_beverages": "packaged drinking water",
    "packaged_drinking_water": "non-alcoholic beverages",
    "high_intensity_sweeteners": "bulk nutritive sweeteners",
    "bulk_nutritive_sweeteners": "high-intensity sweeteners",
    "food_emulsifiers": "hydrocolloids and thickeners",
    "hydrocolloids_and_thickeners": "food emulsifiers, modified food starches",
    "modified_food_starches": "hydrocolloids and thickeners",
    "food_preservatives": "food acidulants",
    "food_acidulants": "food preservatives",
    "starter_cultures_and_probiotics": "food enzymes",
    "food_enzymes": "starter cultures and probiotics",
    "food_protein_ingredients": "dairy ingredients, egg ingredients",
    "egg_ingredients": "food protein ingredients",
    "vitamin_premixes": "mineral fortification ingredients",
    "mineral_fortification_ingredients": "vitamin premixes",
    "savory_taste_enhancers": "spices and seasoning products, food flavourings",
    "food_flavourings": "savory taste enhancers",
    "food_freezing_equipment": "industrial food refrigeration systems",
    "industrial_food_refrigeration_systems": "food freezing equipment",
    "food_filling_and_dosing_machinery": "food form-fill-seal machinery",
    "food_form_fill_seal_machinery": "food filling and dosing machinery, food end-of-line packaging equipment",
    "food_end_of_line_packaging_equipment": "food form-fill-seal machinery",
    "food_inspection_and_sorting_equipment": "food testing instruments",
    "food_testing_instruments": "food inspection and sorting equipment",
    "food_plant_automation_and_control_systems": "food manufacturing execution software, food processing robotics",
    "food_processing_robotics": "food plant automation and control systems",
    "food_manufacturing_execution_software": "food traceability software, food plant automation and control systems",
    "food_traceability_software": "food manufacturing execution software",
    "food_conveying_systems": "bulk ingredient handling and storage systems",
    "bulk_ingredient_handling_and_storage_systems": "food conveying systems",
    "food_thermal_processing_equipment": "food plant steam and heat utility equipment",
    "food_plant_steam_and_heat_utility_equipment": "food thermal processing equipment",
    "food_plant_hygienic_pumps_and_valves": "food plant cleaning and sanitation systems",
    "food_plant_cleaning_and_sanitation_systems": "food plant hygienic pumps and valves",
    "food_drying_and_dehydration_equipment": "food thermal processing equipment",
    "food_mixing_and_blending_equipment": "food size-reduction equipment",
    "food_size_reduction_equipment": "food mixing and blending equipment",
}

_unknown = set(SIBLINGS_FOOD) - {e[0] for e in MARKETS_FOOD}
if _unknown:
    raise ValueError(f"SIBLINGS_FOOD names unknown slugs: {sorted(_unknown)}")


def by_slug_food(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_FOOD:
        if entry[0] == slug:
            return entry
    raise KeyError(
        f"unknown food market slug: {slug}. "
        f"Valid: {', '.join(m[0] for m in MARKETS_FOOD)}"
    )
