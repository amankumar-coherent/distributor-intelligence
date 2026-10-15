"""Brazil Reverse Osmosis Market, Brazil-only distributors.

Built from the user's market brief (2026-10-13): market definition, system
types, components, use cases, end-use industries and exclusions. Physical-goods
distributor rule: companies that stock and resell RO membranes, pumps,
pressure vessels, pretreatment, controls, energy recovery and RO systems.
System integrators and engineering contractors are excluded by that rule.

Same shape as the other batch modules: (slug, market_name, product_scope,
kind), plus BRAZIL_REGIONS, the states / metro areas discovery rotates
through (ASCII names on purpose: they travel in command lines and log files).
"""
from __future__ import annotations

MARKETS_BRAZIL_RO: tuple[tuple[str, str, str, str], ...] = (
    (
        "brazil_reverse_osmosis",
        "Brazil Reverse Osmosis Market",
        # Kept compact: discovery sends system prompt + this text in ONE Google
        # URL (budget 7,550 chars); see wound_markets.py for what over-long
        # scope text does (reply format cut off, 0 companies).
        "PRODUCT SCOPE (required): REVERSE OSMOSIS (RO) EQUIPMENT AND "
        "COMPONENTS - RO systems (seawater, brackish water, industrial / "
        "process water, wastewater, residential) and RO components: "
        "membranes, high-pressure pumps, pressure vessels, pretreatment, "
        "controls and monitoring, energy recovery. For potable, process, "
        "boiler and high-purity water, water reuse and desalination, "
        "serving food and beverage, pharma, chemicals, power, oil and gas, "
        "mining, pulp and paper and electronics. EXCLUDE conventional "
        "filtration with no RO, "
        "standalone ultrafiltration, microfiltration or UV, and water "
        "treatment chemicals unrelated to RO. Terms: distribuidor, osmose "
        "reversa. Name the RO products and brands it distributes.",
        "distributor",
    ),
)

SIBLINGS_BRAZIL_RO: dict[str, str] = {}

BRAZIL_REGIONS: tuple[str, ...] = (
    "Sao Paulo (Sao Paulo, Campinas, Guarulhos, Santos)",
    "Sao Paulo capital and metro (Sao Paulo, Osasco, Santo Andre, Sao Bernardo do Campo)",
    "Sao Paulo interior (Campinas, Ribeirao Preto, Sorocaba, Sao Jose dos Campos)",
    "Rio de Janeiro (Rio de Janeiro, Niteroi, Macae)",
    "Minas Gerais (Belo Horizonte, Uberlandia, Contagem, Betim)",
    "Parana (Curitiba, Londrina, Maringa, Cascavel)",
    "Rio Grande do Sul (Porto Alegre, Caxias do Sul, Pelotas)",
    "Santa Catarina (Joinville, Florianopolis, Blumenau)",
    "Bahia (Salvador, Camacari, Feira de Santana)",
    "Pernambuco (Recife, Ipojuca)",
    "Ceara (Fortaleza)",
    "Distrito Federal (Brasilia)",
    "Goias (Goiania, Anapolis)",
    "Espirito Santo (Vitoria, Serra)",
    "Amazonas (Manaus)",
    "Para (Belem)",
    "Mato Grosso (Cuiaba, Rondonopolis)",
    "Mato Grosso do Sul (Campo Grande)",
    "Maranhao (Sao Luis)",
    "Rio Grande do Norte (Natal)",
    "Paraiba (Joao Pessoa, Campina Grande)",
    "Alagoas (Maceio)",
    "Sergipe (Aracaju)",
    "Piaui (Teresina)",
    "Tocantins (Palmas)",
    "Rondonia (Porto Velho)",
    "Acre (Rio Branco)",
    "Amapa (Macapa)",
    "Roraima (Boa Vista)",
    # Second pass: industrial and port cities, where RO demand concentrates.
    "Juiz de Fora", "Uberaba", "Sao Jose do Rio Preto", "Bauru", "Santos port area",
    "Joinville industrial area", "Blumenau", "Itajai", "Paulinia", "Cubatao",
    "Camacari petrochemical area", "Suape", "Resende", "Volta Redonda",
    "Contagem", "Betim", "Ipatinga", "Canoas", "Gravatai", "Cuiaba", "Rondonopolis",
    "Parauapebas", "Carajas", "Aracruz", "Sao Jose dos Pinhais", "Ponta Grossa",
    "Foz do Iguacu", "Chapeco", "Passo Fundo", "Petrolina", "Mossoro",
)


# Client-facing market definition (the user's own brief, 2026-10-13), shown
# in row 2 of the final Excel instead of a generic rule sentence.
DEFINITIONS: dict[str, str] = {
    "brazil_reverse_osmosis": (
        "The Reverse Osmosis Market covers the Brazil market for reverse osmosis (RO) "
        "systems, equipment and components used to remove dissolved salts, impurities "
        "and contaminants from water through membrane-based filtration. These solutions "
        "serve residential, industrial, commercial and municipal requirements for "
        "potable water production, process water, high-purity water, wastewater "
        "treatment and water reuse. Excludes conventional filtration without RO, "
        "standalone ultrafiltration, microfiltration and UV when not part of an RO "
        "system, water chemicals unrelated to RO, general water supply and sanitation "
        "services, and equipment used only in non-RO treatment."
    ),
}


def by_slug_water(slug: str) -> tuple[str, str, str, str]:
    for entry in MARKETS_BRAZIL_RO:
        if entry[0] == slug:
            return entry
    raise KeyError(f"unknown water market slug: {slug}")
