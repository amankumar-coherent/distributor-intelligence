"""Company names must survive both Kompass listing URL shapes.

The same company appears twice on one page in two forms:
    JSON-LD   https://www.kompass.com/c/<slug>/<id>
    HTML link https://www.kompass.com/z/ww/c/<slug>/<id>

The JSON-LD carries the NAME; the HTML gives the link. Joining them on the
full URL silently failed on any listing that used the /z/ww/ prefix, so
every candidate came back nameless — and nameless candidates are skipped
before screening. Observed live: LNG showed 2,200 screened, 0 accepted,
because 2,223 of its 4,141 candidates had no name.

Joining on the stable company id fixes it.
"""
from __future__ import annotations

from vendor_intel.discovery.kompass_distributors import _extract_company_links

# Kompass nests the list under mainEntity — matching the real page shape
# matters here, because a fixture that parses differently would test
# nothing.
JSONLD = """
<script type="application/ld+json">
{"mainEntity":{"@type":"ItemList","itemListElement":[
 {"item":{"@type":"Organization","name":"Terminale GNL Adriatico",
   "url":"https://www.kompass.com/c/terminale-gnl-adriatico/it0403920",
   "address":{"addressCountry":"IT","addressLocality":"Milano"},
   "telephone":"+39 02 1234567"}},
 {"item":{"@type":"Organization","name":"Vienna Components Trading",
   "url":"https://www.kompass.com/c/vienna-components-trading/at0112233",
   "address":{"addressCountry":"AT"}}}
]}}
</script>
"""

# Links use the /z/ww/ prefix — a different string for the same company.
LINKS_PREFIXED = """
<a href="/z/ww/c/terminale-gnl-adriatico/it0403920/">x</a>
<a href="/z/ww/c/vienna-components-trading/at0112233/">y</a>
"""

# Some listings use the bare form that already matched.
LINKS_BARE = """
<a href="/c/terminale-gnl-adriatico/it0403920/">x</a>
<a href="/c/vienna-components-trading/at0112233/">y</a>
"""

BASE = "https://www.kompass.com/z/ww/a/natural-gas/0746005/"


class TestPrefixedUrls:
    def test_names_resolve_despite_the_z_ww_prefix(self):
        got = _extract_company_links(JSONLD + LINKS_PREFIXED, BASE)
        names = sorted(c.name for c in got)
        assert names == ["Terminale GNL Adriatico", "Vienna Components Trading"]

    def test_no_candidate_is_left_nameless(self):
        """A nameless candidate is dropped before screening, so this is the
        difference between a company being evaluated and vanishing."""
        got = _extract_company_links(JSONLD + LINKS_PREFIXED, BASE)
        assert all(c.name.strip() for c in got)

    def test_other_jsonld_facts_come_across_too(self):
        got = _extract_company_links(JSONLD + LINKS_PREFIXED, BASE)
        by_name = {c.name: c for c in got}
        t = by_name["Terminale GNL Adriatico"]
        assert t.country_hint == "IT"
        assert t.city_hint == "Milano"
        assert t.phone_hint == "+39 02 1234567"


class TestNoRegression:
    def test_bare_urls_still_work(self):
        got = _extract_company_links(JSONLD + LINKS_BARE, BASE)
        assert sorted(c.name for c in got) == [
            "Terminale GNL Adriatico",
            "Vienna Components Trading",
        ]

    def test_link_without_matching_jsonld_stays_nameless(self):
        """Absence of data must not be invented — an unmatched link has no
        name and the screening step is right to skip it."""
        got = _extract_company_links(JSONLD + '<a href="/z/ww/c/unknown-co/xx999999/">z</a>', BASE)
        unknown = [c for c in got if "xx999999" in c.profile_url]
        assert len(unknown) == 1
        assert unknown[0].name == ""

    def test_duplicate_links_are_collapsed(self):
        got = _extract_company_links(JSONLD + LINKS_PREFIXED + LINKS_PREFIXED, BASE)
        assert len(got) == 2
