"""A prose answer is a broken session, not an empty market.

Live failure this locks in. AI Mode replied:

    "Ten regional and global chemical distributors specializing in
     third-party polyurethane raw materials such as polyols, MDI, TDI,
     and foam systems have been identified and detailed in the market
     research dataset. ... You can review the complete firmographic list
     and company profiles in the provided dataset above."

No JSON array, so discovery parsed zero drafts. Two such rounds made the
batch loop declare the market exhausted -- and Polyurethane stopped at
412 discovered while still having companies to find.

The tell is that the model believes it ALREADY delivered the data ("in
the provided dataset above"). It still has candidates; the session has
drifted into summarising, which a profile reset clears. A genuinely dry
market looks different: an empty array, or drafts that are all duplicates.
"""
from __future__ import annotations

from vendor_intel.pipeline.market_research import discovery


SCREENSHOT = (
    "Ten regional and global chemical distributors specializing in "
    "third-party polyurethane raw materials such as polyols, MDI, TDI, and "
    "foam systems have been identified and detailed in the market research "
    "dataset. These companies operate across Europe, the Middle East, "
    "Africa, and Oceania without manufacturing their own branded "
    "polyurethane products. You can review the complete firmographic list "
    "and company profiles in the provided dataset above."
)


def test_the_observed_prose_answer_is_detected():
    assert discovery.looks_like_prose(SCREENSHOT)


def test_a_real_json_answer_is_not_prose():
    """The page echoes the prompt, whose template contains the field
    names, so detection must key on the ANSWER carrying them."""
    reply = (
        "...prompt echo with its JSON template...\n"
        '[{"company": "Ravago", "verdict": "distributor", '
        '"website": "ravago.com"}]'
    )
    assert not discovery.looks_like_prose(reply)


def test_an_empty_answer_is_not_prose():
    """Empty is the genuinely-dry signal and must keep its own handling."""
    assert not discovery.looks_like_prose("")


def test_a_terse_refusal_is_not_prose():
    assert not discovery.looks_like_prose("No further distributors found.")


def test_prose_narrating_a_dataset_is_caught_even_when_worded_differently():
    assert discovery.looks_like_prose(
        "Twelve distributors have been identified and are listed above "
        "in the dataset for your review."
    )


SCREENSHOT_2 = (
    "The provided market research provides a verified list of ten global "
    "distributors operating within the polyurethane sector across regions "
    "including Turkey, Vietnam, South Africa, Australia, and Colombia. Each "
    "profiled entity exclusively resells third-party chemical components "
    "such as polyols, TDI, and MDI without engaging in manufacturing "
    "operations."
)


def test_the_second_observed_wording_is_detected():
    """A keyword-tuned detector missed this one.

    The first failure said "have been identified and detailed in the
    market research dataset"; this one says "The provided market research
    provides a verified list of ten global distributors". Same failure,
    no shared trigger phrase -- which is why detection is structural
    (prose length + market vocabulary + no JSON) rather than a word list.
    """
    assert discovery.looks_like_prose(SCREENSHOT_2)


def test_a_long_answer_containing_an_array_is_not_prose():
    """Length and vocabulary alone must not convict: a real answer often
    has commentary wrapped around the array."""
    reply = (
        "Here are ten global polyurethane distributors across Turkey, "
        "Vietnam, South Africa, Australia and Colombia, each reselling "
        "third-party polyols, TDI and MDI without manufacturing. "
        '[{"company": "Ravago", "verdict": "distributor"}]'
    )
    assert not discovery.looks_like_prose(reply)
