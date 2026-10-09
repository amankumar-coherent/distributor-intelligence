#!/usr/bin/env python3
"""Verify all 501 AP Automation channel partners: active/closed status,
acquired-by, AND genuine market fit (is it really a channel partner of
this market, or did discovery mis-collect it).

Reuses verify_distributors_ai_mode.py's battle-tested machinery
(chunking, checkpointing, resume-by-judged-company, confidence gate,
status/acquired fields) rather than re-implementing it. The one thing
that machinery gets WRONG for this market and has to be patched around:

_declared_kind(market_name) resolves the verification rule by looking
the market name up in market_batches.ALL_MARKETS (the fire/drone
registries). "Accounts Payable Automation Market" is not in there, so it
falls through to the generic "distributor" default -- which demands a
company physically resell OTHER companies' PRODUCTS, and has no
system-integrator carve-out. Running verification under that rule would
judge the market against criteria it was never collected under: most of
the 501 are ERP/accounting IMPLEMENTATION partners (SAP, NetSuite,
Dynamics, Sage resellers and integrators), which discovery explicitly
steered toward "channel_partner" specifically because the stock
software rule excludes system integrators by name. Verifying with the
wrong rule would likely drop most of the list for the same reason it
had to be worked around at discovery time.

Fixed by monkeypatching _declared_kind AND criteria_prompt for the
duration of this run only (process-local; nothing written to
market_batches.py or verify_distributors_ai_mode.py). _declared_kind
forces "software" so the market is read as a channel-partner market at
all. criteria_prompt is REPLACED (not patched/appended) with wording
that explicitly keeps resellers, VARs, system integrators, ERP
implementation and technology/consulting partners per the user's own
market definition point 4 ("A map of the reseller, system integrator,
and technology partner ecosystem") -- the base prompt's generic tail
("DROP any company that manufactures ANY of what it sells", "DROP too:
... consultancies ...") was silently overriding the software kind's own
"KEEP resellers... and system integrators" line, dropping genuine
implementation/channel partners for owning a side tool or being
consultancy-shaped. See _AP_CRITERIA below for the exact wording.

Usage:
    python verify_ap_automation.py
    python verify_ap_automation.py --status
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

import verify_distributors_ai_mode as V  # noqa: E402
import vendor_intel.scraping.google_ai_mode as gam  # noqa: E402

MARKET_NAME = "Accounts Payable Automation Market"
import os as _os
# Overridable so verify_ap_automation_10_browsers.py can point each
# worker at its OWN shard file -- a 10-way shared checkpoint on the one
# full file would overwrite itself, since save_checkpoint() has no file
# lock or merge (see that script's docstring for the full reasoning).
XLSX_PATH = Path(_os.environ.get("AP_VERIFY_FILE_OVERRIDE") or
                 (ROOT / "output" / "market_research" /
                  "accounts-payable-automation-market.xlsx"))

# Measured: the full prompt with the first draft of this sentence was
# 1,662 chars -- more than double the ~700-char point where this file's
# own criteria_prompt() starts getting empty answers from AI Mode (its
# docstring: "past roughly 700 characters the whole pass starts
# returning nothing at all"; confirmed again at 2,096 -> empty). Existing
# per-market "extra" additions in criteria_prompt() cost ~100 chars, so
# this is held to the same budget rather than the ~500 a fuller
# AP-specific explanation would need. This is the market-FIT half of the
# check: the software rule alone would keep a reseller of ANY platform,
# not specifically an AP automation one.
_AP_SCOPE = "Must resell/implement AP automation or an ERP AP module.\n"

# REPLACES the base criteria_prompt() entirely for this market rather than
# patching it. The base generic tail (shared by every market, including
# "software") carries two clauses that silently override the "KEEP
# resellers, VARs, channel partners and system integrators" line for
# software/hybrid kinds:
#   "DROP any company that manufactures ANY of what it sells"
#   "DROP too: ... consultancies ..."
# Both fire on exactly the partner types the user's own market definition
# (#4: "reseller, system integrator, and technology partner ecosystem")
# requires KEEPING -- an SAP/NetSuite/Dynamics implementation partner that
# also ships a small bolt-on tool, or a consulting-shaped firm that
# implements AP modules for clients, is a real channel partner here, not
# a manufacturer or an out-of-scope consultancy. Verified live: dropped
# verdicts quoted reasons like "disqualifying it under the strict rule
# against manufacturing any own products" and "violating the rule to DROP
# consultancies" for companies whose own stated reason was "resells and
# implements" a named AP/ERP vendor's product -- i.e. the KEEP rule was
# satisfied but a DIFFERENT clause in the same prompt overrode it.
# Appending text cannot fix this: the model sees both instructions and
# the stricter one wins. The conflicting clauses are not emitted at all.
_AP_CRITERIA = (
    f"Decide KEEP or DROP for each company for the {MARKET_NAME}.\n"
    "KEEP any reseller, VAR, solution provider, system integrator, ERP "
    "implementation or technology/consulting partner that "
    "resells/implements AP automation or an ERP AP module - even if also "
    "a consultancy or selling its own small tool too.\n"
    "DROP only: the AP vendor itself, BPO/outsourced AP services, "
    "referral-only affiliates, no AP/ERP-AP connection.\n"
    "When unsure, fits_criteria=false. One result per company, never "
    "skip a name.\n"
)


def _patched_declared_kind(market_name: str) -> str:
    if market_name.strip() == MARKET_NAME:
        return "software"
    return V._real_declared_kind(market_name)


def _patched_criteria_prompt(market_name: str) -> str:
    if market_name.strip() != MARKET_NAME:
        return V._real_criteria_prompt(market_name)
    return _AP_CRITERIA + V._STATUS_RULES + V._JSON_SHAPE


# The exported xlsx's own banner row (row 2) is built by a SEPARATE
# function, export_xlsx._definition_banner(), from the shared generic
# _DEF_SOFTWARE_SERVICES text -- which says "EXCLUDED: ... system
# integrators and consultancies building bespoke solutions...". That is
# the rule for an ordinary software-reseller market, not the rule
# actually used here (_AP_CRITERIA above explicitly KEEPS system
# integrators and consultancy-shaped partners). Registering the kind via
# set_distributor_definition() fixed the banner's OPENING line ("CHANNEL
# PARTNERS ARE THE DISTRIBUTORS") but not this leftover EXCLUDED clause,
# which is wrong for this market specifically and would mislead anyone
# reading row 2 of the file about what was actually kept. Patched the
# same way as criteria_prompt: replace, not append, so the wrong clause
# is never emitted.
_AP_BANNER_RULE = (
    "DISTRIBUTOR / CHANNEL PARTNER DEFINITION (AP automation market): "
    "a company that resells, implements or integrates AP automation "
    "software or an ERP AP module for paying customers - resellers, "
    "VARs, solution providers, system integrators, ERP implementation "
    "and technology/consulting partners all qualify, even if also a "
    "consultancy or selling a small tool of their own. EXCLUDED: the AP "
    "vendor itself selling direct, outsourced AP/BPO invoice processing "
    "services, and referral-only affiliates."
)


def _patched_definition_banner(market_name: str):
    """Replaces export_xlsx._definition_banner for this market only.

    write_output() in verify_distributors_ai_mode.py does a LOCAL import
    of _definition_banner from the export_xlsx module on every call (see
    its own docstring: "The text comes from export_xlsx._definition_banner
    ... so the two files cannot drift apart") -- so the module attribute
    actually read at call time is export_xlsx._definition_banner, not
    anything on verify_distributors_ai_mode. Patched there, not on V.
    """
    import vendor_intel.pipeline.market_research.export_xlsx as _export_xlsx
    if market_name.strip() != MARKET_NAME:
        return _export_xlsx._real_definition_banner(market_name)
    # Reuse the real function's scope/note assembly, only replace the
    # rule line (index 1) with the AP-specific wording.
    rows = _export_xlsx._real_definition_banner(market_name)
    rows[1] = (_AP_BANNER_RULE, False)
    return rows


# Measured against the WORST-CASE chunk (the rows with the longest
# payload once encoded, not the first N), using the REPLACED _AP_CRITERIA
# above (959 chars) rather than the short one-liner this used to patch
# onto the base prompt. The full keep-SI/reseller/consultancy rule costs
# more characters than the one-liner did, so the same CHUNK_SIZE=6 that
# was safe before is NOT safe now: 6 measures at 3,849 chars, past the
# 2,096-char empty-answer cliff. 2 keeps the worst case at 1,938, with
# real margin. Slower (251 chunks instead of ~84) but correct beats fast
# and silently dropping genuine channel partners.
_SAFE_CHUNK_SIZE = 2


def main() -> None:
    import argparse
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--status", action="store_true",
                    help="show checkpoint progress only, no queries")
    args = ap.parse_args()

    ckpt_path = XLSX_PATH.with_name(
        XLSX_PATH.stem + "_verify_checkpoint.json")

    if args.status:
        if not ckpt_path.is_file():
            print("no checkpoint yet — verification has not started")
            return
        import json
        data = json.loads(ckpt_path.read_text(encoding="utf-8"))
        verdicts = data.get("verdicts", [])
        kept = sum(1 for v in verdicts if v.get("keep"))
        closed = sum(1 for v in verdicts
                     if str(v.get("status") or "").lower() == "closed")
        acquired = sum(1 for v in verdicts if v.get("acquired_by"))
        print(f"judged: {len(verdicts)} | kept: {kept} | "
              f"closed-dropped: {closed} | acquired: {acquired}")
        return

    if not gam.ai_mode_enabled():
        print("GOOGLE_AI_MODE_ENABLED is not true — aborting.")
        sys.exit(1)

    if not XLSX_PATH.is_file():
        print(f"NOT FOUND: {XLSX_PATH}")
        sys.exit(1)

    # Monkeypatch for this process only. Keep the originals so the patched
    # functions can delegate for every OTHER market this file still knows
    # about, in case anything else imports this module in-process.
    V._real_declared_kind = V._declared_kind
    V._real_criteria_prompt = V.criteria_prompt
    V._declared_kind = _patched_declared_kind
    V.criteria_prompt = _patched_criteria_prompt
    V.CHUNK_SIZE = _SAFE_CHUNK_SIZE

    # discovery.set_distributor_definition() is an in-memory, per-process
    # registry -- it was populated when ap_automation_market.py ran, but
    # THIS is a separate process, so without re-registering here the
    # exported xlsx's own banner row falls through to the generic
    # "physical goods" wording, which literally says "EXCLUDED: channel
    # partners, resellers, VARs ... system integrators" -- the opposite of
    # the rule actually used to judge every row (verified, live file:
    # row 2 read the wrong banner even though 431/501 were correctly kept
    # as channel partners). The verdict logic above was never wrong; only
    # the sheet's own description of itself was.
    from vendor_intel.pipeline.market_research import discovery as _discovery
    _discovery.set_distributor_definition(MARKET_NAME, "software")

    # That registration fixes the banner's OPENING line but leaves the
    # generic "EXCLUDED: ... system integrators and consultancies..."
    # clause in place (wrong for THIS market specifically) -- so the
    # banner-building function itself is patched too, same replace-not-
    # append technique as criteria_prompt above. See
    # _patched_definition_banner's docstring for why this patches
    # export_xlsx directly rather than V.
    import vendor_intel.pipeline.market_research.export_xlsx as _export_xlsx
    _export_xlsx._real_definition_banner = _export_xlsx._definition_banner
    _export_xlsx._definition_banner = _patched_definition_banner

    entry = V.market_from_args(str(XLSX_PATH), MARKET_NAME)
    print(f"verifying {entry['path']}")
    print(f"market name for criteria: {entry['market_name']!r}")
    print("rule: software (channel_partner + system integrators) "
          "+ AP-specific market-fit scope")
    try:
        V.verify_market(entry)
    except Exception as exc:  # noqa: BLE001
        print(f"CRASHED: {type(exc).__name__}: {exc}")
        print("progress is checkpointed — resumes on next run")
    gam.shutdown()
    print("DONE")


if __name__ == "__main__":
    main()
