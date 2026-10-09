# Dry Eyes: 19 companies needing your decision

Everything below is IN the sheet right now. Nothing was deleted on my
judgement — these need a scope call that is yours to make.

## A. 13 punctal-plug distributors — devices, not drugs

Punctal occlusion is a standard dry-eye treatment, so these are genuine
dry-eye companies. But they distribute a **device**, not a pharmaceutical.
My product scope explicitly listed "punctal plugs" as qualifying, so AI
Mode kept them correctly — the question is whether that was the right
scope, given you asked for companies that "sell drug related to dry eyes
treatment market".

    Amarak Medical Equipment          Punctal Plugs
    Amarak Medicals                   Lacrivera VeraPlug Punctal Plugs
    Eye Nation Medical Sdn Bhd        punctal plugs
    Grafton Optical Company Limited   Form Fit Hydrogel Punctum Plugs
    IMEX Clinic S.L.                  I-PLUG punctum plugs
    Inoftal S.A.S.                    OCuSOFT Dry Eye Supplies
    Medical Care for General Trading  Punctum Plugs
    SOS Surgical & Ophthalmic Supplies Katena punctal plugs
    Surgical & Ophthalmic Supplies    EagleVision punctum plugs
    Time Healthcare Solutions         FCI Lacrimal Instruments
    Toomac Ophthalmic                 Parasol Pre-Loaded Punctal Occluders
    United Medical Inc.               Medennium Smart Plug Punctal Plugs
    Valley Medical Company            Punctal Plugs

To remove them, drop "punctal plugs" from the qualifying list in
`discovery._PRODUCT_SCOPES` and re-run:

    .venv\Scripts\python.exe purge_off_scope_companies.py ^
      --market "Global Dry Eyes Treatment Market - Distributors Only" --export

## B. 6 with no dry-eye product in the sheet's own columns

Three are purge survivors whose justifying product is recorded in the
audit log but not in their Core Categories / Specialty Focus / Key
Products cells — so the sheet does not show why they qualify:

    Cardinal Health, Inc.   audit says: Prescription dry eye remedies
    Genop                   audit says: Xailin Hydrate Eye Drops
    Medicals International  audit says: Lubricating solutions

Three arrived AFTER the purge and were never scope-checked:

    Al Hashar Pharmacy
    Insight Surgical Pty Ltd
    Marzam

Re-running the purge (command above) covers group B either way, since it
re-adjudicates every company currently in state.

## What is already good

  * 344 companies from **75 distinct countries** — India 32, South Africa
    24, Colombia 18, UAE 18, Philippines 17, Kenya 13. US is 8th with 11,
    so the US/Europe concentration you reported is gone.
  * 325 of 344 name a real dry-eye drug (Hylo-Comod, Cationorm, Hialid,
    Aquafilm, Refresh, Tears Naturale, Artelac Lipids, RETANE, Restasis,
    Xiidra, Miebo).
  * The original purge removed 118 of 175 wrong companies (equipment,
    lasers, slit lamps, optical retail) and is fully auditable in
    `output/market_research/..._scope_purge_audit.json`.
