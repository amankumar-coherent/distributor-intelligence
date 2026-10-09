# Run these once the four-market queue has finished

Both need the browser profile, so they cannot run while the pipeline is
running — `google_ai_mode` will refuse with `ProfileInUse` rather than
loop, so a mistake is loud rather than silent.

Check first:

    .venv\Scripts\python.exe check_four_markets.py

## 1. Refill the echoed Core Categories cells  (affects >50% of rows)

Measured: **227 of 414** Dry Eyes rows and **33 of 80** Solar Rooftop rows
have `Core Categories` set to the literal market string —

    Core Categories = "Global Dry Eyes Treatment Market - Distributors Only"

— which says only that the company is in the market we asked about. The
column is unusable in those rows, and it hides whether the company sells
the right product.

Already fixed for NEW companies: the prompt no longer interpolates the
market into the field it asks about, and `_clean_category` blanks an echo.
This repairs rows already banked.

    .venv\Scripts\python.exe backfill_categories.py --all --export

Existing good values are never overwritten, and a cell whose replacement
cannot be established is left exactly as it is rather than blanked — so
this can only add information. Backs up state and re-exports.

## 2. Decide on the 13 punctal-plug distributors  (Dry Eyes)

See `DRY_EYES_REVIEW.md`. Punctal occlusion is a real dry-eye treatment,
but it is a **device, not a drug**, and you asked for drug distributors.
They qualify under the scope I wrote, so this is a scope decision, not a
bug — which is why I left them in rather than deleting them.

To remove: drop `punctal plugs` from the qualifying list in
`discovery._PRODUCT_SCOPES`, then

    .venv\Scripts\python.exe purge_off_scope_companies.py ^
      --market "Global Dry Eyes Treatment Market - Distributors Only" --export

## Not pending — already done

  * Dry Eyes: 414 discovered, **373 verified**, 75 countries, banner +
    product scope in both workbooks.
  * The original off-scope purge: 175 -> 57, 118 equipment/optical
    distributors removed, fully audited.
