# Duplicate companies: 13-20% of rows are the same company repeated

Found by auditing the finished workbooks. The pipeline dedupes on a
normalised company NAME, so rows whose names genuinely differ get
through even when they are obviously one company:

```
  401 rows -> 325 unique (76 redundant across 53 domains)
    powernsun.com
        PowernSun
        Powernsun (Power n Sun Solar Systems)
        Power & Sun
        Power n Sun
        Power & Sun FZCO
        Powernsun Group
        Power N Sun Solar Equipment's Trading LLC
    segensolar.co.za
        SegenSolar (Pty) Ltd
        Segen South Africa
        Segen Solar (Pty) Ltd
        Segen Solar
        Segen South Africa (Pty) Ltd
    solarandsolar.hu
        Solar & Solar Kft.
        Solar&Solar Wholesale
        Solar&Solar Ltd
```

## Impact

| Market | Rows | Unique companies | Redundant |
|---|---|---|---|
| Dry Eyes | 414 | 363 | 51 (12%) |
| Solar Rooftop | 401 | 324 | 77 (19%) |
| Polyurethane | 193 | 154 | 39 (20%) |

## Why it is not applied automatically

It deletes rows from finished data, and a shared domain is not always
one company. A multinational runs one domain across genuinely distinct
national distributors, so the rule used is **same domain AND same
country**. Verified against the real data:

```
MERGED   (one German company, three ways):
   Krannich Solar Group / Krannich Solar / Krannich Solar GmbH & Co. KG

KEPT     (five real national distributors):
   Krannich Solar GmbH & Co. KG      Germany
   Krannich Solar Turkey             Turkey
   Krannich Solar d.o.o.             Croatia
   Krannich Solar Energy (Pty) Ltd   South Africa
   Krannich Solar Private Limited    India
```

Within a merged group the most complete row wins and any field it
leaves blank is filled from its siblings, so merging never loses a
value. State is backed up first.

## To apply

```powershell
.venv\Scripts\python.exe dedupe_by_website.py --all --dry-run     # review
.venv\Scripts\python.exe dedupe_by_website.py --all --export      # apply
```

Re-run verification afterwards if you want the verified workbooks
regenerated from the deduped list.

## Already fixed for future collection

The pipeline now dedupes on **website + country** at collection time, not
only on the company name, so CPaaS and everything after it cannot
accumulate these. Confirmed live on Polyurethane: 24 duplicates rejected
at source within the first few batches, each logged as
"skip (same website as an existing company)".

Country aliases are normalised too. The live data held both "UAE" (12
rows) and "United Arab Emirates" (65), "USA" (5) and "United States"
(26), "UK" (2) and "United Kingdom" (13) -- a duplicate written with the
other spelling would otherwise have escaped the check entirely.

The counts above are what remains from BEFORE those fixes.
