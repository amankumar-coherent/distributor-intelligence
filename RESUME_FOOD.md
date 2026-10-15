# Resume the ProFood India batch on another PC

75 markets (India companies only, distributors, cap 250 each, 15 browsers at a
time). Paused 2026-10-12 with 16 markets started, 3,434 found, 2,449 verified.
Its progress is in the repo, so a fresh clone continues where it stopped:

| Saved progress | Path |
|---|---|
| Companies found per market | `data/market_research/<market>.json` |
| Indian states/cities already searched | `data/region_coverage/food_profile_<slug>.json` |
| Exports, verification checkpoints, verified lists | `output/market_research/<market>*` |

## One-time setup on the new PC

1. Clone, then create the virtual environment in the repo folder:
   `python -m venv .venv` and `.venv\Scripts\pip install -r requirements.txt`
   (plus `.venv\Scripts\patchright install chromium`).
2. Copy `.env.example` to `.env` and fill in your keys (the real `.env` is never
   committed).
3. CAPTCHA solver: put the captcha-raptor extension on the PC and point `.env` at it:
   `GOOGLE_AI_MODE_EXTENSIONS=C:\path\to\captcha-raptor\extension`

## Resume

```
powershell -ExecutionPolicy Bypass -File launch\resume_food.ps1
```

This starts the batch and its watchdog in the background (no console window).
Markets already at 250 are skipped, already-searched regions are not searched
again, and companies already verified are not re-checked. When every market is
done the watchdog runs the final active-status + ownership check and writes the
polished Excel + JSON to `output/final/food/`.

Progress: `.venv\Scripts\python.exe verify_all_drones.py --status`, or
`logs\food_progress.log` (updated every 10 minutes).

To save progress again before moving PCs:
`.venv\Scripts\python.exe scripts\save_batch_progress.py --batch food --git-add`
then commit and push.
