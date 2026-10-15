# ProFood India - 75 markets, India companies only, distributors (rule per
# market in food_markets.py), cap 250 verified each, 15 browsers at a time; a
# finished market's slot starts the next. Re-running RESUMES: markets already at
# 250 are skipped, found/verified companies and searched regions are kept
# (data/market_research, data/region_coverage, output/market_research).
# Portable: works from any clone location. Start it with launch\resume_food.ps1.
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root
New-Item -ItemType Directory -Force -Path (Join-Path $root "logs") | Out-Null
$env:PYTHONIOENCODING = "utf-8"
# Each browser may rebuild its own profile (fresh cookies) this many times
# per run when CAPTCHA / "Verifying your request" pages persist.
$env:GOOGLE_AI_MODE_MAX_RESETS = "10"
& .venv\Scripts\python.exe -u run_fire_10_browsers.py --batch food --rolling --wave 15 --target 250 `
    2>&1 | Out-File -Encoding utf8 -Append "logs\run_food_batch.log"
