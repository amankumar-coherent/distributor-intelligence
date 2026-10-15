# Every 10 min: progress to logs\food_progress.log, relaunch the food batch if
# its launcher died, reset browsers whose log is silent 25+ min, and run the
# final status + ownership check (Excel + JSON) when all markets are done.
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root
New-Item -ItemType Directory -Force -Path (Join-Path $root "logs") | Out-Null
$env:PYTHONIOENCODING = "utf-8"
& .venv\Scripts\python.exe -u scripts\batch_watchdog.py --batch food --ps1 launch\run_food_batch.ps1 --target 250 `
    2>&1 | Out-File -Encoding utf8 -Append "logs\run_food_watchdog.log"
