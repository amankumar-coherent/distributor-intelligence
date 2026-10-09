# Run the four remaining markets to a target of 400 each, unattended.
#
#   Dry Eyes Treatment / Solar Rooftop / Polyurethane / CPaaS
#
# Why a watchdog loop rather than a single invocation: these runs take
# hours and the failure modes are transient -- a CAPTCHA streak, a wedged
# browser, a dropped connection. Every layer below is already resumable
# (discovery checkpoints per batch, verification per chunk of 10), so
# re-entering run_all_markets.py after a crash continues from the last
# checkpoint instead of restarting. The loop just keeps re-entering until
# every market reports its target or the attempt budget runs out.
#
# SINGLE INSTANCE ONLY -- see the guard below. Two copies of this script
# ran at once (2:34:39, twice) and the result was two run_all_markets
# trees, two discovery processes and two verifiers all driving ONE browser
# profile. That surfaces as an endless
#     "BrowserType.launch_persistent_context: Opening in existing browser
#      session ... the profile is already in use"
# loop which burns the whole session-reset budget, because no amount of
# relaunching can win a lock another process legitimately holds. The
# markets share one profile and one IP by design; concurrency is never
# valid here.
#
# NOTE: progress is checked by check_four_markets.py, a real file. An
# earlier version embedded that Python in a PowerShell @"..."@ here-string,
# which PowerShell parsed as its own code ("The 'from' keyword is not
# supported in this version of the language") and the watchdog never ran.

$ErrorActionPreference = "Continue"
$py = ".venv\Scripts\python.exe"
$maxAttempts = 12

# --- single-instance guard ------------------------------------------------
$lockPath = Join-Path $PSScriptRoot "run_four_markets.lock"
$mine = $PID

if (Test-Path $lockPath) {
    $existing = (Get-Content $lockPath -ErrorAction SilentlyContinue | Select-Object -First 1)
    $alive = $null
    if ($existing) { $alive = Get-Process -Id $existing -ErrorAction SilentlyContinue }
    if ($alive) {
        Write-Output "ALREADY RUNNING as PID $existing - refusing to start a second copy."
        Write-Output "Two copies fight over one browser profile and neither makes progress."
        Write-Output "To take over: Stop-Process -Id $existing ; Remove-Item '$lockPath'"
        exit 1
    }
    Write-Output "stale lock from PID $existing (no longer running) - taking it over"
    Remove-Item $lockPath -Force -ErrorAction SilentlyContinue
}
Set-Content -Path $lockPath -Value $mine -Encoding utf8

# Also clear anything left over from a previous crashed run, so the first
# launch is not fighting an orphaned Chrome for the profile lock.
Get-CimInstance Win32_Process |
    Where-Object { $_.CommandLine -like '*ai_mode_chrome_profile*' } |
    ForEach-Object {
        Write-Output "killing orphaned automation Chrome PID $($_.ProcessId)"
        try { Stop-Process -Id $_.ProcessId -Force -ErrorAction Stop } catch {}
    }

try {
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        Write-Output ""
        Write-Output "################################################################"
        Write-Output "# WATCHDOG attempt $attempt/$maxAttempts  ($(Get-Date -Format 'HH:mm:ss'))"
        Write-Output "################################################################"

        & $py run_all_markets.py --target 400 --only "dry eyes" --only "solar rooftop" --only "polyurethane" --only "cpaas"
        $code = $LASTEXITCODE

        # Ask the pipeline's own state how far each market got. The exit
        # code alone is not enough: run_all_markets exits 0 even when a
        # market stopped early, because a genuinely exhausted market is
        # not an error.
        Write-Output ""
        Write-Output "--- progress after attempt $attempt ---"
        $short = & $py check_four_markets.py

        if ($short -eq "0") {
            Write-Output ""
            Write-Output "All four markets reached 400. Done."
            break
        }

        Write-Output ""
        Write-Output "$short market(s) still short of 400 (exit code $code) - retrying."
        Write-Output "Progress is checkpointed; the retry continues rather than restarting."
        Start-Sleep -Seconds 60
    }

    Write-Output ""
    Write-Output "=============================================================="
    Write-Output "FINAL STATE"
    Write-Output "=============================================================="
    & $py check_four_markets.py | Out-Null
    Write-Output ""
    Write-Output "Outputs in output\market_research\"
}
finally {
    Remove-Item $lockPath -Force -ErrorAction SilentlyContinue
}
