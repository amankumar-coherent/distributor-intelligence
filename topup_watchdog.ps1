# Top all six markets to 400 VERIFIED companies, unattended.
#
# top_up_to_target.py already loops discover -> verify per market and
# stops a market that stops yielding. This outer loop exists for the case
# where the whole PROCESS dies (a crash, a wedged driver, the machine
# sleeping): everything underneath checkpoints, so re-entering costs
# nothing and resumes from the last completed batch or chunk.
#
# Refusals and rate limits no longer need a human: both discovery and
# verification now rebuild the browser profile when AI Mode declines,
# because the refusal lives in the session cookies and a browser restart
# alone never cleared it. Staged extensions (captcha-raptor) sit outside
# the profile and survive every reset.

$ErrorActionPreference = "Continue"
$py = ".venv\Scripts\python.exe"
$maxAttempts = 10

for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
    Write-Output ""
    Write-Output ("#" * 64)
    Write-Output "# WATCHDOG attempt $attempt/$maxAttempts  ($(Get-Date -Format 'HH:mm:ss'))"
    Write-Output ("#" * 64)

    & $py top_up_to_target.py --target 400

    # Ask the checkpoints, not the exit code: top_up exits 0 having done
    # its best even when a market is genuinely exhausted.
    $short = & $py verified_shortfall.py
    if ($short -eq "0") {
        Write-Output ""
        Write-Output "All six markets reached 400 verified."
        break
    }

    Write-Output ""
    Write-Output "$short market(s) still short of 400 — retrying (progress is saved)"
    Start-Sleep -Seconds 60
}

Write-Output ""
Write-Output ("-" * 62)
Write-Output "FINAL VERIFIED COUNTS"
Write-Output ("-" * 62)
& $py verified_shortfall.py --report
