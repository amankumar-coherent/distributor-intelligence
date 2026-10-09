# Wait for Google's AI Mode quota to reset, then finish the office-contact fill.
#
# The quota is per-account and resets on Google's own schedule, so the only
# thing to do is probe cheaply and start the real run the moment a probe
# gets a genuine answer. Probing with a three-word prompt costs one
# request; starting the full run against an exhausted quota costs a
# browser launch and a CAPTCHA solve to learn the same thing.
#
# Safe to leave running: fill_office_contacts.py checkpoints every chunk
# and --only-empty skips rows already filled, so if the quota runs out
# again mid-run this script simply waits and resumes.

$ErrorActionPreference = "Continue"
$py = ".venv\Scripts\python.exe"
$maxHours = 12
$probeEvery = 900          # 15 minutes

$markets = @(
    "--market", "Global CPaaS (Communication Platform as a Service) Market - Distributors Only",
    "--market", "Global Solar Rooftop Market - Distributors Only"
)

$deadline = (Get-Date).AddHours($maxHours)
$round = 0

while ((Get-Date) -lt $deadline) {
    $round++
    Write-Output "[$(Get-Date -Format 'HH:mm:ss')] probe $round ..."

    $state = & $py probe_ai_mode_quota.py
    if ($state -eq "AVAILABLE") {
        Write-Output "[$(Get-Date -Format 'HH:mm:ss')] quota available - starting the fill"
        & $py fill_office_contacts.py @markets --export --only-empty
        $code = $LASTEXITCODE

        # Exit 0 means the markets finished. Any other exit (including the
        # quota stop) means there is more to do, so go back to waiting.
        if ($code -eq 0) {
            Write-Output "[$(Get-Date -Format 'HH:mm:ss')] fill completed"
            break
        }
        Write-Output "[$(Get-Date -Format 'HH:mm:ss')] stopped early (exit $code) - waiting for more quota"
    }
    else {
        Write-Output "[$(Get-Date -Format 'HH:mm:ss')] still rate limited"
    }

    Start-Sleep -Seconds $probeEvery
}

Write-Output ""
Write-Output "FINAL COVERAGE"
& $py office_coverage.py
