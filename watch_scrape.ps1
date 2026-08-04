# Watch avocado scrape logs in a second PowerShell window (does not stop the run).
# Usage:
#   .\watch_scrape.ps1
#   .\watch_scrape.ps1 -ProgressOnly

param(
    [switch]$ProgressOnly,
    [int]$Tail = 30
)

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$liveLog = Join-Path $root "output\avocado_oil\scrape_live.log"
$progress = Join-Path $root "output\avocado_oil\scrape_progress.jsonl"

if ($ProgressOnly) {
    if (-not (Test-Path $progress)) {
        Write-Host "Progress file not found yet: $progress"
        Write-Host "It is created when a run starts with the updated run_avocado_scrape.py"
        exit 1
    }
    Write-Host "Watching progress (one line per company): $progress"
    Get-Content $progress -Wait -Tail $Tail
} else {
    if (-not (Test-Path $liveLog)) {
        Write-Host "Live log not found yet: $liveLog"
        Write-Host "Current run may still write only to the terminal."
        Write-Host "For the running job, try: Get-Content `$env:USERPROFILE\.cursor\projects\...\terminals\10.txt -Wait -Tail 40"
        exit 1
    }
    Write-Host "Watching full log: $liveLog"
    Get-Content $liveLog -Wait -Tail $Tail
}
