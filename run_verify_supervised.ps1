param(
    [string]$Only = ""
)

$ErrorActionPreference = "Continue"
Set-Location "C:\Users\aman.kumar\Downloads\vendor-intel"

$maxRestarts = 60
$restart = 0

while ($restart -lt $maxRestarts) {
    $restart++
    Write-Host "=== supervised run attempt $restart ==="

    # Kill any leftover automation Chromium from a previous crashed attempt
    Get-CimInstance Win32_Process -Filter "Name = 'chrome.exe'" |
        Where-Object { $_.CommandLine -like "*ai_mode_chrome_profile*" } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    Start-Sleep -Seconds 2

    $logFile = "verify_supervised_$restart.log"
    if ($Only -ne "") {
        python verify_distributors_ai_mode.py $Only *> $logFile
    } else {
        python verify_distributors_ai_mode.py *> $logFile
    }
    $exitCode = $LASTEXITCODE

    $tail = Get-Content $logFile -Tail 5 -ErrorAction SilentlyContinue
    $doneAll = ($tail -join "`n") -match "ALL MARKETS DONE"

    Write-Host "attempt $restart exit code: $exitCode"
    if ($doneAll) {
        Write-Host "=== all markets verified — stopping supervisor ==="
        break
    }
    Write-Host "not finished yet — restarting in 5s (will resume from checkpoints)"
    Start-Sleep -Seconds 5
}
