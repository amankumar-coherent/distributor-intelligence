$logFile = "C:\Users\aman.kumar\Downloads\vendor-intel\verify_run13.log"

while ($true) {
    $found = Select-String -Path $logFile -Pattern "ALL MARKETS DONE" -Quiet -ErrorAction SilentlyContinue
    if ($found) {
        Write-Output "ALL MARKETS DONE detected"
        break
    }

    $procs = Get-CimInstance Win32_Process -Filter "Name='python.exe'"
    $match = $procs | Where-Object { $_.CommandLine -like "*verify_distributors_ai_mode*" }
    if (-not $match) {
        Write-Output "PROCESS DIED"
        break
    }

    Start-Sleep -Seconds 90
    Write-Output "still running..."
}
