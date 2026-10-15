# Resume the ProFood India batch (75 markets) from where it stopped.
#   powershell -ExecutionPolicy Bypass -File launch\resume_food.ps1
# Starts the batch launcher and its watchdog in the background with no
# console window (conhost --headless: on Windows 11 a normal start opens a
# Windows Terminal window, and closing that window kills the run).
# Progress: python verify_all_drones.py --status   |  logs\food_progress.log
$root = Split-Path -Parent $PSScriptRoot
foreach ($f in "run_food_batch.ps1", "run_food_watchdog.ps1") {
    $ps1 = Join-Path $PSScriptRoot $f
    $r = Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{
        CommandLine      = "conhost.exe --headless powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$ps1`""
        CurrentDirectory = $root
    }
    "started $f (pid $($r.ProcessId), rc $($r.ReturnValue))"
}
