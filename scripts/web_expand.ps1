# Web-search expand helper — you only run this; Python does live web search.
#
# Usage (from D:\vendor-intel):
#   .\scripts\web_expand.ps1 -Query "Global Firewall Market" -Target 320 -Batch a
#   .\scripts\web_expand.ps1 -Query "Global Smartwatch Market" -Batch b
#   .\scripts\web_expand.ps1 -Query "Global Green Chemical Market" -Mode pipeline

param(
    [Parameter(Mandatory = $true)]
    [string]$Query,

    [string]$Country = "global",
    [int]$Target = 320,
    [ValidateSet("a", "b", "all")]
    [string]$Batch = "a",
    [int]$MaxQueries = 80,
    [ValidateSet("harvest", "pipeline")]
    [string]$Mode = "harvest",
    [string]$Final = "",
    [switch]$NoMerge
)

$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

$py = Join-Path $Root ".venv\Scripts\python.exe"
if (-not (Test-Path $py)) {
    Write-Error "Missing venv at $py — create it first."
    exit 1
}

$argsList = @(
    "run_web_expand.py",
    "--query", $Query,
    "--country", $Country,
    "--target", "$Target",
    "--batch", $Batch,
    "--max-queries", "$MaxQueries",
    "--mode", $Mode
)
if ($Final) { $argsList += @("--final", $Final) }
if ($NoMerge) { $argsList += "--no-merge" }

Write-Host "Running: $py $($argsList -join ' ')" -ForegroundColor Cyan
& $py @argsList
exit $LASTEXITCODE
