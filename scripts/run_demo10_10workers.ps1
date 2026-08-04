# Launch 10 market workers (1 market per PowerShell window), staggered to avoid rate-limit storms.
# Output: D:\vendor-intel\output\run5_demo10\<slug>\  (csv / json / xlsx)
#
# Usage:
#   cd D:\vendor-intel
#   .\scripts\start_services.ps1
#   .\scripts\run_demo10_10workers.ps1

$ErrorActionPreference = "Stop"
Set-Location "D:\vendor-intel"

$outRoot = "output\run5_demo10"
New-Item -ItemType Directory -Force -Path $outRoot | Out-Null

$env:MARKET_QUERY_OUTPUT_DIR = $outRoot
$env:CSV_OUTPUT_DIR = $outRoot
$env:GOOGLE_AI_SCRAPER_ENABLED = "true"
$env:GOOGLE_AI_SCRAPER_GAP_FILL = "true"
$env:GOOGLE_AI_SCRAPER_LLM_CLEAN = "true"
$env:DISCOVERY_COUNTRY_SWEEP = "true"
# Soften LinkedIn / search pressure when 10 workers run
$env:LINKEDIN_MCP_DELAY_SEC = "5"
$env:LINKEDIN_MCP_DELAY_JITTER = "2"
$env:GAP_FILL_CONCURRENT = "1"
$env:PIPELINE_ENRICH_CONCURRENT = "2"
$env:PIPELINE_CLASSIFY_CONCURRENT = "2"

$py = "D:\vendor-intel\.venv\Scripts\python.exe"
$markets = @(
  "Global Avocado Oil Market",
  "Global Gel Ice Pack Market",
  "Global High Energy Laser Market",
  "Global Spectral Sensor Market",
  "Global Technical Foam Market",
  "Global Cyber-Physical Systems Security Market",
  "Global Firewall Market",
  "Global Green Chemical Market",
  "Global Smartwatch Market",
  "Global Sustainable Battery Materials Market"
)

# Stagger start (seconds) so workers don't all hit Google/LinkedIn at once
$staggerSec = 45

Write-Host "Output: $outRoot"
Write-Host "Workers: $($markets.Count) (stagger ${staggerSec}s)"
Write-Host ""

for ($i = 0; $i -lt $markets.Count; $i++) {
  $q = $markets[$i]
  $n = $i + 1
  $title = "VI-W$n"
  # Each window: set env, run one market with quality+maximum, no old phase1/2 reuse
  $cmd = @"
`$Host.UI.RawUI.WindowTitle = '$title - $q'
Set-Location 'D:\vendor-intel'
`$env:MARKET_QUERY_OUTPUT_DIR = '$outRoot'
`$env:CSV_OUTPUT_DIR = '$outRoot'
`$env:GOOGLE_AI_SCRAPER_ENABLED = 'true'
`$env:GOOGLE_AI_SCRAPER_GAP_FILL = 'true'
`$env:GOOGLE_AI_SCRAPER_LLM_CLEAN = 'true'
`$env:DISCOVERY_COUNTRY_SWEEP = 'true'
`$env:LINKEDIN_MCP_DELAY_SEC = '5'
`$env:GAP_FILL_CONCURRENT = '1'
Write-Host '=== Worker $n/10 ===' -ForegroundColor Cyan
Write-Host 'Market: $q'
Write-Host 'Out: $outRoot'
& '$py' run_query.py --query `"$q`" --country global --mode quality --cap maximum
Write-Host ''
Write-Host 'Worker $n finished. Press Enter to close.'
Read-Host
"@
  Start-Process powershell.exe -ArgumentList @("-NoExit", "-Command", $cmd)
  Write-Host "[$n/10] started: $q"
  if ($i -lt $markets.Count - 1) {
    Write-Host "  waiting ${staggerSec}s before next worker..."
    Start-Sleep -Seconds $staggerSec
  }
}

Write-Host ""
Write-Host "All 10 workers launched. Watch each window; outputs under $outRoot"
Write-Host "Resume a single market later:"
Write-Host '  $env:MARKET_QUERY_OUTPUT_DIR="output\run5_demo10"'
Write-Host '  .\.venv\Scripts\python.exe run_query.py --resume --query "Global Firewall Market"'
