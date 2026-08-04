# Run 10 global markets into a FRESH folder (ignores old output/* runs).
# Checkpoints: batch DB skips done markets; post_market_progress.json resumes mid-enrich.
#
# First start (fresh discovery — do not reuse old Phase 1/2):
#   .\scripts\run_demo10_fresh.ps1
# Resume after Ctrl+C (same batch-id):
#   .\scripts\run_demo10_fresh.ps1 -Resume

param(
    [switch]$Resume
)

$ErrorActionPreference = "Stop"
Set-Location "D:\vendor-intel"

$env:CSV_OUTPUT_DIR = "output\run5_demo10"
$env:GOOGLE_AI_SCRAPER_ENABLED = "true"
$env:GOOGLE_AI_SCRAPER_GAP_FILL = "true"
$env:GOOGLE_AI_SCRAPER_LLM_CLEAN = "true"
$env:DISCOVERY_COUNTRY_SWEEP = "true"

New-Item -ItemType Directory -Force -Path $env:CSV_OUTPUT_DIR | Out-Null
Copy-Item -Force "queries\demo10_global_markets.txt" "queries\demo10_global_markets.txt" -ErrorAction SilentlyContinue

$batchId = "demo10-gai-jul30"
$py = ".\.venv\Scripts\python.exe"

Write-Host "Output root: $env:CSV_OUTPUT_DIR"
Write-Host "Batch id:    $batchId"
Write-Host "Markets:     queries\demo10_global_markets.txt"
Write-Host ""

# Optional helpers (start if you use them)
# docker compose up -d
# Start-Process -FilePath ".\.venv\Scripts\linkedin-scraper-mcp.exe" -ArgumentList "--transport","streamable-http","--port","8001" -WindowStyle Minimized

if ($Resume) {
    Write-Host "RESUME — skip markets already done; continue post_market checkpoints"
    & $py run_batch.py `
        --batch-id $batchId `
        --mode quality `
        --cap maximum `
        --max-concurrent 1 `
        --yes
} else {
    Write-Host "FRESH START — new folder (MARKET_QUERY_OUTPUT_DIR / CSV_OUTPUT_DIR)"
    & $py run_batch.py `
        --input queries\demo10_global_markets.txt `
        --batch-id $batchId `
        --mode quality `
        --cap maximum `
        --max-concurrent 1 `
        --no-resume-checkpoint `
        --yes
}

Write-Host ""
Write-Host "Status:  $py status_batch.py --batch-id $batchId"
Write-Host "Outputs: $env:CSV_OUTPUT_DIR\<market_slug>\  (xlsx / csv / json)"
