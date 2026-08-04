# Start / verify vendor-intel services (Windows PowerShell)
# Usage: cd D:\vendor-intel; .\scripts\start_services.ps1

$ErrorActionPreference = "Continue"
Set-Location "D:\vendor-intel"

function Test-HttpOk([string]$Url, [int]$TimeoutSec = 4) {
  try {
    $r = Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec $TimeoutSec
    return $true
  } catch {
    # 406 on MCP /mcp is still "alive"
    $code = 0
    try { $code = [int]$_.Exception.Response.StatusCode } catch {}
    if ($code -in 401, 403, 405, 406) { return $true }
    return $false
  }
}

Write-Host "`n=== 1) SearXNG (port 8080) ===" -ForegroundColor Cyan
$searxUp = Test-HttpOk "http://127.0.0.1:8080/"
if (-not $searxUp) {
  Write-Host "Starting SearXNG via docker compose..."
  docker compose up -d searxng
  Start-Sleep -Seconds 5
  $searxUp = Test-HttpOk "http://127.0.0.1:8080/"
}
if ($searxUp) {
  try {
    $sx = Invoke-RestMethod "http://127.0.0.1:8080/search?q=ping&format=json" -TimeoutSec 10
    Write-Host "OK — SearXNG up, sample results=$(@($sx.results).Count)" -ForegroundColor Green
  } catch {
    Write-Host "WARN — SearXNG HTTP up but search failed: $($_.Exception.Message)" -ForegroundColor Yellow
  }
} else {
  Write-Host "FAIL — SearXNG not reachable. Run: docker compose up -d" -ForegroundColor Red
}

Write-Host "`n=== 2) LinkedIn MCP (port 8001) ===" -ForegroundColor Cyan
$liListen = Get-NetTCPConnection -LocalPort 8001 -State Listen -ErrorAction SilentlyContinue
if (-not $liListen) {
  Write-Host "Starting linkedin-scraper-mcp..."
  $exe = "D:\vendor-intel\.venv\Scripts\linkedin-scraper-mcp.exe"
  if (-not (Test-Path $exe)) {
    Write-Host "FAIL — missing $exe" -ForegroundColor Red
  } else {
    Start-Process -FilePath $exe -ArgumentList "--transport","streamable-http","--port","8001" -WindowStyle Minimized
    Start-Sleep -Seconds 4
  }
}
# 406 = server alive (streamable-http needs special Accept header)
$liAlive = Test-HttpOk "http://127.0.0.1:8001/mcp"
if ($liAlive -or (Get-NetTCPConnection -LocalPort 8001 -State Listen -EA SilentlyContinue)) {
  Write-Host "OK — LinkedIn MCP listening on :8001" -ForegroundColor Green
} else {
  Write-Host "FAIL — LinkedIn MCP not up. Manual:" -ForegroundColor Red
  Write-Host '  .\.venv\Scripts\linkedin-scraper-mcp.exe --transport streamable-http --port 8001'
}

Write-Host "`n=== 3) Google AI scraper (port 15551) ===" -ForegroundColor Cyan
$gai = Test-HttpOk "http://127.0.0.1:15551/health"
if ($gai) {
  try {
    $h = Invoke-RestMethod "http://127.0.0.1:15551/health" -TimeoutSec 4
    Write-Host "OK — Google AI scraper (extension=$($h.extension_status))" -ForegroundColor Green
  } catch {
    Write-Host "OK — :15551 responds" -ForegroundColor Green
  }
} else {
  Write-Host "WARN — Google AI scraper not on :15551. Start extension backend if you need it:" -ForegroundColor Yellow
  Write-Host '  cd C:\Users\aman.kumar\Downloads\google-ai-scraper-main\google-ai-scraper-main\server'
  Write-Host '  uv run google-ai-scraper --backend --port 15551'
}

Write-Host "`nDone. Next: .\scripts\run_demo10_10workers.ps1`n"
