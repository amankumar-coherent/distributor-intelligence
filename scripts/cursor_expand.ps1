# Market expand — default backend is ChatGPT (OPENAI_API_KEY / OpenAI SDK).
#
# Usage (from D:\vendor-intel):
#   .\scripts\cursor_expand.ps1 -Query "Global Firewall Market" -Batch a
#   .\scripts\cursor_expand.ps1 -Query "Global Firewall Market" -Batch a -Enrich
#   .\scripts\cursor_expand.ps1 -Query "Global Firewall Market" -Backend cursor -Batch a

param(
    [Parameter(Mandatory = $true)]
    [string]$Query,

    [string]$Country = "global",
    [ValidateSet("a", "b", "all")]
    [string]$Batch = "a",
    [int]$Target = 320,
    [ValidateSet("chatgpt", "openai", "cursor", "harvest")]
    [string]$Backend = "chatgpt",
    [string]$Final = "",
    [string]$Model = "",
    [int]$MaxFill = 80,
    [switch]$Enrich,
    [switch]$FillExisting,
    [switch]$NoStream
)

$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

$py = Join-Path $Root ".venv\Scripts\python.exe"
if (-not (Test-Path $py)) {
    Write-Error "Missing venv at $py"
    exit 1
}

$argsList = @(
    "run_cursor_expand.py",
    "--backend", $Backend,
    "--query", $Query,
    "--country", $Country,
    "--batch", $Batch,
    "--target", "$Target"
)
if ($Final) { $argsList += @("--final", $Final) }
if ($Model) { $argsList += @("--model", $Model) }
if ($MaxFill -ne 80) { $argsList += @("--max-fill", "$MaxFill") }
if ($Enrich) { $argsList += "--enrich" }
if ($FillExisting) { $argsList += "--fill-existing" }
if ($NoStream) { $argsList += "--no-stream" }

Write-Host "Running expand (backend=$Backend)…" -ForegroundColor Cyan
& $py @argsList
exit $LASTEXITCODE
