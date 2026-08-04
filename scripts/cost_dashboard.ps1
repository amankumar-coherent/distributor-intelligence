# Open OpenAI cost dashboard (Streamlit)
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root
$py = Join-Path $Root ".venv\Scripts\python.exe"
& $py -m streamlit run cost_dashboard.py
exit $LASTEXITCODE
