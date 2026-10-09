# Re-verify all six markets under the third-party-distributor-only rule.
#
# What changed: a company that MANUFACTURES any of what it sells is no
# longer a distributor here, even when it also runs a distribution arm and
# carries other brands. The old criteria said "pure manufacturers", so a
# hybrid passed; that is the company being removed.
#
# Runs in Chromium so the captcha-raptor extension loads (enterprise-managed
# Chrome refuses command-line unpacked extensions). GOOGLE_AI_MODE_BROWSER
# and GOOGLE_AI_MODE_EXTENSIONS in .env already select it.
#
# RATE LIMITS. Google answers with "You've reached the request limit for AI
# responses" — a normal page, not an error. The verifier now recognises it,
# rebuilds the PROFILE (the limit lives in the session cookies, so a browser
# restart alone never cleared it) and retries the same chunk. A chunk that
# fails for backend reasons is NOT checkpointed, so the offset rewinds and
# a re-run re-asks those companies rather than recording them as rejected.
#
# The loop below keeps re-entering each market until every company has a
# real verdict, which is what clears a backlog left by an earlier failure.

$ErrorActionPreference = "Continue"
$py = ".venv\Scripts\python.exe"
$maxAttempts = 6

$markets = @(
    @{ slug = "global-silicon-carbide-market-distributors-only";                           name = "Global Silicon Carbide Market" },
    @{ slug = "global-wearable-glucometer-market-distributors-only";                       name = "Global Wearable Glucometer Market" },
    @{ slug = "global-dry-eyes-treatment-market-distributors-only";                        name = "Global Dry Eyes Treatment Market" },
    @{ slug = "global-solar-rooftop-market-distributors-only";                             name = "Global Solar Rooftop Market" },
    @{ slug = "global-polyurethane-market-distributors-only";                              name = "Global Polyurethane Market" },
    @{ slug = "global-cpaas-communication-platform-as-a-service-market-distributors-only"; name = "Global CPaaS (Communication Platform as a Service) Market" }
)

foreach ($m in $markets) {
    $file = "output\market_research\$($m.slug).xlsx"
    if (-not (Test-Path $file)) {
        Write-Output "SKIP (no discovery export): $($m.name)"
        continue
    }

    Write-Output ""
    Write-Output "=============================================================="
    Write-Output "$($m.name)   $(Get-Date -Format 'HH:mm:ss')"
    Write-Output "=============================================================="

    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        & $py verify_distributors_ai_mode.py --file $file --market-name $m.name

        # Exit code alone is not enough: the verifier exits 0 having
        # written a file even when a chunk failed for backend reasons, so
        # ask the checkpoint whether every company actually has a verdict.
        $left = & $py unverified_count.py $m.slug
        if ($left -eq "0") {
            Write-Output "  all companies verified"
            break
        }

        Write-Output "  $left companies still unverified after attempt $attempt/$maxAttempts"
        Write-Output "  (progress is checkpointed; the retry re-asks only those chunks)"
        Start-Sleep -Seconds 60
    }
}

Write-Output ""
Write-Output "=============================================================="
Write-Output "FINAL STATE"
Write-Output "=============================================================="
foreach ($m in $markets) {
    $left = & $py unverified_count.py $m.slug
    Write-Output ("  {0,-56} {1} unverified" -f $m.name, $left)
}
