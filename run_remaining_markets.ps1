# Run every remaining market sequentially.
#
# The Dry Eyes purge already ran (175 -> 57 companies, audit log in
# output\market_research\). This continues from there: discovery tops each
# market back up to target under the new product-scope constraint, then
# verifies.
#
# Everything is resumable -- discovery checkpoints per batch, verification
# per chunk of 10 -- so re-running this after an interruption continues
# rather than restarting.

$ErrorActionPreference = "Continue"
& ".venv\Scripts\python.exe" run_all_markets.py --target 400
Write-Output ""
Write-Output "ALL DONE. Outputs in output\market_research\"
