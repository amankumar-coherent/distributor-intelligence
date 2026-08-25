#!/usr/bin/env python3
"""Run ChatGPT expand for 5 markets, one by one, using Chrome Profile 2 scraper (:15561).

Other project keeps :15551. This run uses:
  GOOGLE_AI_SCRAPER_BASE_URL=http://127.0.0.1:15561
  DeepSeek cleans Google AI Overview to useful sentences only
  Fill ALL table rows (target/max-fill 500)

From D:\\vendor-intel:

  .\\.venv\\Scripts\\python.exe run_five_chatgpt_markets.py
"""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
MARKETS = [
    "Global Flexible Packaging Market",
    "Global Wearable Medical Devices Market",
    "Global Semiconductor Market",
    "Global Liquefied Natural Gas Market",
    "Global GLP-1 Receptor Agonist Market",
]


def main() -> int:
    os.chdir(ROOT)
    py = ROOT / ".venv" / "Scripts" / "python.exe"
    if not py.exists():
        py = Path(sys.executable)

    env = os.environ.copy()
    env["PYTHONPATH"] = str(ROOT / "src")
    env["GOOGLE_AI_SCRAPER_ENABLED"] = "true"
    env["GOOGLE_AI_SCRAPER_BASE_URL"] = "http://127.0.0.1:15561"
    env["GOOGLE_AI_SCRAPER_URL"] = "http://127.0.0.1:15561"
    env["GOOGLE_AI_SCRAPER_LLM_CLEAN"] = "true"
    env["GOOGLE_AI_SCRAPER_LLM_CLEAN_ONLY_WHEN_MISSING"] = "false"
    env["GOOGLE_AI_SCRAPER_DISCOVERY"] = "true"
    env["GOOGLE_AI_SCRAPER_GAP_FILL"] = "true"

    health_py = (
        "import sys; sys.path.insert(0, 'src'); import httpx\n"
        "from vendor_intel.clients.google_ai_scraper import google_ai_base_url\n"
        "url = google_ai_base_url().rstrip('/') + '/health'\n"
        "try:\n"
        "    r = httpx.get(url, timeout=5.0)\n"
        "    data = r.json() if r.status_code == 200 else {}\n"
        "except Exception as e:\n"
        "    print('HEALTH FAIL', e); raise SystemExit(1)\n"
        "print(url, data)\n"
        "if not data.get('extension_connected'):\n"
        "    print('Extension not connected on 15561. Set Server URL to http://localhost:15561 and refresh.')\n"
        "    raise SystemExit(1)\n"
    )
    health_file = ROOT / "_tmp_scraper_health.py"
    health_file.write_text(health_py, encoding="utf-8")
    try:
        chk = subprocess.run([str(py), str(health_file)], cwd=str(ROOT), env=env)
        if chk.returncode != 0:
            print(
                "Start scraper on 15561 and point THIS Chrome extension to http://localhost:15561",
                flush=True,
            )
            return 1
    finally:
        try:
            health_file.unlink(missing_ok=True)
        except Exception:
            pass

    failed: list[str] = []
    for i, market in enumerate(MARKETS, 1):
        print(f"\n=== [{i}/{len(MARKETS)}] {market} ===\n", flush=True)
        cmd = [
            str(py),
            str(ROOT / "run_cursor_expand.py"),
            "--backend",
            "chatgpt",
            "--query",
            market,
            "--country",
            "global",
            "--batch",
            "all",
            "--target",
            "500",
            "--max-fill",
            "500",
            "--max-queries",
            "180",
            "--fill-existing",
        ]
        rc = subprocess.run(cmd, cwd=str(ROOT), env=env).returncode
        if rc != 0:
            print(f"FAILED: {market} (exit {rc})", flush=True)
            failed.append(market)
        else:
            print(f"DONE: {market}", flush=True)

    print(f"\nBatch complete: {len(MARKETS) - len(failed)}/{len(MARKETS)} ok", flush=True)
    if failed:
        print("Failed: " + "; ".join(failed), flush=True)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
