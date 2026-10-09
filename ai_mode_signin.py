#!/usr/bin/env python3
"""One-time Google sign-in for the AI Mode automation profile.

Google challenges ANONYMOUS AI Mode (udm=50) traffic almost immediately —
measured: plain google.com loaded fine from the automation profile while a
udm=50 URL went straight to /sorry/ "unusual traffic" from the same profile
at the same moment. Signing in once removes that, and the session persists
in the profile directory for every later pipeline run.

Usage:
    .\\.venv\\Scripts\\python.exe ai_mode_signin.py

A Chrome window opens on the Google sign-in page. Sign in, then press
Ctrl+C (or wait — it detects the signed-in state on its own).
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT / "src") not in sys.path:
    sys.path.insert(0, str(ROOT / "src"))


def main() -> int:
    import run_cursor_expand as rce

    rce._load_env()

    from vendor_intel.scraping.google_ai_mode import (
        _profile_dir,
        ai_mode_enabled,
        open_for_signin,
        shutdown,
    )

    if not ai_mode_enabled():
        print("GOOGLE_AI_MODE_ENABLED is not true in .env — nothing to sign in to.")
        return 1

    print(f"automation profile: {_profile_dir()}")
    try:
        ok = open_for_signin()
    finally:
        shutdown()
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
