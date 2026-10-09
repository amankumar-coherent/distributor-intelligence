#!/usr/bin/env python3
"""Create ten independent Chromium profiles, each with captcha-raptor.

Ten markets run in parallel, and a Chromium profile cannot be shared:
launch_persistent_context takes an exclusive lock, so two processes on
one profile produce the about:blank tabs seen earlier in this project.
Each market therefore gets its own profile directory.

Extensions are staged per profile as "<profile>_extensions" (see
google_ai_mode._staged_extensions_dir), so captcha-raptor is copied into
each one rather than shared. That copy is what survives a profile reset:
when a session is rebuilt after a rate limit, the staged directory sits
OUTSIDE the profile and is reinstalled automatically.

Usage:
    python setup_drone_profiles.py            # create all ten
    python setup_drone_profiles.py --check    # report state, change nothing
"""
from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

from drone_markets import MARKETS  # noqa: E402

DATA = ROOT / "data"
# The profile the six finished markets used. Its staged captcha-raptor is
# the copy source, so every new profile starts with a solver known to
# work in this environment.
SOURCE_EXT = DATA / "ai_mode_chrome_profile_extensions" / "captcha-raptor"


def profile_dir(slug: str) -> Path:
    """Base profile path for a market.

    google_ai_mode appends "_chromium" when extensions force the bundled
    build, so this is the BASE name and the real directory will be
    "<slug>_chromium".
    """
    return DATA / f"drone_profile_{slug}"


def staged_ext_dir(slug: str) -> Path:
    """Where google_ai_mode looks for this profile's staged extensions.

    Derived from the UNSUFFIXED profile name. google_ai_mode._staging_root
    builds it from _profile_dir(), i.e. before the "_chromium" suffix is
    applied, so the staged copy is shared by both the Chrome and Chromium
    builds. Guessing "<name>_chromium_extensions" put the extensions where
    nothing looked for them, and ten browsers launched with no CAPTCHA
    solver at all.
    """
    base = profile_dir(slug)
    return base.parent / f"{base.name}_extensions"


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--check", action="store_true")
    args = ap.parse_args()

    if not SOURCE_EXT.is_dir():
        print(f"ERROR: captcha-raptor not found at {SOURCE_EXT}")
        print("The working profile must have it staged before cloning.")
        raise SystemExit(1)

    manifest = SOURCE_EXT / "manifest.json"
    if not manifest.is_file():
        print(f"ERROR: {SOURCE_EXT} has no manifest.json - not an extension")
        raise SystemExit(1)

    print(f"source extension: {SOURCE_EXT}")
    print(f"markets: {len(MARKETS)}")
    print("")

    for slug, market, _scope in MARKETS:
        ext = staged_ext_dir(slug)
        have = (ext / "manifest.json").is_file()
        if args.check:
            print(f"  {slug:22} extension staged: {have}")
            continue
        if have:
            print(f"  {slug:22} already staged")
            continue
        ext.parent.mkdir(parents=True, exist_ok=True)
        shutil.copytree(SOURCE_EXT, ext, dirs_exist_ok=True)
        print(f"  {slug:22} staged -> {ext.name}")

    if not args.check:
        print("")
        print("Profiles themselves are created on first launch by")
        print("launch_persistent_context; only the extensions need staging.")


if __name__ == "__main__":
    main()
