#!/usr/bin/env python3
"""Add / list / test Chrome extensions in the AI Mode automation profile.

Why this exists: adding an extension to the AI Mode scraper is easy to get
subtly wrong, and the failure is SILENT — Chrome accepts
`--load-extension` and starts normally while the extension never actually
registers. Two traps in particular:

  1. The path must be the folder that DIRECTLY contains manifest.json.
     Pointing at a repo root whose manifest sits in `extension/` loads
     nothing (this script auto-detects that layout and corrects it).
  2. Chrome will not enable an unpacked extension in an automation profile
     that previously ran WITHOUT it — measured repeatedly: existing profile
     => 0 service workers, fresh profile with identical flags => 1 worker.
     So the profile has to be rebuilt after adding one. `--add` does that
     for you.

Usage:
    # RECOMMENDED — install into the profile so it persists, then verify
    python manage_ai_mode_extensions.py --install "C:\\path\\to\\extension" --test

    # add extensions BY HAND in a real Chrome window (persists now)
    python manage_ai_mode_extensions.py --open

    # what is configured / installed, and does it actually load?
    python manage_ai_mode_extensions.py --list

    # env-var route (legacy: --load-extension, does not persist in profile)
    python manage_ai_mode_extensions.py --add "C:\\path\\to\\extension" --test

    # remove one
    python manage_ai_mode_extensions.py --remove "C:\\path\\to\\extension"

    # just verify what the browser really registered
    python manage_ai_mode_extensions.py --test

    # wipe browsing state, KEEPING installed extensions
    python manage_ai_mode_extensions.py --reset-profile

--install vs --add: `--install` COPIES the extension into the automation
profile, so it survives rebuilds and shows up in chrome://extensions like
a normal install. `--add` only records a path in .env, which Chrome loads
via --load-extension — a command-line-only load Chrome never persists.
Prefer --install.

IMPORTANT — what an extension can and cannot fix here. A CAPTCHA-solver
extension only helps when Google actually renders a solvable reCAPTCHA
widget. captcha-raptor, for example, injects only into
`*://*.google.com/recaptcha/*` and clicks image tiles. When Google's
/sorry/ page reports "Cannot contact reCAPTCHA", the widget never loads at
all, so there is nothing for ANY solver extension to act on — that block is
avoided by pacing (GOOGLE_AI_MODE_QUERY_DELAY), not solved by an extension.
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

ENV_PATH = ROOT / ".env"
ENV_KEY = "GOOGLE_AI_MODE_EXTENSIONS"


def _resolve(raw: str) -> Path | None:
    """Return the folder that actually holds manifest.json, or None."""
    path = Path(raw.strip().strip('"'))
    if not path.is_dir():
        print(f"  ERROR: not a directory: {path}")
        return None
    if (path / "manifest.json").is_file():
        return path.resolve()
    nested = path / "extension"
    if (nested / "manifest.json").is_file():
        print(f"  note: using {nested} (manifest.json is in the extension/ subfolder)")
        return nested.resolve()
    print(f"  ERROR: no manifest.json in {path} (or its extension/ subfolder)")
    return None


def _read_env_paths() -> list[str]:
    if not ENV_PATH.is_file():
        return []
    for line in ENV_PATH.read_text(encoding="utf-8").splitlines():
        if line.startswith(f"{ENV_KEY}="):
            value = line.split("=", 1)[1].strip()
            return [p for p in value.split(";") if p.strip()]
    return []


def _write_env_paths(paths: list[str]) -> None:
    """Rewrite (or append) the ENV_KEY line, preserving everything else."""
    joined = ";".join(paths)
    new_line = f"{ENV_KEY}={joined}"
    lines = ENV_PATH.read_text(encoding="utf-8").splitlines() if ENV_PATH.is_file() else []
    for i, line in enumerate(lines):
        if line.startswith(f"{ENV_KEY}="):
            lines[i] = new_line
            break
    else:
        lines.append("")
        lines.append("# Chrome extensions loaded into the AI Mode automation browser.")
        lines.append(new_line)
    ENV_PATH.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"  .env updated: {ENV_KEY}={joined or '(empty)'}")


def _list_profile_state() -> None:
    """Report what the PROFILE itself holds — the persistent part."""
    from vendor_intel.scraping.google_ai_mode import (
        _installed_extension_ids,
        active_profile_dir,
        staged_extension_paths,
    )

    profile = Path(active_profile_dir())
    print(f"Automation profile (next launch will use this): {profile}")
    print(f"  exists: {profile.exists()}")

    staged = staged_extension_paths()
    print("Installed via --install (persists, survives resets):")
    if not staged:
        print("  (none)")
    for p in staged:
        print(f"  OK  {p}")

    ids = _installed_extension_ids(profile)
    print("Installed in the profile (chrome://extensions, persists):")
    if not ids:
        print("  (none recorded yet — add one with --open or --install)")
    for i in ids:
        print(f"  {i}")
    print()


def _test_live() -> int:
    """Launch the automation browser and report what really registered."""
    from vendor_intel.placeholders.load_keys import apply_env_overrides

    apply_env_overrides()
    from vendor_intel.scraping import google_ai_mode as aim

    configured = aim.extension_paths() + aim.staged_extension_paths()
    installed = aim._installed_extension_ids(Path(aim.active_profile_dir()))
    print(f"  resolved extension paths: {configured or 'NONE'}")
    print(f"  profile-installed extension ids: {installed or 'NONE'}")
    if not configured and not installed:
        print("  nothing to test — install one with --install or --open first")
        return 1

    print("  launching automation browser...")
    aim._get_page()
    time.sleep(6)  # MV3 service workers register lazily
    ctx = aim._context

    # Attribute workers to OUR extensions. A bare worker count is a false
    # positive: Chrome always runs its own component extensions (pdf,
    # hangout_services, network_speech_synthesis), which previously made a
    # never-loaded solver report as active.
    ours = aim._own_extension_workers(ctx, configured)

    print(f"  all workers: {len(ctx.service_workers)} (incl. Chrome's own components)")
    for w in ctx.service_workers:
        mine = "  <-- ours" if w.url in ours else ""
        print(f"    {w.url[:95]}{mine}")
    print(f"  OUR extension workers: {len(ours)}")

    # Ask Chrome itself, which is the only ground truth here.
    try:
        page = ctx.new_page()
        page.goto("chrome://extensions-internals/", wait_until="load", timeout=20_000)
        time.sleep(2)
        data = json.loads(page.inner_text("body"))
        rows = [e for e in data if e.get("location") != "COMPONENT"]
        print(f"  chrome://extensions-internals reports {len(rows)} non-component extension(s):")
        for e in rows:
            print(f"    {e.get('name')} [{e.get('location')}]")
        if not rows:
            print("    NONE — the extension is genuinely not installed")
    except Exception as exc:
        print(f"  (could not read extensions-internals: {exc})")

    print(f"  EXTENSION ACTIVE: {bool(ours)}")
    aim.shutdown()
    return 0 if ours else 1


def main() -> None:
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    ap.add_argument("--install", action="append", default=[], metavar="PATH",
                    help="RECOMMENDED: copy an extension INTO the profile so it "
                         "persists across runs (repeatable).")
    ap.add_argument("--open", action="store_true",
                    help="Open the automation profile's Chrome on "
                         "chrome://extensions so you can add extensions by hand. "
                         "Manual installs now persist.")
    ap.add_argument("--reset-profile", action="store_true",
                    help="Wipe the profile's browsing state but KEEP installed "
                         "extensions (use if Google has flagged the session).")
    ap.add_argument("--add", action="append", default=[], metavar="PATH",
                    help="Legacy: record a path in .env for --load-extension. "
                         "Does not persist in the profile; prefer --install.")
    ap.add_argument("--remove", action="append", default=[], metavar="PATH",
                    help="Extension folder to remove from .env (repeatable).")
    ap.add_argument("--list", action="store_true",
                    help="Show configured, staged, and profile-installed extensions")
    ap.add_argument("--test", action="store_true",
                    help="Launch the browser and verify what actually registered")
    args = ap.parse_args()

    if not (args.add or args.remove or args.list or args.test
            or args.install or args.open or args.reset_profile):
        ap.print_help()
        return

    current = _read_env_paths()

    if args.list:
        print("Configured via .env (GOOGLE_AI_MODE_EXTENSIONS, --load-extension):")
        if not current:
            print("  (none)")
        for p in current:
            ok = (Path(p) / "manifest.json").is_file()
            print(f"  {'OK ' if ok else 'BAD'} {p}")
            if not ok:
                print("      no manifest.json here — this will silently not load")
        print()
        _list_profile_state()

    if args.install:
        from vendor_intel.scraping.google_ai_mode import install_extension

        for raw in args.install:
            print(f"Installing into the profile: {raw}")
            if install_extension(raw):
                print("  OK — persists across runs")
            else:
                print("  FAILED — see the message above")
        print()

    if args.reset_profile:
        from vendor_intel.scraping.google_ai_mode import (
            _rebuild_profile_preserving_staged,
            active_profile_dir,
        )

        target = Path(active_profile_dir())
        if target.exists():
            print(f"Resetting browsing state, keeping extensions: {target}")
            _rebuild_profile_preserving_staged(target)
            print("  done — profile will re-seed from real Chrome on next run")
        else:
            print(f"  profile does not exist yet: {target}")
        print()

    if args.open:
        from vendor_intel.placeholders.load_keys import apply_env_overrides

        apply_env_overrides()
        from vendor_intel.scraping.google_ai_mode import open_profile_for_extensions

        open_profile_for_extensions()
        print()
        _list_profile_state()

    changed = False

    for raw in args.add:
        print(f"Adding: {raw}")
        resolved = _resolve(raw)
        if not resolved:
            continue
        if str(resolved) in current:
            print("  already configured, skipping")
            continue
        current.append(str(resolved))
        changed = True
        print(f"  added: {resolved}")

    for raw in args.remove:
        print(f"Removing: {raw}")
        target = str(Path(raw.strip().strip('"')).resolve())
        before = len(current)
        current = [p for p in current if str(Path(p).resolve()) != target]
        if len(current) == before:
            print("  was not configured, nothing to remove")
        else:
            changed = True
            print("  removed")

    if changed:
        _write_env_paths(current)
        # No automatic profile deletion any more: it used to wipe manually
        # installed extensions, which is the exact problem --install and
        # --open exist to solve. Reset explicitly if you actually need it.
        print("  note: --add uses --load-extension, which Chrome does not persist.")
        print("        If it does not register, install it instead:")
        print("          python manage_ai_mode_extensions.py --install <path>")

    if args.test:
        print()
        print("Live verification:")
        sys.exit(_test_live())


if __name__ == "__main__":
    main()
