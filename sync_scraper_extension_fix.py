#!/usr/bin/env python3
"""Copy the patched google-ai-scraper content.js over every other copy of it
on this machine.

Why this exists: there are several checked-out copies of the scraper
extension in Downloads (Coherent-Quadrant, chrome-scraper-extensions,
vendor-intel/Brands V2, ...). Chrome loads ONE of them as an unpacked
extension, and it is not obvious which. Patching a single copy therefore
looks like it did nothing — the reload picks up a different, unpatched file.

Only Chrome (MV3) copies are synced; firefox-extension/ directories are left
alone because their content.js may diverge from the Chrome one.

Usage:
    python sync_scraper_extension_fix.py            # dry run
    python sync_scraper_extension_fix.py --apply
"""
import argparse
import filecmp
import shutil
from pathlib import Path

SOURCE = Path(
    r"C:\Users\aman.kumar\Downloads\Coherent-Quadrant\google-ai-scraper-main"
    r"\google-ai-scraper-main\extension\content.js"
)
SEARCH_ROOT = Path(r"C:\Users\aman.kumar\Downloads")
# Marker proving a file carries the truncation-detection fix.
MARKER = "promisesRosterButHasNone"


def extension_name(manifest_path: Path) -> str:
    import json

    try:
        return str(json.loads(manifest_path.read_text(encoding="utf-8")).get("name") or "")
    except Exception:  # noqa: BLE001
        return ""


def find_targets() -> list[Path]:
    """Chrome copies of the SAME extension as SOURCE.

    Matching on manifest name matters: Downloads also contains a different,
    much smaller extension ("Coherent Quadrant — AI Overview Collector",
    176 lines vs 872) whose content.js would be destroyed by this sync.
    """
    source_name = extension_name(SOURCE.parent / "manifest.json")
    if not source_name:
        raise SystemExit(f"cannot read extension name from {SOURCE.parent / 'manifest.json'}")

    targets = []
    for path in SEARCH_ROOT.rglob("content.js"):
        if path.resolve() == SOURCE.resolve():
            continue
        # Chrome unpacked extensions live in an "extension" directory next to
        # a manifest.json; skip Firefox variants and unrelated content.js.
        if path.parent.name != "extension":
            continue
        manifest = path.parent / "manifest.json"
        if not manifest.is_file():
            continue
        if extension_name(manifest) != source_name:
            continue
        targets.append(path)
    return sorted(targets)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--apply", action="store_true", help="Write the files (otherwise dry-run)")
    args = parser.parse_args()

    if not SOURCE.is_file():
        raise SystemExit(f"source not found: {SOURCE}")
    source_text = SOURCE.read_text(encoding="utf-8")
    if MARKER not in source_text:
        raise SystemExit(
            f"source {SOURCE} does not contain {MARKER!r} — refusing to "
            "propagate an unpatched file"
        )

    targets = find_targets()
    print(f"source : {SOURCE}")
    print(f"targets: {len(targets)} other Chrome extension copies\n")

    changed = 0
    for target in targets:
        already = MARKER in target.read_text(encoding="utf-8", errors="replace")
        same = filecmp.cmp(SOURCE, target, shallow=False)
        status = "up-to-date" if same else ("patched-but-differs" if already else "NEEDS FIX")
        print(f"  [{status:19}] {target}")
        if same:
            continue
        changed += 1
        if args.apply:
            backup = target.with_suffix(".js.bak")
            if not backup.exists():
                shutil.copy2(target, backup)
            shutil.copy2(SOURCE, target)

    print()
    if not changed:
        print("All copies already identical to the patched source.")
    elif args.apply:
        print(f"Updated {changed} copies (originals saved as content.js.bak).")
        print("Now reload the extension in chrome://extensions.")
    else:
        print(f"{changed} copies would be updated. Re-run with --apply.")


if __name__ == "__main__":
    main()
