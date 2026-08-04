#!/usr/bin/env python3
"""Interactive log watcher for active market scrapes."""
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = ROOT / "output"

# Reconfigure stdout/stderr for clean Unicode output on Windows
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


def get_active_logs() -> list[tuple[str, Path]]:
    if not OUTPUT_DIR.is_dir():
        return []
    logs = []
    for d in OUTPUT_DIR.iterdir():
        if d.is_dir():
            log_file = d / "scrape_live.log"
            if log_file.is_file():
                logs.append((d.name, log_file))
    return logs


def tail_file(filepath: Path, num_lines: int = 40):
    print(f"\n--- Tail of {filepath.name} ({filepath.parent.name}) ---")
    try:
        with open(filepath, "r", encoding="utf-8", errors="replace") as f:
            # Simple tail implementation
            lines = f.readlines()
            for line in lines[-num_lines:]:
                print(line, end="")
            
            # Keep reading new lines
            f.seek(0, os.SEEK_END)
            while True:
                line = f.readline()
                if not line:
                    time.sleep(0.5)
                    continue
                print(line, end="")
                sys.stdout.flush()
    except KeyboardInterrupt:
        print("\n\nStopped watching logs.")
    except Exception as e:
        print(f"\nError reading log file: {e}")


def main():
    print("=" * 60)
    print("  Market Scrape Log Watcher")
    print("=" * 60)
    
    logs = get_active_logs()
    if not logs:
        print(f"\nNo active scrape logs found in: {OUTPUT_DIR.resolve()}")
        print("Logs are created when you start running the scrape scripts (*_scrape.py).")
        input("\nPress Enter to exit...")
        return

    # Sort logs by modification time (most recent first)
    logs.sort(key=lambda x: x[1].stat().st_mtime, reverse=True)

    print("\nSelect a market log to view:")
    for idx, (market_name, path) in enumerate(logs, 1):
        mtime = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(path.stat().st_mtime))
        # Highlight the most recently active one
        recent_label = " (Most Active)" if idx == 1 else ""
        print(f"  [{idx}] {market_name:<30} (Last Updated: {mtime}){recent_label}")

    choice = input(f"\nEnter choice [1-{len(logs)}] (default 1): ").strip()
    if not choice:
        selected_idx = 0
    else:
        try:
            selected_idx = int(choice) - 1
            if selected_idx < 0 or selected_idx >= len(logs):
                print("Invalid selection. Exiting.")
                return
        except ValueError:
            print("Invalid input. Exiting.")
            return

    market_name, log_path = logs[selected_idx]
    print(f"\nWatching logs for: {market_name}")
    print("Press Ctrl+C to exit watching.")
    tail_file(log_path)


if __name__ == "__main__":
    main()
