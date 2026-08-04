#!/usr/bin/env python3
"""Print Agent Reach + vendor-intel integration status."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from dotenv import load_dotenv

load_dotenv(ROOT / ".env")


def main() -> None:
    from vendor_intel.integrations.agent_reach_client import (
        agent_reach_enabled,
        agent_reach_root,
        doctor_summary,
        exa_search_enabled,
        full_channel_report,
    )
    from vendor_intel.integrations.linkedin_mcp import linkedin_mcp_enabled, linkedin_mcp_status

    print("=== Vendor-intel × Agent Reach ===")
    print(f"Root: {agent_reach_root()}")
    print(f"AGENT_REACH_ENABLED: {agent_reach_enabled()}")
    print(f"AGENT_REACH_EXA_SEARCH: {exa_search_enabled()}")
    print(f"AGENT_REACH_LINKEDIN_MCP: {linkedin_mcp_enabled()}")
    print(f"LinkedIn MCP: {linkedin_mcp_status()}")
    print()

    summary = doctor_summary()
    for key, val in summary.items():
        safe = str(val).encode("ascii", "replace").decode("ascii")
        print(f"  {key}: {safe}")

    report = full_channel_report()
    if report:
        print()
        print(report.encode("ascii", "replace").decode("ascii"))


if __name__ == "__main__":
    main()
