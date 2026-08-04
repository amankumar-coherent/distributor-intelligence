"""Agent Reach integration — Jina web read + optional Exa search via mcporter.

Agent-Reach-main lives at repo root: Cmi/Agent-Reach-main (override with AGENT_REACH_ROOT).
Enable with AGENT_REACH_ENABLED=true (default). Jina read needs no extra setup.
Exa search: AGENT_REACH_EXA_SEARCH=true + npm install -g mcporter + mcporter config add exa https://mcp.exa.ai/mcp

Wired into vendor-intel:
  - page_fetch / contact_page / linkedin_enrich — Jina Reader for any URL
  - search_router — Exa semantic search when DDGS/SearXNG/Bing return few hits
  - linkedin_mcp — search_people, get_company_profile, get_company_employees (mcporter)
"""
from __future__ import annotations

import asyncio
import json
import os
import re
import subprocess
import sys
from pathlib import Path
from typing import Any

from vendor_intel.config import _project_root

_web_channel: Any | None = None
_path_ready = False


def _env_true(name: str, default: bool = True) -> bool:
    raw = os.getenv(name, "").strip().lower()
    if not raw:
        return default
    return raw in ("1", "true", "yes", "on")


def agent_reach_enabled() -> bool:
    return _env_true("AGENT_REACH_ENABLED", default=True)


def _discover_agent_reach_root() -> Path | None:
    """Find Agent-Reach-main by walking up from this file, cwd, and project root."""
    seen: set[str] = set()
    seeds: list[Path] = [
        Path(__file__).resolve(),
        Path.cwd().resolve(),
        _project_root(),
    ]
    cmi_raw = (os.getenv("CMI_ROOT") or os.getenv("REPO_ROOT") or "").strip()
    if cmi_raw:
        seeds.insert(0, Path(cmi_raw).resolve())

    def _check(parent: Path) -> Path | None:
        key = str(parent).lower()
        if key in seen:
            return None
        seen.add(key)
        candidate = parent / "Agent-Reach-main"
        if (candidate / "agent_reach").is_dir():
            return candidate
        return None

    for seed in seeds:
        if not seed:
            continue
        start = seed if seed.is_dir() else seed.parent
        for parent in [start, *start.parents]:
            hit = _check(parent)
            if hit:
                return hit
            if parent.name.lower() == "cmi":
                hit = _check(parent)
                if hit:
                    return hit

    # vendor-intel often symlinked to coherent — Agent-Reach sits in sibling Cmi/ folder
    for seed in seeds:
        for parent in seed.parents:
            if parent.name.lower() == "desktop":
                cmi_reach = parent / "Cmi" / "Agent-Reach-main"
                if (cmi_reach / "agent_reach").is_dir():
                    return cmi_reach
                break

    fallback = Path.home() / "Desktop" / "Cmi" / "Agent-Reach-main"
    if (fallback / "agent_reach").is_dir():
        return fallback
    return None


def agent_reach_root() -> Path:
    raw = (os.getenv("AGENT_REACH_ROOT") or "").strip()
    if raw:
        return Path(raw).resolve()
    found = _discover_agent_reach_root()
    if found:
        return found
    return _project_root().parent.parent / "Agent-Reach-main"


def _ensure_path() -> bool:
    global _path_ready
    if _path_ready:
        return True
    if not agent_reach_enabled():
        return False
    root = agent_reach_root()
    if not (root / "agent_reach").is_dir():
        return False
    root_str = str(root)
    if root_str not in sys.path:
        sys.path.insert(0, root_str)
    _path_ready = True
    return True


def _get_web_channel() -> Any | None:
    global _web_channel
    if _web_channel is not None:
        return _web_channel
    if not _ensure_path():
        return None
    try:
        from agent_reach.channels.web import WebChannel

        _web_channel = WebChannel()
        return _web_channel
    except Exception:
        return None


async def read_url_markdown(url: str, *, timeout: float = 35.0) -> str:
    """Read any URL as clean markdown via Agent Reach Jina Reader."""
    if not agent_reach_enabled() or not url.strip():
        return ""
    channel = _get_web_channel()
    if channel is None:
        return ""

    def _read() -> str:
        try:
            return channel.read(url) or ""
        except Exception:
            return ""

    try:
        text = await asyncio.wait_for(asyncio.to_thread(_read), timeout=timeout)
        return (text or "").strip()[:20000]
    except Exception:
        return ""


def _mcporter_probe() -> tuple[bool, str]:
    """Return (ok, detail) for mcporter + configured MCP servers."""
    if not _ensure_path():
        return False, "Agent-Reach-main not on path"
    try:
        from agent_reach.probe import probe_command

        probe = probe_command("mcporter", ["config", "list"], timeout=8, package="mcporter")
        if probe.status == "missing":
            return False, "mcporter not installed (npm install -g mcporter)"
        if not probe.ok:
            return False, probe.hint or probe.output or probe.status or "mcporter error"
        out = (probe.output or "").lower()
        if "exa" not in out:
            return False, "exa MCP not configured (mcporter config add exa https://mcp.exa.ai/mcp)"
        return True, "mcporter + exa MCP"
    except Exception as exc:
        return False, str(exc)


def _mcporter_available() -> bool:
    ok, _ = _mcporter_probe()
    return ok


def _mcporter_has(name: str) -> bool:
    """True when mcporter lists an MCP server (e.g. linkedin, exa)."""
    if not _ensure_path():
        return False
    try:
        from agent_reach.probe import probe_command

        probe = probe_command("mcporter", ["config", "list"], timeout=8, package="mcporter")
        return probe.ok and name.lower() in (probe.output or "").lower()
    except Exception:
        return False


def exa_search_enabled() -> bool:
    return agent_reach_enabled() and _env_true("AGENT_REACH_EXA_SEARCH", default=False)


async def exa_web_search(query: str, *, num_results: int = 5) -> list[dict[str, str]]:
    """Optional Exa semantic search via mcporter (requires AGENT_REACH_EXA_SEARCH=true)."""
    if not exa_search_enabled() or not query.strip():
        return []
    if not _mcporter_available():
        return []

    safe_q = query.replace('"', "'").replace("\\", " ")[:500]
    call_arg = f'exa.web_search_exa(query: "{safe_q}", numResults: {max(1, min(num_results, 10))})'

    def _run() -> str:
        cmds = ["mcporter.cmd", "mcporter"] if sys.platform == "win32" else ["mcporter"]
        for cmd in cmds:
            try:
                proc = subprocess.run(
                    [cmd, "call", call_arg],
                    capture_output=True,
                    text=True,
                    timeout=45,
                    check=False,
                )
                out = (proc.stdout or proc.stderr or "").strip()
                if out:
                    return out
            except Exception:
                continue
        return ""

    raw = await asyncio.to_thread(_run)
    if not raw:
        return []

    hits: list[dict[str, str]] = []
    # mcporter may return JSON or plain text with URLs
    try:
        data = json.loads(raw)
        items = data if isinstance(data, list) else data.get("results") or data.get("items") or []
        if isinstance(items, list):
            for item in items[:num_results]:
                if not isinstance(item, dict):
                    continue
                hits.append(
                    {
                        "title": str(item.get("title") or item.get("name") or "")[:200],
                        "link": str(item.get("url") or item.get("link") or item.get("href") or ""),
                        "snippet": str(item.get("snippet") or item.get("text") or item.get("description") or "")[:500],
                    }
                )
            if hits:
                return hits
    except json.JSONDecodeError:
        pass

    for m in re.finditer(r"https?://[^\s\)\]\"']+", raw):
        link = m.group(0).rstrip(".,;")
        if link and not any(h.get("link") == link for h in hits):
            hits.append({"title": link[:80], "link": link, "snippet": raw[:300]})
        if len(hits) >= num_results:
            break
    return hits


# Channels vendor-intel enrichment uses via Jina (web/linkedin/github public URLs).


def doctor_summary() -> dict[str, str]:
    """Quick status for web + exa + enrichment-relevant Agent Reach channels."""
    out: dict[str, str] = {"root": str(agent_reach_root()), "enabled": str(agent_reach_enabled())}
    if not agent_reach_enabled():
        out["web"] = "disabled"
        out["exa"] = "disabled"
        return out
    if not _ensure_path():
        out["web"] = "missing (Agent-Reach-main not found)"
        out["exa"] = "missing"
        return out
    channel = _get_web_channel()
    if channel:
        try:
            status, msg = channel.check()
            out["web"] = f"{status}: {msg[:120]}"
        except Exception as exc:
            out["web"] = f"error: {exc}"
    else:
        out["web"] = "unavailable"
    if exa_search_enabled():
        exa_ok, exa_detail = _mcporter_probe()
        out["exa"] = f"ok ({exa_detail})" if exa_ok else f"off ({exa_detail})"
    else:
        out["exa"] = "disabled (set AGENT_REACH_EXA_SEARCH=true)"
    try:
        from vendor_intel.integrations.linkedin_mcp import linkedin_mcp_status

        out["linkedin_mcp"] = linkedin_mcp_status()
    except Exception as exc:
        out["linkedin_mcp"] = f"error: {exc}"
    try:
        from agent_reach.channels import get_channel

        for name in ("linkedin", "github"):
            ch = get_channel(name)
            if ch:
                status, msg = ch.check()
                out[name] = f"{status}: {str(msg)[:100]}"
    except Exception:
        pass
    return out


def full_channel_report() -> str:
    """Plain-text Agent Reach doctor report (all platforms)."""
    if not _ensure_path():
        return ""
    try:
        from agent_reach.config import Config
        from agent_reach.doctor import check_all

        results = check_all(Config())
        lines = ["=== Agent Reach channels (all platforms) ==="]
        for key in sorted(results.keys()):
            r = results[key]
            status = r.get("status", "?")
            desc = r.get("name", key)
            msg = str(r.get("message") or "")[:100].replace("\n", " ")
            backend = r.get("active_backend") or ""
            suffix = f" [{backend}]" if backend else ""
            lines.append(f"  {status:5} {key:12} {desc}{suffix} — {msg}")
        return "\n".join(lines)
    except Exception as exc:
        return f"Agent Reach full doctor failed: {exc}"


def log_agent_reach_status() -> None:
    """Print one-line Agent Reach status at pipeline startup."""
    if not agent_reach_enabled():
        print("  [agent-reach] Disabled (AGENT_REACH_ENABLED=false)", flush=True)
        return
    s = doctor_summary()
    root = s.get("root", "?")
    if "not found" in s.get("web", "").lower() or "missing" in s.get("web", "").lower():
        print(f"  [agent-reach] Not found at {root} — set AGENT_REACH_ROOT", flush=True)
        return
    web = str(s.get("web", "?"))[:80].encode("ascii", "replace").decode("ascii")
    print(f"  [agent-reach] Web (Jina): {web}", flush=True)
    if exa_search_enabled():
        exa = str(s.get("exa", "?")).encode("ascii", "replace").decode("ascii")
        print(f"  [agent-reach] Exa search: {exa}", flush=True)
    if "linkedin_mcp" in s:
        li_mcp = str(s["linkedin_mcp"]).encode("ascii", "replace").decode("ascii")
        print(f"  [agent-reach] LinkedIn MCP: {li_mcp}", flush=True)
    for ch in ("linkedin", "github"):
        if ch in s:
            line = str(s[ch])[:80].encode("ascii", "replace").decode("ascii")
            print(f"  [agent-reach] {ch.title()}: {line}", flush=True)
