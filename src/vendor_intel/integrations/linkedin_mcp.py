"""LinkedIn MCP (linkedin-scraper-mcp via mcporter) for vendor-intel enrichment.

Requires:
  pip install linkedin-scraper-mcp
  linkedin-scraper-mcp --transport streamable-http --port 8001
  mcporter config add linkedin http://localhost:8001/mcp

Enable: AGENT_REACH_LINKEDIN_MCP=true (and AGENT_REACH_ENABLED=true)
"""
from __future__ import annotations

import asyncio
import json
import os
import random
import re
import subprocess
import sys
from pathlib import Path
from typing import Any

from vendor_intel.integrations.agent_reach_client import (
    _ensure_path,
    _env_true,
    _mcporter_has,
    agent_reach_enabled,
)

_EXEC_TITLE = re.compile(
    r"\b(ceo|chief executive|ciso|chief security|founder|co-founder|president|"
    r"managing director|head of security|vp security|director of security)\b",
    re.I,
)

_server_prefix_cache: str | None = None
_prefix_checked = False
_rate_lock = asyncio.Lock()
_last_mcp_call_at: float = 0.0

# Direct streamable-http session (avoids mcporter CLI zombies that block the MCP queue)
_http_lock = asyncio.Lock()
_http_cm: Any = None
_http_session: Any = None
_http_session_ready = False


def linkedin_mcp_delay() -> float:
    """Seconds to wait between LinkedIn MCP calls (avoid blocks)."""
    try:
        base = float(os.getenv("LINKEDIN_MCP_DELAY_SEC", "4.0"))
    except ValueError:
        base = 4.0
    try:
        jitter = float(os.getenv("LINKEDIN_MCP_DELAY_JITTER", "1.5"))
    except ValueError:
        jitter = 1.5
    if jitter > 0:
        return max(0.5, base + random.uniform(0, jitter))
    return max(0.5, base)


async def _wait_linkedin_rate_limit() -> None:
    global _last_mcp_call_at
    delay = linkedin_mcp_delay()
    async with _rate_lock:
        import time

        now = time.monotonic()
        wait = delay - (now - _last_mcp_call_at)
        if wait > 0:
            await asyncio.sleep(wait)
        _last_mcp_call_at = time.monotonic()


def linkedin_mcp_enabled() -> bool:
    return (
        agent_reach_enabled()
        and _env_true("AGENT_REACH_LINKEDIN_MCP", default=False)
        and _ensure_path()
    )


def _linkedin_mcp_port() -> tuple[str, int]:
    host = (os.getenv("LINKEDIN_MCP_HOST") or "127.0.0.1").strip() or "127.0.0.1"
    try:
        port = int(os.getenv("LINKEDIN_MCP_PORT", "8001"))
    except ValueError:
        port = 8001
    return host, port


def _linkedin_mcp_port_open() -> bool:
    """True when streamable-http LinkedIn MCP is accepting TCP on the configured port."""
    import socket

    host, port = _linkedin_mcp_port()
    try:
        with socket.create_connection((host, port), timeout=1.5):
            return True
    except OSError:
        return False


def _mcporter_config_lists_linkedin() -> bool:
    """Read local mcporter.json without invoking mcporter CLI (avoids Windows probe flakes)."""
    here = Path(__file__).resolve()
    candidates = [
        Path.cwd() / "config" / "mcporter.json",
        here.parents[3] / "config" / "mcporter.json",  # project root
    ]
    for path in candidates:
        try:
            if not path.is_file():
                continue
            data = json.loads(path.read_text(encoding="utf-8"))
            servers = data.get("mcpServers") or data.get("servers") or {}
            if isinstance(servers, dict) and any(k.lower() == "linkedin" for k in servers):
                return True
        except Exception:
            continue
    return False


def linkedin_mcp_available() -> bool:
    """True when LinkedIn MCP enrichment can run.

    Prefer a live :8001 listener. Do not rely solely on ``mcporter config list`` —
    that CLI probe often false-negatives on Windows even when MCP is up.
    """
    if not linkedin_mcp_enabled():
        return False
    if _linkedin_mcp_port_open():
        return True
    if _mcporter_config_lists_linkedin():
        return True
    return _mcporter_has("linkedin")


def linkedin_mcp_status() -> str:
    if not linkedin_mcp_enabled():
        return "disabled (set AGENT_REACH_LINKEDIN_MCP=true)"
    host, port = _linkedin_mcp_port()
    if _linkedin_mcp_port_open():
        prefix = _linkedin_server_prefix() or "linkedin"
        return f"ok (server={prefix}, {host}:{port} up)"
    if _mcporter_has("linkedin") or _mcporter_config_lists_linkedin():
        return f"configured but {host}:{port} not listening — start linkedin-scraper-mcp"
    return "off (mcporter linkedin MCP not configured)"


def _mcporter_commands() -> list[str]:
    if sys.platform == "win32":
        return ["mcporter.cmd", "mcporter"]
    return ["mcporter"]


def _linkedin_server_prefix() -> str | None:
    """Return mcporter server name for LinkedIn tools (default: linkedin)."""
    global _server_prefix_cache, _prefix_checked
    if _prefix_checked:
        return _server_prefix_cache
    _prefix_checked = True
    # Prefer live server / local config over flaky mcporter CLI probe
    if _linkedin_mcp_port_open() or _mcporter_config_lists_linkedin():
        _server_prefix_cache = "linkedin"
        return _server_prefix_cache
    if not _mcporter_has("linkedin"):
        _server_prefix_cache = None
        return None
    try:
        from agent_reach.probe import probe_command

        probe = probe_command("mcporter", ["config", "list"], timeout=8, package="mcporter")
        out = (probe.output or "").lower()
        for candidate in ("linkedin-scraper", "mcp-server-linkedin", "linkedin"):
            if candidate in out:
                _server_prefix_cache = candidate
                return _server_prefix_cache
    except Exception:
        pass
    _server_prefix_cache = "linkedin"
    return _server_prefix_cache


def _format_call_args(**kwargs: Any) -> str:
    parts: list[str] = []
    for key, val in kwargs.items():
        if val is None or val == "":
            continue
        if isinstance(val, bool):
            parts.append(f"{key}: {'true' if val else 'false'}")
        elif isinstance(val, (int, float)):
            parts.append(f"{key}: {val}")
        else:
            safe = str(val).replace("\\", " ").replace('"', "'")[:500]
            parts.append(f'{key}: "{safe}"')
    return ", ".join(parts)


def _linkedin_mcp_http_url() -> str:
    host, port = _linkedin_mcp_port()
    # Prefer IPv4 loopback; "localhost" can resolve to ::1 and miss the listener
    if host in ("127.0.0.1", "localhost", "0.0.0.0", "::1"):
        host = "127.0.0.1"
    return (os.getenv("LINKEDIN_MCP_URL") or f"http://{host}:{port}/mcp").strip()


def _call_timeout_sec() -> float:
    try:
        timeout = float(os.getenv("LINKEDIN_MCP_CALL_TIMEOUT_SEC", "120"))
    except ValueError:
        timeout = 120.0
    return max(30.0, timeout)


async def _close_http_session() -> None:
    """No-op placeholder — HTTP sessions are opened per call (cancel-scope safe)."""
    return


async def _invoke_tool_http(tool: str, **params: Any) -> Any:
    """Call LinkedIn MCP tools via streamable-http (preferred over mcporter CLI)."""
    from datetime import timedelta

    from mcp import ClientSession
    from mcp.client.streamable_http import streamablehttp_client

    timeout = _call_timeout_sec()
    url = _linkedin_mcp_http_url()
    async with _http_lock:
        try:
            async with streamablehttp_client(
                url,
                timeout=timedelta(seconds=min(60.0, timeout)),
                sse_read_timeout=timedelta(seconds=max(timeout + 30.0, 180.0)),
            ) as (read, write, _get_sid):
                async with ClientSession(read, write) as session:
                    await session.initialize()
                    result = await asyncio.wait_for(
                        session.call_tool(tool, params or {}),
                        timeout=timeout,
                    )
        except Exception as exc:
            print(f"    [linkedin_mcp] http {tool} failed: {exc}", flush=True)
            return None

        structured = getattr(result, "structuredContent", None)
        if isinstance(structured, dict) and structured:
            return structured
        texts: list[str] = []
        for block in getattr(result, "content", None) or []:
            text = getattr(block, "text", None)
            if text:
                texts.append(str(text))
        if not texts:
            return None
        return _parse_mcp_payload("\n".join(texts))


async def _mcporter_call(tool_expr: str, *, timeout: float | None = None) -> str:
    if timeout is None:
        timeout = _call_timeout_sec()
    timeout = max(30.0, float(timeout))

    def _kill_tree(pid: int) -> None:
        """Kill mcporter + child node processes (Windows subprocess timeout often leaves them)."""
        if pid <= 0:
            return
        try:
            if sys.platform == "win32":
                subprocess.run(
                    ["taskkill", "/PID", str(pid), "/T", "/F"],
                    capture_output=True,
                    check=False,
                    timeout=15,
                )
            else:
                import signal

                os.killpg(os.getpgid(pid), signal.SIGKILL)
        except Exception:
            try:
                os.kill(pid, 9)
            except Exception:
                pass

    def _run() -> str:
        last_err = ""
        for cmd in _mcporter_commands():
            proc: subprocess.Popen[str] | None = None
            try:
                kwargs: dict[str, Any] = {
                    "stdout": subprocess.PIPE,
                    "stderr": subprocess.PIPE,
                    "text": True,
                    "encoding": "utf-8",
                    "errors": "replace",
                }
                if sys.platform == "win32":
                    kwargs["creationflags"] = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
                else:
                    kwargs["start_new_session"] = True
                proc = subprocess.Popen([cmd, "call", tool_expr], **kwargs)
                try:
                    out, err = proc.communicate(timeout=timeout)
                except subprocess.TimeoutExpired:
                    _kill_tree(proc.pid)
                    try:
                        proc.communicate(timeout=5)
                    except Exception:
                        pass
                    last_err = f"timeout after {timeout:.0f}s"
                    continue
                text = (out or err or "").strip()
                if text:
                    return text
                last_err = err.strip() if err else f"exit {proc.returncode}"
            except FileNotFoundError:
                last_err = f"{cmd} not found"
            except Exception as exc:
                if proc and proc.poll() is None:
                    _kill_tree(proc.pid)
                last_err = str(exc)
        return last_err

    return await asyncio.to_thread(_run)


async def _invoke_tool(tool: str, **params: Any) -> Any:
    if not linkedin_mcp_available():
        return None
    await _wait_linkedin_rate_limit()

    # Prefer direct HTTP — mcporter CLI leaves zombie node processes that
    # serialize behind LinkedIn MCP's single-tool lock and cause TIMEOUT skips.
    prefer_http = _env_true("LINKEDIN_MCP_HTTP", default=True)
    if prefer_http and _linkedin_mcp_port_open():
        data = await _invoke_tool_http(tool, **params)
        if data is not None:
            return data

    prefix = _linkedin_server_prefix()
    if not prefix:
        return None
    arg_str = _format_call_args(**params)
    expr = f"{prefix}.{tool}({arg_str})" if arg_str else f"{prefix}.{tool}()"
    raw = await _mcporter_call(expr)
    if not raw or raw.startswith("exit ") or "not found" in raw.lower():
        return None
    if raw.startswith("timeout after"):
        print(f"    [linkedin_mcp] {raw}", flush=True)
        return None
    return _parse_mcp_payload(raw)


def _parse_mcp_payload(raw: str) -> Any:
    raw = raw.strip()
    if not raw:
        return None
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        pass
    for m in re.finditer(r"\{[\s\S]*\}|\[[\s\S]*\]", raw):
        try:
            return json.loads(m.group(0))
        except json.JSONDecodeError:
            continue
    return {"text": raw}


def _unwrap_items(data: Any) -> list[dict[str, Any]]:
    if data is None:
        return []
    if isinstance(data, list):
        return [x for x in data if isinstance(x, dict)]
    if isinstance(data, dict):
        for key in (
            "results", "people", "employees", "items", "companies",
            "profiles", "data", "content",
        ):
            val = data.get(key)
            if isinstance(val, list):
                return [x for x in val if isinstance(x, dict)]
        if any(k in data for k in ("name", "title", "headline", "url", "linkedin_url")):
            return [data]
    return []


def _linkedin_company_slug(value: str) -> str:
    """Extract LinkedIn company URL slug from a URL, path, or raw slug string."""
    raw = (value or "").strip()
    if not raw:
        return ""
    m = re.search(r"(?:linkedin\.com)?/company/([^/?#\s]+)", raw, re.I)
    if m:
        slug = m.group(1).strip().strip("/")
        if slug.lower() in ("about", "people", "posts", "jobs", "life"):
            return ""
        return slug
    # Already a slug (no spaces / URL / path)
    if "://" not in raw and " " not in raw and "/" not in raw:
        return raw.strip("/")
    return ""


def _slugs_from_mcp_payload(data: Any) -> list[str]:
    """Pull company slugs from search/profile MCP payloads (url, references, text)."""
    found: list[str] = []
    seen: set[str] = set()

    def _add(val: str) -> None:
        slug = _linkedin_company_slug(val)
        if slug and slug.lower() not in seen:
            seen.add(slug.lower())
            found.append(slug)

    def _scan_text(blob: str) -> None:
        if not blob:
            return
        for m in re.finditer(r"linkedin\.com/company/([A-Za-z0-9_-]+)", blob, re.I):
            _add(m.group(0))
        # Bare /company/slug paths in scraped text
        for m in re.finditer(r"/company/([A-Za-z0-9_-]+)", blob, re.I):
            _add(m.group(1))

    if isinstance(data, dict):
        _add(str(data.get("url") or ""))
        refs = data.get("references")
        if isinstance(refs, dict):
            for items in refs.values():
                if not isinstance(items, list):
                    continue
                for item in items:
                    if isinstance(item, dict):
                        _add(str(item.get("value") or item.get("url") or item.get("href") or ""))
                        _add(str(item.get("path") or ""))
                        kind = str(item.get("kind") or "").lower()
                        if "company" in kind or kind in ("link", "url", "href"):
                            _add(str(item.get("value") or ""))
                    else:
                        _add(str(item))
        elif isinstance(refs, list):
            for item in refs:
                _add(str(item))
        sections = data.get("sections")
        if isinstance(sections, dict):
            for v in sections.values():
                _scan_text(str(v) if v else "")
        _scan_text(str(data.get("text") or ""))
    elif isinstance(data, str):
        _scan_text(data)
    elif isinstance(data, list):
        for item in data:
            found.extend(_slugs_from_mcp_payload(item))
    return found


_LI_NAV_NOISE = re.compile(
    r"^(home|about|posts|jobs|people|life|products|follow|message|visit website|"
    r"overview|website|industry|company size|headquarters|founded|specialties|"
    r"locations?|primary|get directions|i.?m interested|see all|show more)$",
    re.I,
)

_GENERIC_SPECIALTY = re.compile(
    r"^(software|technology|manufacturing|services?|consulting|retail|wholesale|"
    r"distribution|industrial|chemical|engineering|construction|healthcare|"
    r"education|finance|marketing|sales|operations|developers?|open source)$",
    re.I,
)

_BRAND_CUE = re.compile(
    r"(?:key brands?|brands? (?:we |they )?(?:represent|distribute|carry|partner(?:s)? with)|"
    r"authorized\s+(?:distributor|reseller|dealer)\s+(?:of|for)|"
    r"exclusive\s+(?:distributor|partner)\s+(?:of|for)|"
    r"distributor\s+of|reseller\s+of|partner(?:ship)?\s+with)\s*[:\-]?\s*([^\n.]{5,200})",
    re.I,
)


def _parse_about_firmographics(text: str) -> dict[str, Any]:
    """Parse founded / HQ / employees / locations / about / specialties from LinkedIn about text."""
    out: dict[str, Any] = {}
    if not (text or "").strip():
        return out
    t = text.replace("\r", "\n")
    lines = [ln.strip() for ln in t.split("\n") if ln.strip()]

    # Label-on-own-line then value on next line (common LinkedIn about layout)
    for i, ln in enumerate(lines):
        low = ln.lower().rstrip(":")
        nxt = lines[i + 1] if i + 1 < len(lines) else ""
        if low in ("founded", "year founded", "established") and nxt:
            m = re.search(r"(19|20)\d{2}", nxt)
            if m and not out.get("founded_year"):
                out["founded_year"] = m.group(0)
        if low in ("headquarters", "hq", "head office") and nxt and len(nxt) >= 3:
            if not out.get("headquarters") and not _LI_NAV_NOISE.match(nxt):
                out["headquarters"] = re.sub(r"\s+", " ", nxt).strip(" .;-")
        if low in ("company size", "employees", "staff") and nxt and not out.get("employee_count"):
            out["employee_count"] = re.sub(r"\s+", " ", nxt).strip(" .;-")
        if low == "specialties" and nxt:
            specs = [s.strip() for s in re.split(r"\s*,\s*", nxt) if s.strip()]
            cleaned_specs: list[str] = []
            for s in specs:
                s = re.sub(r"^(and\s+)", "", s, flags=re.I).strip()
                if len(s) >= 2 and not _GENERIC_SPECIALTY.match(s):
                    cleaned_specs.append(s)
            if cleaned_specs:
                out["specialties"] = "; ".join(cleaned_specs[:12])

    # Structured "Founded: 1969" and prose "since being established in 1969" /
    # "founded in 2007" / "incorporated in 1998" / "est. 1969". Allow short filler
    # words (in / on / back in / during) but nothing else, so distant years in
    # narrative text are not mistaken for the founding year. Include 18xx (e.g. Nokia 1865).
    m = re.search(
        r"(?:\bfounded|\bestablished|\bincorporated|\byear\s+founded|\best\.)"
        r"\s*(?:in|on|back\s+in|during)?\s*[:\-]?\s*((?:18|19|20)\d{2})",
        t,
        re.I,
    )
    if m and not out.get("founded_year"):
        out["founded_year"] = m.group(1)

    m = re.search(
        r"(?:headquarters|hq|head office|based in)\s*[:\-]?\s*([^\n|]{3,80})",
        t,
        re.I,
    )
    if m and not out.get("headquarters"):
        hq = re.sub(r"\s+", " ", m.group(1)).strip(" .;-")
        if hq and not _LI_NAV_NOISE.match(hq):
            out["headquarters"] = hq

    m = re.search(
        r"(?:company size|employees?|staff)\s*[:\-]?\s*([^\n|]{2,40})",
        t,
        re.I,
    )
    if m and not out.get("employee_count"):
        out["employee_count"] = re.sub(r"\s+", " ", m.group(1)).strip(" .;-")

    # Locations block → cities / office list (feeds Excel "Cities / Regions")
    loc_cities: list[str] = []
    in_locations = False
    for i, ln in enumerate(lines):
        low = ln.lower()
        if re.match(r"^locations?\s*(\(\d+\))?$", low):
            in_locations = True
            continue
        if in_locations:
            if _LI_NAV_NOISE.match(ln) or low in ("primary",):
                continue
            if re.match(r"^(website|industry|company size|founded|specialties|overview)\b", low):
                in_locations = False
                continue
            # Street / city lines
            cleaned = re.sub(r"\s+", " ", ln).strip(" .;-")
            cleaned = re.sub(r"^get directions?\s*(to\s+)?", "", cleaned, flags=re.I).strip()
            if len(cleaned) < 4 or cleaned.lower().startswith("get direction"):
                continue
            # Prefer "City, State" / address-looking lines
            if re.search(r"\d", cleaned) or "," in cleaned or re.search(
                r"\b(alabama|california|texas|florida|new york|ontario|england|germany|india|usa|uk|us)\b",
                cleaned,
                re.I,
            ):
                if cleaned not in loc_cities and cleaned.lower() != str(out.get("headquarters") or "").lower():
                    loc_cities.append(cleaned)
            if len(loc_cities) >= 6:
                break
    if loc_cities:
        out["offices"] = "; ".join(loc_cities[:6])
        # Compact city/region labels for Cities / Regions when addresses are long
        compact: list[str] = []
        for loc in loc_cities:
            # "3217 Messer Airport Hwy, Birmingham, Alabama 35222, US" → Birmingham, Alabama
            parts = [p.strip() for p in loc.split(",") if p.strip()]
            if len(parts) >= 3 and re.search(r"\d", parts[0]):
                city_region = ", ".join(parts[1:3]).strip()
                city_region = re.sub(r"\s+\d{4,}\s*", " ", city_region).strip(" ,")
                if city_region:
                    compact.append(city_region)
            elif len(parts) >= 2:
                compact.append(", ".join(parts[:2]))
            else:
                compact.append(loc)
        # unique preserve order
        seen: set[str] = set()
        uniq = []
        for c in compact:
            k = c.lower()
            if k not in seen:
                seen.add(k)
                uniq.append(c)
        if uniq:
            out["cities_regions"] = "; ".join(uniq[:6])

    # Company-written About / Overview (skip chrome / nav)
    overview = ""
    # Prefer text after "Overview" until Website/Industry
    m = re.search(
        r"(?is)\boverview\b\s*(.+?)(?=\n(?:website|industry|company size|headquarters|founded|specialties|locations)\b)",
        t,
    )
    if m:
        overview = re.sub(r"\s+", " ", m.group(1)).strip()
    if not overview or len(overview) < 40:
        # Tagline-ish paragraphs early on the page (company description)
        for p in re.split(r"\n{2,}", t):
            p = re.sub(r"\s+", " ", p).strip()
            if len(p) < 40:
                continue
            if _LI_NAV_NOISE.match(p.split()[0] if p.split() else ""):
                continue
            if re.search(
                r"followers|associated members|skip to|keyboard shortcuts|notifications|try premium",
                p,
                re.I,
            ):
                continue
            if re.search(r"^(home|about|posts|jobs)\b", p, re.I):
                continue
            overview = p[:4000]
            break
    if overview:
        # Drop leading company-name-only residue if duplicated
        out["about_description"] = overview[:4000]

    # Key-brand cues in the about prose ONLY (not LinkedIn Specialties — those are topics)
    brand_bits: list[str] = []
    for m in _BRAND_CUE.finditer(t):
        chunk = m.group(1).strip()
        for part in re.split(r"[,;/]| and ", chunk):
            part = part.strip(" .;-")
            if 2 <= len(part) <= 60 and not _GENERIC_SPECIALTY.match(part):
                brand_bits.append(part)
    # Quoted product / brand names in overview
    about_for_brands = str(out.get("about_description") or "")
    for m in re.finditer(r"[\"“]([A-Z][^\"”]{1,40})[\"”]", about_for_brands):
        brand_bits.append(m.group(1).strip())
    if brand_bits:
        seen_b: set[str] = set()
        brands: list[str] = []
        for b in brand_bits:
            k = b.lower()
            if k not in seen_b:
                seen_b.add(k)
                brands.append(b)
        out["key_brands_from_about"] = "; ".join(brands[:8])

    return out


def _pick_str(d: dict[str, Any], *keys: str) -> str:
    for k in keys:
        v = d.get(k)
        if v is not None and str(v).strip():
            return str(v).strip()
    return ""


def _normalize_person(item: dict[str, Any]) -> dict[str, str]:
    name = _pick_str(item, "name", "full_name", "fullName", "contact_person")
    title = _pick_str(item, "title", "headline", "job_title", "jobTitle", "contact_role")
    url = _pick_str(item, "url", "linkedin_url", "linkedinUrl", "profile_url", "profileUrl")
    if url and not url.startswith("http"):
        url = f"https://www.linkedin.com/in/{url.lstrip('/')}"
    if not name:
        first = _pick_str(item, "first_name", "firstName")
        last = _pick_str(item, "last_name", "lastName")
        name = f"{first} {last}".strip()
    out: dict[str, str] = {}
    if name and len(name) >= 3 and "*" not in name:
        out["contact_person"] = name
    if title:
        out["contact_role"] = title
    if url:
        out["linkedin"] = url
    return out


def _pick_exec_person(people: list[dict[str, Any]]) -> dict[str, str] | None:
    scored: list[tuple[int, dict[str, str]]] = []
    for item in people:
        parsed = _normalize_person(item)
        if not parsed.get("contact_person"):
            continue
        title = parsed.get("contact_role", "")
        score = 2 if _EXEC_TITLE.search(title) else 1
        if re.search(r"\bciso\b", title, re.I):
            score += 3
        if re.search(r"\bceo\b|chief executive|founder", title, re.I):
            score += 2
        scored.append((score, parsed))
    if not scored:
        return None
    scored.sort(key=lambda x: x[0], reverse=True)
    return scored[0][1]


def _normalize_company_profile(data: Any) -> dict[str, Any]:
    """Normalize mcp-server-linkedin company payload into firmographic fields."""
    if not isinstance(data, dict):
        return {}

    out: dict[str, Any] = {}

    # New scraper format: {url, sections: {about: "..."}, references: {...}}
    sections = data.get("sections")
    about_text = ""
    if isinstance(sections, dict):
        about_text = str(sections.get("about") or sections.get("About") or "")
        if not about_text:
            about_text = "\n".join(str(v) for v in sections.values() if v)
    if about_text:
        # Dead / missing LinkedIn pages return chrome chrome, not firmographics
        if re.search(r"page isn.?t available|no longer exists|page not found", about_text, re.I):
            return {}
        out.update(_parse_about_firmographics(about_text))

    url = _pick_str(data, "url", "linkedin_url", "linkedinUrl")
    slug = _linkedin_company_slug(url)
    if slug:
        out["linkedin_url"] = f"https://www.linkedin.com/company/{slug}"
    else:
        for s in _slugs_from_mcp_payload(data):
            out["linkedin_url"] = f"https://www.linkedin.com/company/{s}"
            break

    # Legacy / structured fields if present
    root = data
    for key in ("company", "profile", "about", "data"):
        nested = data.get(key)
        if isinstance(nested, dict) and len(nested) > 2:
            root = {**data, **nested}
            break

    founded = _pick_str(root, "founded_year", "foundedYear", "founded", "year_founded")
    if founded and not out.get("founded_year"):
        m = re.search(r"(19|20)\d{2}", founded)
        out["founded_year"] = m.group(0) if m else founded
    hq = _pick_str(
        root, "headquarters", "hq", "location", "headquarters_location", "headquarter"
    )
    if hq and not out.get("headquarters"):
        out["headquarters"] = hq
    employees = _pick_str(
        root, "employee_count", "employeeCount", "employees", "company_size", "staff_count"
    )
    if employees and not out.get("employee_count"):
        out["employee_count"] = employees
    about = _pick_str(root, "about_description", "about", "description", "overview", "tagline")
    if about and not out.get("about_description"):
        out["about_description"] = about[:4000]
    for key in ("offices", "cities_regions", "specialties", "key_brands_from_about"):
        if out.get(key):
            continue
        val = _pick_str(root, key)
        if val:
            out[key] = val
    return out


async def search_people(
    keyword: str,
    *,
    limit: int = 5,
    current_company: str = "",
) -> list[dict[str, Any]]:
    if not keyword.strip():
        return []
    params: dict[str, Any] = {"keyword": keyword, "limit": max(1, min(limit, 10))}
    if current_company:
        params["current_company"] = current_company
    data = await _invoke_tool("search_people", **params)
    return _unwrap_items(data)


async def search_companies(keyword: str, *, limit: int = 5) -> list[dict[str, Any]]:
    """Search LinkedIn companies. mcp-server-linkedin expects ``keywords`` (not keyword/limit)."""
    if not keyword.strip():
        return []
    # Official tool signature: search_companies(keywords: str)
    data = await _invoke_tool("search_companies", keywords=keyword.strip())
    if data is None:
        return []

    # Prefer company slugs from references (path like /company/accufoamcps/)
    hits: list[dict[str, Any]] = []
    for slug in _slugs_from_mcp_payload(data)[: max(1, min(limit, 10))]:
        hits.append(
            {
                "name": slug.replace("-", " "),
                "url": f"https://www.linkedin.com/company/{slug}",
                "linkedin_url": f"https://www.linkedin.com/company/{slug}",
                "slug": slug,
            }
        )
    if hits:
        return hits

    items = _unwrap_items(data)
    if items:
        return items[: max(1, min(limit, 10))]
    if isinstance(data, dict):
        return [data]
    return []


async def get_company_profile(
    *,
    linkedin_url: str = "",
    company_name: str = "",
) -> dict[str, Any]:
    """Fetch company about page via MCP.

    mcp-server-linkedin ``get_company_profile`` requires the LinkedIn **URL slug**
    (e.g. ``accufoamcps``), not a display name and not ``linkedin_url``.
    """
    slug = _linkedin_company_slug(linkedin_url) or _linkedin_company_slug(company_name)
    if not slug and company_name and " " not in company_name.strip() and "://" not in company_name:
        # Treat bare token as slug
        slug = company_name.strip().strip("/")
    if not slug:
        return {}
    data = await _invoke_tool("get_company_profile", company_name=slug)
    if isinstance(data, dict):
        profile = _normalize_company_profile(data)
        if not profile.get("linkedin_url"):
            profile["linkedin_url"] = f"https://www.linkedin.com/company/{slug}"
        return profile
    return {}


async def get_company_employees(
    linkedin_url: str,
    *,
    keyword: str = "",
    limit: int = 5,
) -> list[dict[str, Any]]:
    slug = _linkedin_company_slug(linkedin_url)
    if not slug:
        return []
    params: dict[str, Any] = {"company_name": slug}
    if keyword:
        # Official arg is keywords
        params["keywords"] = keyword
    data = await _invoke_tool("get_company_employees", **params)
    return _unwrap_items(data)


async def resolve_contact_via_mcp(
    company: str,
    *,
    linkedin_company_url: str = "",
) -> dict[str, str] | None:
    if not linkedin_mcp_available() or not company.strip():
        return None

    co_url = linkedin_company_url.strip()
    if co_url and "linkedin.com/company" in co_url:
        for kw in ("CISO", "Chief Security Officer", "CEO", "Founder", "President"):
            try:
                employees = await get_company_employees(co_url, keyword=kw, limit=8)
                hit = _pick_exec_person(employees)
                if hit:
                    return hit
            except Exception:
                continue
            await asyncio.sleep(linkedin_mcp_delay())

    for kw in (
        f"{company} CISO",
        f"{company} Chief Security Officer",
        f"{company} CEO",
        f"{company} Founder",
    ):
        try:
            people = await search_people(kw, limit=8, current_company=company)
            hit = _pick_exec_person(people)
            if hit:
                return hit
        except Exception:
            continue
        await asyncio.sleep(linkedin_mcp_delay())
    return None


async def enrich_company_via_mcp(
    company_name: str,
    *,
    linkedin_url: str = "",
) -> tuple[dict[str, Any], str]:
    """Search LinkedIn for the company (if needed), open profile, return firmographics.

    Flow:
      1) If we already have a /company/{slug} URL → scrape that profile
      2) Else search_companies(keywords=display name) → pick best slug → scrape profile
    """
    if not linkedin_mcp_available() or not company_name.strip():
        return {}, linkedin_url

    profile: dict[str, Any] = {}
    resolved_url = (linkedin_url or "").strip()
    slug = _linkedin_company_slug(resolved_url)

    if slug:
        print(f"    [linkedin_mcp] profile slug={slug}", flush=True)
        profile = await get_company_profile(company_name=slug)
    else:
        print(f"    [linkedin_mcp] search companies keywords={company_name!r}", flush=True)
        hits = await search_companies(company_name, limit=5)
        name_l = company_name.lower().strip()
        chosen_slug = ""
        for hit in hits:
            hit_slug = (
                _linkedin_company_slug(_pick_str(hit, "url", "linkedin_url", "linkedinUrl"))
                or str(hit.get("slug") or "").strip()
            )
            hit_name = _pick_str(hit, "name", "company_name", "title").lower()
            if hit_slug and (not hit_name or name_l in hit_name or hit_name in name_l):
                chosen_slug = hit_slug
                break
        if not chosen_slug and hits:
            chosen_slug = (
                _linkedin_company_slug(_pick_str(hits[0], "url", "linkedin_url"))
                or str(hits[0].get("slug") or "").strip()
            )
        # Also accept slugs embedded in raw search payload
        if not chosen_slug:
            for s in _slugs_from_mcp_payload(hits[0] if hits else {}):
                chosen_slug = s
                break

        if chosen_slug:
            print(f"    [linkedin_mcp] search hit slug={chosen_slug} -> profile", flush=True)
            profile = await get_company_profile(company_name=chosen_slug)
            resolved_url = f"https://www.linkedin.com/company/{chosen_slug}"
        else:
            # Last resort: try display name as slug (works for Accufoam-like names)
            guess = re.sub(r"[^A-Za-z0-9]+", "", company_name).lower()
            if guess:
                print(f"    [linkedin_mcp] fallback slug guess={guess}", flush=True)
                profile = await get_company_profile(company_name=guess)

    if profile.get("linkedin_url"):
        resolved_url = str(profile["linkedin_url"])
    return profile, resolved_url
