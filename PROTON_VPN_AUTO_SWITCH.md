# Automatic Proton VPN server switching for browser automation (Windows)

This guide explains how to make a browser-automation project switch its
Proton VPN server automatically. It switches on a fixed schedule, and also
whenever Google (or any site) starts blocking. It needs **no admin rights**:
the code drives the **Proton VPN Windows app itself** through Windows UI
Automation.

Every switch runs these steps in order:

1. **Pause.** All automation browsers finish their current request and close.
2. **Disconnect** the current Proton server (the code clicks the button in the Proton app).
3. **Connect** to the next country in a rotation list.
4. **Check** that the public IP really changed and the internet works.
5. **Resume.** The browsers reopen with **fresh profiles** (no old cookies) on the new IP.

It was built and tested in the `distributor-intelligence` project, which runs
10–20 Patchright/Chromium browsers that query Google AI Mode. The switch takes
about 15 seconds, plus the time the browsers need to finish their current request.

---

## 1. Requirements

| What | Detail |
|---|---|
| OS | Windows 10/11 |
| VPN | **Proton VPN app** (tested on v5.1.9) with a paid plan (Plus) that can pick countries |
| Python | 3.10+ with `pywinauto` (`pip install pywinauto`) |
| Admin rights | **Not needed** |
| Browsers | Any automation that you can tell to "close and wait" (Playwright / Patchright / Selenium) |

### Why the app is driven instead of the command line

- The Proton VPN Windows app has **no official CLI**.
- The cleaner alternative, switching between WireGuard config files with
  `wireguard.exe /installtunnelservice`, needs **admin rights**. On a
  company-managed laptop you usually don't have them.
- The Proton app exposes stable **UI Automation ids**, so pywinauto can click
  its buttons reliably.

---

## 2. Proton VPN app settings (important)

1. **Kill switch: OFF.** With it on, a connect that fails cuts the **whole
   internet**: the Wi-Fi looks disconnected and there's no internet. This was
   the first problem we hit. Proton main screen → **Kill switch** → Off.
2. **Protocol:** Smart, or whatever connects on your network. In the test,
   WireGuard (UDP) worked once the Kill switch was off. If nothing connects,
   try **Stealth**, then WireGuard (TCP), then OpenVPN (TCP).
3. **Check manually first:** click Connect in the app yourself and make sure it
   reaches **"Protected"** with a new IP. If a manual connect fails, the code
   can't fix that: it's the network or the company's security software.
4. Keep the Proton app **open (or minimised)** and signed in. The code starts
   it if it isn't running.

---

## 3. Proton app UI Automation ids (v5.1.9)

These are the hooks the code uses. They were found by listing the window's
elements with pywinauto (see §9 if a Proton update changes them).

| Element | Automation id | Notes |
|---|---|---|
| Window | title `Proton VPN` | |
| Status text | `ProtectionTitle` | `"Protected"` when connected, `"Unprotected"` when off; other text while connecting |
| Connect button (when off) | `ConnectionCardConnectButton` | text `"Connect"` |
| Disconnect button (when on) | `ConnectionCardDisconnectButton` | **different id** from the Connect button |
| Country search box | `SearchTextBox` | type a country name to bring its row into view |
| Connect to a country | `Connect_to_<code>` | e.g. `Connect_to_US`, `Connect_to_DE`. **United Kingdom is `UK`, not `GB`** |
| Current IP / country (when off) | `PART_ScrambledText` | the home card shows IP, country, provider |

**Two traps we hit:**
- While connecting, the status is neither "Protected" nor "Unprotected". Only
  **`ProtectionTitle == "Protected"`** means connected. Counting "connecting"
  as connected leads to "connected, but the IP didn't change".
- After connecting, **`ConnectionCardConnectButton` no longer exists**. Use
  `ConnectionCardDisconnectButton`.

---

## 4. Architecture

```
 automation browsers (N workers)              VPN guard (1 process for the whole PC)
 ───────────────────────────────              ─────────────────────────────────────
 before every request:                        every 30 s:
   1. if data/vpn/PAUSE exists:                 score = weighted blocks in last 10 min
        close my browser, wait until gone      if 20 min since last switch  -> switch()
   2. if VPN switch # changed since my         elif score >= limit and 15 min since
      profile was last used:                        last switch              -> switch()
        delete my profile, open fresh
 on a block (sign-in wall, CAPTCHA,          switch():
 verify page, no answer):                       create PAUSE, wait for browsers to close
   append a line to data/vpn/blocks.log        click Disconnect -> wait "Unprotected"
                                                search country, click Connect_to_XX
                                                wait "Protected", check public IP changed
                                                switches += 1 (state.json), delete PAUSE
```

Shared files, all in `data/vpn/`:

| File | Written by | Purpose |
|---|---|---|
| `PAUSE` | guard | Exists while a switch is in progress. Browsers close and wait. |
| `blocks.log` | browsers | One line per block: `<epoch> <pid> <kind>` |
| `state.json` | guard | `{"next": rotation index, "switches": N, "country", "ip", "last_switch"}` |

Every browser profile folder contains a `.vpn_gen` file holding the switch
number it last ran under. When `state.json` shows a newer number, that
browser deletes its profile before its next request.

---

## 5. File 1: `scripts/vpn_switch.py` (drives Proton)

```python
#!/usr/bin/env python3
"""Proton VPN server switching by driving the Proton VPN Windows app (UI
Automation, no admin). switch(): pause browsers -> Disconnect -> connect the
next country -> verify new public IP -> resume.

Usage:
    python scripts/vpn_switch.py --status
    python scripts/vpn_switch.py --switch [--country US]
    python scripts/vpn_switch.py --connect
"""
from __future__ import annotations

import argparse, json, os, subprocess, time, urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
VPN_DIR = ROOT / "data" / "vpn"
PAUSE_FILE = VPN_DIR / "PAUSE"
STATE_FILE = VPN_DIR / "state.json"
LOG_FILE = ROOT / "logs" / "vpn_switch.log"
LAUNCHER = r"C:\Program Files\Proton\VPN\ProtonVPN.Launcher.exe"

# Countries with many Proton servers. Proton codes: UK (not GB).
ROTATION = (("US", "United States"), ("NL", "Netherlands"), ("DE", "Germany"),
            ("UK", "United Kingdom"), ("CA", "Canada"), ("SE", "Sweden"),
            ("CH", "Switzerland"), ("FR", "France"), ("SG", "Singapore"),
            ("JP", "Japan"), ("AU", "Australia"), ("ES", "Spain"),
            ("IT", "Italy"), ("BE", "Belgium"), ("AT", "Austria"),
            ("DK", "Denmark"), ("NO", "Norway"), ("PL", "Poland"))
PAUSE_WAIT_SEC = int(os.environ.get("VPN_PAUSE_WAIT_SEC") or 180)
# Regex matching the command lines of YOUR automation Chrome processes
# (their --user-data-dir). Never match your own everyday browser.
BROWSER_PATTERN = os.environ.get("VPN_BROWSER_PATTERN") or r"_profile_|ai_mode_pverify|ai_mode_status"


def log(msg):
    line = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    print(line, flush=True)
    LOG_FILE.parent.mkdir(exist_ok=True)
    with LOG_FILE.open("a", encoding="utf-8") as fh:
        fh.write(line + "\n")


def public_ip(timeout=10.0):
    for url in ("https://api.ipify.org", "https://ifconfig.me/ip"):
        try:
            with urllib.request.urlopen(url, timeout=timeout) as r:
                return r.read().decode().strip()
        except Exception:
            continue
    return ""


def _window():
    from pywinauto import Desktop
    desk = Desktop(backend="uia")
    try:
        w = desk.window(title="Proton VPN"); w.wait("exists", timeout=5); return w
    except Exception:
        subprocess.Popen([LAUNCHER])
        w = desk.window(title="Proton VPN"); w.wait("exists", timeout=60); return w


def _by_id(w, aid):
    return w.child_window(auto_id=aid)


def status():
    """Connected ONLY when ProtectionTitle == 'Protected' (not while connecting)."""
    try:
        title = _by_id(_window(), "ProtectionTitle").window_text().strip()
    except Exception:
        title = ""
    return {"protection": title, "connected": title.lower() == "protected"}


def _wait(cond, timeout, what):
    end = time.time() + timeout
    while time.time() < end:
        try:
            if cond():
                return True
        except Exception:
            pass
        time.sleep(2)
    log(f"timeout waiting for: {what}")
    return False


def disconnect():
    if status()["protection"].lower() == "unprotected":
        return True
    log("clicking Disconnect")
    _by_id(_window(), "ConnectionCardDisconnectButton").invoke()
    return _wait(lambda: status()["protection"].lower() == "unprotected", 60, "Unprotected")


def connect(code, name):
    w = _window()
    box = _by_id(w, "SearchTextBox")
    box.set_focus(); box.set_edit_text(name); time.sleep(2)   # bring the country row into view
    btn = _by_id(w, f"Connect_to_{code}")
    if not _wait(lambda: btn.exists(), 15, f"Connect_to_{code}"):
        return False
    log(f"connecting to {name} ({code})")
    btn.invoke()
    ok = _wait(lambda: status()["connected"], 120, f"connected to {name}")
    try:
        box.set_edit_text("")
    except Exception:
        pass
    return ok


def _load_state():
    try:
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {"next": 0, "switches": 0}


def _save_state(st):
    VPN_DIR.mkdir(parents=True, exist_ok=True)
    tmp = STATE_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(st, indent=1), encoding="utf-8")
    os.replace(tmp, STATE_FILE)


def _ps(cmd, timeout=60):
    return subprocess.run(["powershell", "-NoProfile", "-Command", cmd],
                          capture_output=True, text=True, timeout=timeout).stdout.strip()


def _pipeline_browsers():
    try:
        return int(_ps("(Get-CimInstance Win32_Process -Filter \"Name='chrome.exe'\" | "
                       f"Where-Object {{ $_.CommandLine -match '{BROWSER_PATTERN}' }}).Count") or 0)
    except Exception:
        return -1


def _close_pipeline_browsers():
    _ps("Get-CimInstance Win32_Process -Filter \"Name='chrome.exe'\" | Where-Object { "
        f"$_.CommandLine -match '{BROWSER_PATTERN}' }} | ForEach-Object {{ "
        "Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }", timeout=120)


def pause(reason):
    VPN_DIR.mkdir(parents=True, exist_ok=True)
    PAUSE_FILE.write_text(f"{time.strftime('%H:%M:%S')} {reason}", encoding="utf-8")
    log(f"PAUSE set ({reason}) - waiting for automation browsers to close")
    end = time.time() + PAUSE_WAIT_SEC
    while time.time() < end:
        if _pipeline_browsers() == 0:
            log("all automation browsers closed")
            return
        time.sleep(5)
    log(f"browsers still open after {PAUSE_WAIT_SEC}s - closing them")
    _close_pipeline_browsers(); time.sleep(5)


def resume():
    try:
        PAUSE_FILE.unlink()
    except FileNotFoundError:
        pass
    log("PAUSE removed - browsers reopen and the run resumes")


def switch(reason="manual", country=None):
    """Pause -> Disconnect -> next country -> verify IP -> resume.
    The pause is ALWAYS lifted (finally), so a UI hiccup never freezes the run."""
    st = _load_state(); old_ip = public_ip(); ok = False
    pause(reason)
    try:
        disconnect() or log("disconnect did not complete")
        for _ in range(3):
            if country:
                code, name = next(((c, n) for c, n in ROTATION if c == country.upper()),
                                  (country.upper(), country))
            else:
                code, name = ROTATION[st["next"] % len(ROTATION)]
                st["next"] = (st["next"] + 1) % len(ROTATION)
            if connect(code, name):
                new_ip, end = "", time.time() + 60      # tunnel needs a moment to carry traffic
                while time.time() < end:
                    new_ip = public_ip(timeout=8)
                    if new_ip and new_ip != old_ip:
                        break
                    time.sleep(4)
                if new_ip and new_ip != old_ip:
                    st.update(switches=st.get("switches", 0) + 1, last_switch=time.time(),
                              country=name, ip=new_ip)
                    log(f"SWITCHED: {old_ip or '?'} -> {new_ip} ({name}), switch #{st['switches']}")
                    ok = True
                    break
                log(f"connected to {name} but IP check failed (old {old_ip}, new {new_ip or 'none'})")
            country = None
        _save_state(st)
    except Exception as exc:
        log(f"switch error: {type(exc).__name__}: {exc}")
    finally:
        resume()
    return ok


def ensure_connected():
    if status()["connected"]:
        return True
    st = _load_state(); code, name = ROTATION[st["next"] % len(ROTATION)]
    st["next"] = (st["next"] + 1) % len(ROTATION); _save_state(st)
    return connect(code, name)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--status", action="store_true"); ap.add_argument("--switch", action="store_true")
    ap.add_argument("--connect", action="store_true"); ap.add_argument("--country")
    a = ap.parse_args()
    if a.status:
        print(status(), "ip:", public_ip())
    elif a.switch:
        print("ok" if switch("manual", a.country) else "FAILED")
    elif a.connect:
        print("ok" if ensure_connected() else "FAILED", public_ip())
```

---

## 6. File 2: `scripts/vpn_guard.py` (decides when to switch)

```python
#!/usr/bin/env python3
"""One guard for the whole PC. Switches the Proton server
  - every ROTATE_MIN minutes, whether or not there are blocks, and
  - earlier when the weighted block score in the last WINDOW_MIN minutes
    reaches SCORE_LIMIT (and COOLDOWN_MIN has passed since the last switch).
Blocks come from data/vpn/blocks.log ("<epoch> <pid> <kind>").
Env: VPN_SCORE_LIMIT 30, VPN_WINDOW_MIN 10, VPN_COOLDOWN_MIN 15,
     VPN_ROTATE_MIN 20 (0 = off), VPN_CHECK_SEC 30."""
from __future__ import annotations
import os, sys, time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import vpn_switch as V  # noqa: E402

WEIGHT = {"signin": 3.0, "verify": 2.0, "no_answer": 1.0, "captcha": 0.5}
SCORE_LIMIT = float(os.environ.get("VPN_SCORE_LIMIT") or 30)
WINDOW_MIN = float(os.environ.get("VPN_WINDOW_MIN") or 10)
COOLDOWN_MIN = float(os.environ.get("VPN_COOLDOWN_MIN") or 15)
ROTATE_MIN = float(os.environ.get("VPN_ROTATE_MIN") or 20)
CHECK_SEC = float(os.environ.get("VPN_CHECK_SEC") or 30)
BLOCKS = V.VPN_DIR / "blocks.log"


def score(since):
    counts, total = {}, 0.0
    try:
        lines = BLOCKS.read_text(encoding="utf-8", errors="replace").splitlines()[-5000:]
    except FileNotFoundError:
        return 0.0, counts
    for line in lines:
        p = line.split()
        if len(p) < 3:
            continue
        try:
            t = float(p[0])
        except ValueError:
            continue
        if t >= since:
            counts[p[2]] = counts.get(p[2], 0) + 1
            total += WEIGHT.get(p[2], 1.0)
    return total, counts


def _do_switch(reason):
    try:
        V.switch(reason)
    except Exception as exc:
        V.log(f"switch failed: {type(exc).__name__}: {exc}")
        if V.PAUSE_FILE.exists():
            V.resume()


def main():
    V.log(f"vpn_guard started: score>={SCORE_LIMIT}/{WINDOW_MIN}min (cooldown {COOLDOWN_MIN}) "
          f"and every {ROTATE_MIN:.0f} min")
    if V.PAUSE_FILE.exists():            # a crashed earlier switch must not freeze the run
        V.resume()
    try:
        if not V.status()["connected"]:
            V.log("VPN not connected - connecting first")
            _do_switch("guard start: VPN was not connected")
    except Exception as exc:
        V.log(f"start check failed: {exc}")
    last_switch = time.time()
    while True:
        time.sleep(CHECK_SEC)
        now = time.time()
        s, counts = score(max(now - WINDOW_MIN * 60, last_switch))
        if ROTATE_MIN and now - last_switch >= ROTATE_MIN * 60:
            V.log(f"scheduled switch (score {s:.1f} {counts})")
            _do_switch(f"scheduled every {ROTATE_MIN:.0f} min"); last_switch = time.time()
        elif s >= SCORE_LIMIT and now - last_switch >= COOLDOWN_MIN * 60:
            V.log(f"block score {s:.1f} {counts} - switching")
            _do_switch(f"blocks: score {s:.1f}"); last_switch = time.time()


if __name__ == "__main__":
    main()
```

---

## 7. File 3: hooks in YOUR browser code

Add these three things to whatever function sends one request through a
browser. In this project that's `google_ai_mode.ask()`. Adapt `shutdown()`,
`rebuild_profile()` and `active_profile_dir()` to your own browser wrapper.

```python
import os, time, json
from pathlib import Path

VPN_DIR = Path(r"<project root>") / "data" / "vpn"
VPN_PAUSE = VPN_DIR / "PAUSE"
VPN_BLOCKS = VPN_DIR / "blocks.log"


def report_block(kind: str) -> None:
    """Call on every block: 'signin', 'verify', 'captcha', 'no_answer'."""
    try:
        VPN_DIR.mkdir(parents=True, exist_ok=True)
        with VPN_BLOCKS.open("a", encoding="utf-8") as fh:
            fh.write(f"{time.time():.0f} {os.getpid()} {kind}\n")
    except Exception:
        pass


def wait_while_vpn_switch() -> None:
    """Step 1 of a switch: close MY browser and wait until the switch is done."""
    if not VPN_PAUSE.exists():
        return
    shutdown()                       # close this worker's browser
    while VPN_PAUSE.exists():
        time.sleep(5)


def fresh_profile_after_vpn_switch() -> None:
    """After a switch, delete this browser's profile (cookies, history) once,
    so the site sees a brand-new visitor on the new IP. Not a failure reset."""
    try:
        gen = int(json.loads((VPN_DIR / "state.json").read_text())["switches"])
    except Exception:
        return
    profile = Path(active_profile_dir())
    marker = profile / ".vpn_gen"
    try:
        seen = int(marker.read_text().strip())
    except Exception:
        seen = None
    if seen != gen and (seen is not None or (profile / "Default").exists()):
        shutdown()
        rebuild_profile(profile)     # delete + recreate the user-data dir (keep extensions if any)
    profile.mkdir(parents=True, exist_ok=True)
    marker.write_text(str(gen))


def ask(prompt):
    wait_while_vpn_switch()          # <- before every request
    fresh_profile_after_vpn_switch() # <- before every request
    ...                              # your normal request code
    # inside it, where you detect a block:
    #   if signin_wall:  report_block("signin")
    #   if captcha:      report_block("captcha")
    #   if verify_page:  report_block("verify")
    #   if no answer:    report_block("no_answer")
```

What counts as a Google block (text found on the page):

- sign-in wall: `"sign in to continue"`, `"sign in to verify you"`
- verify page: `"verifying your request"`
- CAPTCHA: the `/sorry/` page
- no answer: the request returned nothing usable (for example the canned
  "Here are top web results for exploring this topic" page)

---

## 8. Running it

```powershell
pip install pywinauto

# one-off checks
python scripts\vpn_switch.py --status            # {'protection': 'Protected', 'connected': True} ip: ...
python scripts\vpn_switch.py --switch            # full switch with pause/resume
python scripts\vpn_switch.py --switch --country DE

# guard for the whole run (start it next to your browser workers)
python scripts\vpn_guard.py
```

Launching it hidden in the background (as in this project), via a `.ps1`
started through WMI:

```powershell
# logs\run_vpn_guard.ps1
Set-Location "C:\path\to\project"
$env:PYTHONIOENCODING = "utf-8"
& .venv\Scripts\python.exe -u scripts\vpn_guard.py 2>&1 | Out-File -Encoding utf8 -Append "logs\vpn_guard.log"
```
```powershell
Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{
  CommandLine = "conhost.exe --headless powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\path\to\project\logs\run_vpn_guard.ps1"
  CurrentDirectory = "C:\path\to\project" }
```

The guard must run in the **same Windows desktop session** as the Proton app,
because UI Automation can't see other sessions. A process started this way
from your own logged-in user is fine. A Windows service or a scheduled task
running "whether user is logged on or not" is **not**.

**Logs:** `logs\vpn_switch.log` shows every pause, disconnect, connect and
`SWITCHED: old_ip -> new_ip (Country), switch #N` line.

---

## 9. Troubleshooting

| Symptom | Cause / fix |
|---|---|
| Wi-Fi seems to disconnect / no internet when connecting | Proton **Kill switch ON** + the connect failed → turn the Kill switch **off** |
| Manual Connect in the app also fails | The network or security software blocks VPNs → try **Stealth** protocol, try a phone hotspot, ask IT |
| `connected but IP check failed (new = old)` | A "connecting" state was counted as connected → only `ProtectionTitle == "Protected"` counts (already in the code) |
| `timeout waiting for: Connect_to_GB` | Proton uses **UK** for the United Kingdom |
| `ElementNotFoundError ConnectionCardConnectButton` while connected | When connected the button is `ConnectionCardDisconnectButton` |
| Pipeline frozen, browsers don't come back | A stale `data\vpn\PAUSE`: delete it. The guard deletes a stale one at start, and `switch()` always removes it in `finally` |
| A Proton update broke clicking | Re-list the ids (below) and update §3 |
| Switch waited 180 s for browsers | Some browsers were mid-request or stuck in a CAPTCHA; after `VPN_PAUSE_WAIT_SEC` they are closed by force (they reopen on their next request) |

To list the Proton window's ids after an app update:

```python
from pywinauto import Desktop
w = Desktop(backend="uia").window(title="Proton VPN")
for el in w.descendants():
    ct = el.element_info.control_type
    if ct in ("Button", "Edit", "Text", "ListItem"):
        print(ct, "|", el.window_text()[:50], "|", el.element_info.automation_id)
```

---

## 10. Results and lessons from the real run (be realistic)

Test switch: Sweden 159.26.108.145 became Switzerland 212.104.214.56 in
**14 seconds**. Over an evening of 20 browsers on Google AI Mode the guard
made 10+ switches (CH, FR, SG, JP, AU, IT, AT, DK, PL ...) with no manual help.

| Period | Sign-in walls | CAPTCHAs | No answer |
|---|---|---|---|
| Before VPN, home IP, 2 h | 47 | 40 | 12 |
| First 17 min through VPN | **0** | 20 | **0** |
| Later, 20 browsers through VPN, ~2 h | 300 | 567 | 182 |

What we learned:

1. **A new IP helps at first, but Google re-blocks in minutes when the traffic is heavy.**
   With 20 browsers firing queries from one IP, each new Proton IP got sign-in
   walls within 2–12 minutes.
2. **Fresh profiles alone didn't stop it.** A 10-minute test with 46 freshly
   wiped profiles still saw 67 sign-in walls. The main driver is **requests
   per IP**, plus Proton's IP ranges being known to Google, not cookies.
3. **What actually reduces blocks:** fewer parallel browsers per IP (for example
   10 instead of 20), switching on a schedule (every 20 min), and a fresh
   profile after every switch. The run still made steady progress through the
   blocks thanks to the automatic switching and retries.
4. **Always lift the pause in `finally`**, and have the guard delete a stale
   `PAUSE` on start. A stuck pause file freezes every worker.
5. **One guard per PC.** Never let each worker switch the VPN itself: the VPN
   is shared by the whole machine, and N workers would fight over it.
6. A VPN switch drops **every** connection on the PC for a few seconds, including
   other programs. Keep that in mind on a shared or work machine.

---

## 11. Checklist to port this to another project

- [ ] Proton VPN app installed, signed in, **Kill switch off**, a manual connect works
- [ ] `pip install pywinauto`
- [ ] Copy `scripts/vpn_switch.py` and `scripts/vpn_guard.py` (§5, §6)
- [ ] Set `VPN_BROWSER_PATTERN` to match only your automation browsers' `--user-data-dir`
- [ ] Add the three hooks to your request function (§7): `wait_while_vpn_switch()`, `fresh_profile_after_vpn_switch()`, `report_block(kind)`
- [ ] Test: `python scripts/vpn_switch.py --status`, then `--switch`, and check `logs/vpn_switch.log` shows `SWITCHED`
- [ ] Start the guard next to your workers; tune `VPN_ROTATE_MIN`, `VPN_SCORE_LIMIT` and the number of parallel browsers
