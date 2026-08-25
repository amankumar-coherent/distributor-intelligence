import json
from pathlib import Path

roots = [
    Path(r"C:\Users\aman.kumar\AppData\Local\Google\Chrome\User Data"),
    Path(r"C:\Users\aman.kumar\AppData\Local\Microsoft\Edge\User Data"),
]
for root in roots:
    if not root.exists():
        continue
    print("===", root)
    for pref in list(root.glob("*/Preferences")) + list(root.glob("*/Secure Preferences")):
        try:
            data = json.loads(pref.read_text(encoding="utf-8"))
        except Exception as e:
            print("FAIL", pref, type(e).__name__, e)
            continue
        ext = data.get("extensions") or {}
        settings = ext.get("settings") or {}
        print(pref.parent.name, pref.name, "settings", len(settings), "ext_keys", list(ext)[:15])
        for eid, meta in settings.items():
            if not isinstance(meta, dict):
                continue
            manifest = meta.get("manifest") or {}
            name = manifest.get("name") if isinstance(manifest, dict) else ""
            path = str(meta.get("path") or "")
            blob = f"{name} {path} {eid}".lower()
            if any(k in blob for k in ("ai", "scraper", "overview", "google-ai", "15551")):
                print(
                    " HIT",
                    pref.parent.name,
                    eid,
                    "state=",
                    meta.get("state"),
                    "name=",
                    name,
                    "path=",
                    path[:120],
                )
        # also print all names
        for eid, meta in settings.items():
            if not isinstance(meta, dict):
                continue
            manifest = meta.get("manifest") or {}
            name = manifest.get("name") if isinstance(manifest, dict) else ""
            if name:
                print("  ALL", pref.parent.name, eid[:16], repr(name)[:80], "from=", meta.get("from_webstore"))
