# Chrome scraper & captcha extensions

Local Chrome extensions used with vendor-intel / Google AI Overview scraping.

## Contents

| Folder | Purpose | Load unpacked path |
|--------|---------|-------------------|
| `google-ai-scraper/extension` | Google AI Overview scraper | Load this folder |
| `captcha-raptor/extension` | Free captcha helper | Load this folder |
| `hektcaptcha/public` | hektCaptcha | Load this folder |
| `lets-solve-captcha/src` | Let's Solve Captcha | Load this folder |
| `nopecha-extension-main` | NopeCHA docs only (use Chrome Web Store for real extension) | N/A |

## Google AI scraper backend

```powershell
cd google-ai-scraper/server
python -m venv .venv
.\.venv\Scripts\pip install -e .
.\.venv\Scripts\google-ai-scraper.exe --backend --port 15551
```

Point the extension options at `http://localhost:15551`.

## Load in Chrome

1. Open `chrome://extensions`
2. Enable **Developer mode**
3. **Load unpacked** → select the folder listed above
