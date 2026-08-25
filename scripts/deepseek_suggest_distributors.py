"""Ask DeepSeek for known packaging DISTRIBUTOR suggestions (verify before add)."""
import json
import os
from pathlib import Path

from dotenv import load_dotenv
from openai import OpenAI

load_dotenv(r"D:\vendor-intel\.env")
client = OpenAI(
    api_key=os.environ["OPENAI_API_KEY"],
    base_url=os.environ.get("OPENAI_BASE_URL", "https://api.deepseek.com/v1"),
)
prompt = """List 120 real packaging DISTRIBUTORS / wholesalers / merchants (NOT manufacturers/OEMs/converters) worldwide.
Return JSON: {"companies":[{"company":"","website":"https://...","country":"US","note":"distributor evidence"}]}
Only include companies you are confident are primarily distributors/wholesalers of packaging supplies.
Prefer independent packaging distributors, catalog merchants, industrial packaging wholesalers.
Do NOT invent websites. If unsure of exact domain, omit the company.
Avoid Amcor, Berry, Sealed Air, Smurfit, WestRock, DS Smith, Sonoco, Pregis, Ranpak (manufacturers).
Country is ISO2 code.
"""
r = client.chat.completions.create(
    model=os.environ.get("OPENAI_MODEL", "deepseek-chat"),
    messages=[{"role": "user", "content": prompt}],
    temperature=0.2,
    response_format={"type": "json_object"},
)
text = r.choices[0].message.content or "{}"
out = Path(
    r"D:\vendor-intel\output\web_expand\global_flexible_packaging_market_global"
    r"\deepseek_distributor_suggestions.json"
)
out.write_text(text, encoding="utf-8")
data = json.loads(text)
print("suggested", len(data.get("companies", [])))
for c in data.get("companies", [])[:25]:
    print("-", c.get("company"), c.get("website"), c.get("country"))
