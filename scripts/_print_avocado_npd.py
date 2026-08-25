# -*- coding: utf-8 -*-
import json
from pathlib import Path

rows = json.loads(Path(r"D:\vendor-intel\output\_npd_linkedin.json").read_text(encoding="utf-8"))
i = 0
for x in rows:
    if x["m"] != "avocado":
        continue
    if "*" in x["person"]:
        continue
    print(x["person"], "|", x["role"][:40], "|", x["company"][:45])
    i += 1
    if i >= 40:
        break
print("printed", i)
