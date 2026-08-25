# -*- coding: utf-8 -*-
import json
import sys
from pathlib import Path

mkt = sys.argv[1]
rows = json.loads(Path(r"D:\vendor-intel\output\_npd_linkedin.json").read_text(encoding="utf-8"))
i = 0
for x in rows:
    if x["m"] != mkt:
        continue
    if "*" in x["person"]:
        continue
    print(x["person"], "|", x["role"][:38], "|", x["company"][:42])
    i += 1
    if i >= 35:
        break
print("printed", i, "of", mkt)
