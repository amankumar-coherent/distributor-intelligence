# -*- coding: utf-8 -*-
import json
import sys
from pathlib import Path

mkt = sys.argv[1]
skip = int(sys.argv[2]) if len(sys.argv) > 2 else 0
rows = json.loads(Path(r"D:\vendor-intel\output\_npd_linkedin.json").read_text(encoding="utf-8"))
clean = [x for x in rows if x["m"] == mkt and "*" not in x["person"]]
print("clean", len(clean), mkt)
for x in clean[skip : skip + 30]:
    print(x["person"], "|", x["role"][:36], "|", x["company"][:40])
