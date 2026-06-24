# -*- coding: utf-8 -*-
"""Create an ordered, indexed worklist from strings.json.

- keys.json : ordered list of the exact original strings (source of truth)
- worklist.tsv : INDEX<TAB>GERMAN(display-safe) for human/AI translation
Real newlines/tabs in a string are shown as \\n / \\t so each entry is one line.
The translator returns INDEX<TAB>ENGLISH lines; merge.py maps by INDEX.
"""
import os, json

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
T = os.path.join(ROOT, "_translation")

strings = json.load(open(os.path.join(T, "strings.json"), encoding="utf-8"))
keys = sorted(strings.keys())
json.dump(keys, open(os.path.join(T, "keys.json"), "w", encoding="utf-8"),
          ensure_ascii=False, indent=0)

def disp(s):
    return s.replace("\\", "\\\\").replace("\n", "\\n").replace("\t", "\\t")

with open(os.path.join(T, "worklist.tsv"), "w", encoding="utf-8", newline="") as f:
    for i, k in enumerate(keys):
        f.write("{}\t{}\n".format(i, disp(k)))

print("keys:", len(keys))
print("wrote keys.json and worklist.tsv")
