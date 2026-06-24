# -*- coding: utf-8 -*-
"""Merge INDEX<TAB>ENGLISH part files into translations.json.

Reads every _translation/parts/*.tsv, parses 'INDEX<TAB>ENGLISH', un-escapes
\\n \\t \\\\, and maps to the exact original key via keys.json[INDEX].
Reports any indices still missing so coverage is visible.
"""
import os, json, glob

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
T = os.path.join(ROOT, "_translation")

keys = json.load(open(os.path.join(T, "keys.json"), encoding="utf-8"))

def undisp(s):
    # Values are kept verbatim. The source represents in-string line breaks as
    # the literal 2-char escape sequence \n (backslash + n) inside single-line
    # string literals, and our TSV English mirrors that exactly. So no
    # transformation: real newlines would corrupt the Python/XAML literals.
    return s

trans = {}
seen = set()
for fp in sorted(glob.glob(os.path.join(T, "parts", "*.tsv"))):
    for line in open(fp, encoding="utf-8"):
        line = line.rstrip("\n").rstrip("\r")
        if not line or "\t" not in line:
            continue
        idx, eng = line.split("\t", 1)
        try:
            idx = int(idx)
        except ValueError:
            continue
        if idx < 0 or idx >= len(keys):
            continue
        trans[keys[idx]] = undisp(eng)
        seen.add(idx)

json.dump(trans, open(os.path.join(T, "translations.json"), "w", encoding="utf-8"),
          ensure_ascii=False, indent=1, sort_keys=True)

missing = [i for i in range(len(keys)) if i not in seen]
print("translated:", len(seen), "/", len(keys))
print("missing indices:", len(missing))
if missing:
    print("first missing:", missing[:20])
