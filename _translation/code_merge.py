# -*- coding: utf-8 -*-
"""Merge code-string translation parts into translations_code.json.

Parts are 'INDEX<TAB>ENGLISH' lines in _translation/code_parts/*.tsv. ENGLISH
uses \\n \\t \\\\ escapes; these are DECODED here (to real newline/tab/backslash)
so the map holds decoded values -- apply_german.py re-escapes them per literal.
Keys come from code_keys.json[INDEX] (the exact decoded German source string).
"""
import os, json, glob

T = os.path.dirname(os.path.abspath(__file__))
keys = json.load(open(os.path.join(T, "code_keys.json"), encoding="utf-8"))

def undisp(s):
    out, i = [], 0
    while i < len(s):
        c = s[i]
        if c == "\\" and i + 1 < len(s):
            n = s[i + 1]
            if n == "n": out.append("\n"); i += 2; continue
            if n == "t": out.append("\t"); i += 2; continue
            if n == "\\": out.append("\\"); i += 2; continue
        out.append(c); i += 1
    return "".join(out)

trans, seen = {}, set()
for fp in sorted(glob.glob(os.path.join(T, "code_parts", "*.tsv"))):
    for line in open(fp, encoding="utf-8"):
        line = line.rstrip("\n").rstrip("\r")
        if not line or "\t" not in line:
            continue
        idx, eng = line.split("\t", 1)
        try:
            idx = int(idx)
        except ValueError:
            continue
        if 0 <= idx < len(keys):
            trans[keys[idx]] = undisp(eng)
            seen.add(idx)

json.dump(trans, open(os.path.join(T, "translations_code.json"), "w", encoding="utf-8"),
          ensure_ascii=False, indent=1, sort_keys=True)
missing = [i for i in range(len(keys)) if i not in seen]
print("translated:", len(seen), "/", len(keys), "| skipped/missing:", len(missing))
