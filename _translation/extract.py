# -*- coding: utf-8 -*-
"""Extract translatable UI strings from BKT Python + XAML sources.

Collects values of safe-to-translate attributes only. Never touches id=,
image_mso=, idMso=, on_action, etc. Produces a JSON dict of unique strings.
"""
import os, re, json, sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Python attributes that hold user-facing text.
PY_ATTRS = ["label", "supertip", "screentip", "description", "title"]
# Matches:  label = "..."  or  label='...'  (single line, single/double quotes)
PY_RE = re.compile(
    r'\b(' + "|".join(PY_ATTRS) + r')\s*=\s*(u?)(["\'])((?:\\.|(?!\3).)*)\3'
)

# XAML attributes that hold user-facing text.
XAML_ATTRS = ["Content", "Header", "Text", "ToolTip", "Title"]
XAML_RE = re.compile(
    r'\b(' + "|".join(XAML_ATTRS) + r')\s*=\s*"([^"]*)"'
)

INCLUDE_DIRS = ["features", "bkt", "modules"]
SKIP_DIRS = {".git", "bin", "_translation"}

def iter_files(exts):
    for base in INCLUDE_DIRS:
        for dirpath, dirnames, filenames in os.walk(os.path.join(ROOT, base)):
            dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS]
            for fn in filenames:
                if os.path.splitext(fn)[1].lower() in exts:
                    yield os.path.join(dirpath, fn)

def looks_translatable(s):
    s = s.strip()
    if not s:
        return False
    # skip pure format/placeholder/non-letter strings
    if not re.search(r'[A-Za-zÄÖÜäöüß]', s):
        return False
    # skip strings that are obviously code-ish identifiers only
    if re.match(r'^[\w\.]+$', s) and ' ' not in s and len(s) < 4:
        return False
    return True

def main():
    strings = {}
    occ = 0
    # Python
    for fp in iter_files({".py"}):
        try:
            txt = open(fp, encoding="utf-8").read()
        except Exception:
            txt = open(fp, encoding="latin-1").read()
        for m in PY_RE.finditer(txt):
            val = m.group(4)
            if looks_translatable(val):
                strings.setdefault(val, "")
                occ += 1
    # XAML
    for fp in iter_files({".xaml"}):
        try:
            txt = open(fp, encoding="utf-8").read()
        except Exception:
            txt = open(fp, encoding="latin-1").read()
        for m in XAML_RE.finditer(txt):
            val = m.group(2)
            if looks_translatable(val):
                strings.setdefault(val, "")
                occ += 1

    out = os.path.join(ROOT, "_translation", "strings.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(strings, f, ensure_ascii=False, indent=1, sort_keys=True)
    print("Total occurrences:", occ)
    print("Unique strings:", len(strings))
    print("Written to:", out)

if __name__ == "__main__":
    main()
