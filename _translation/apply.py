# -*- coding: utf-8 -*-
"""Apply DE->EN translations to BKT sources.

Only rewrites the *value* of safe text attributes. Preserves attribute name,
surrounding whitespace, quote char, and optional u-prefix. Never alters id=,
image_mso=, idMso=, callbacks, etc. (they are not matched at all).

Usage:
    python apply.py            # apply in place
    python apply.py --dry      # report counts only, change nothing
"""
import os, re, json, sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DRY = "--dry" in sys.argv
# Optional: apply to a different install root (translations.json still read from REPO)
ROOT = REPO
if "--root" in sys.argv:
    ROOT = sys.argv[sys.argv.index("--root") + 1]

PY_ATTRS = ["label", "supertip", "screentip", "description", "title"]
PY_RE = re.compile(
    r'\b(' + "|".join(PY_ATTRS) + r')(\s*=\s*)(u?)(["\'])((?:\\.|(?!\4).)*)\4'
)
XAML_ATTRS = ["Content", "Header", "Text", "ToolTip", "Title"]
XAML_RE = re.compile(
    r'\b(' + "|".join(XAML_ATTRS) + r')(\s*=\s*)"([^"]*)"'
)

INCLUDE_DIRS = ["features", "bkt", "modules"]
SKIP_DIRS = {".git", "bin", "_translation"}

def iter_files(exts):
    for base in INCLUDE_DIRS:
        for dp, dn, fns in os.walk(os.path.join(ROOT, base)):
            dn[:] = [d for d in dn if d not in SKIP_DIRS]
            for fn in fns:
                if os.path.splitext(fn)[1].lower() in exts:
                    yield os.path.join(dp, fn)

def main():
    trans = json.load(open(os.path.join(REPO, "_translation", "translations.json"),
                            encoding="utf-8"))
    # only entries that actually change something
    trans = {k: v for k, v in trans.items() if v and v != k}

    stats = {"files_changed": 0, "py_repl": 0, "xaml_repl": 0, "missing": set()}

    def py_sub(m):
        attr, eq, up, q, val = m.group(1), m.group(2), m.group(3), m.group(4), m.group(5)
        if val in trans:
            new = trans[val]
            # Escape the delimiter quote if it appears unescaped in the value
            # (e.g. apostrophes inside a single-quoted literal). The source
            # value never contained the delimiter (regex matched cleanly), but a
            # translation might (Python's, o'clock, ...).
            new = re.sub(r'(?<!\\)' + re.escape(q), '\\\\' + q, new)
            stats["py_repl"] += 1
            return "{}{}{}{}{}{}".format(attr, eq, up, q, new, q)
        return m.group(0)

    def xaml_sub(m):
        attr, eq, val = m.group(1), m.group(2), m.group(3)
        if val in trans:
            new = trans[val].replace('"', '&quot;')  # XAML attr is "-delimited
            stats["xaml_repl"] += 1
            return '{}{}"{}"'.format(attr, eq, new)
        return m.group(0)

    for fp in iter_files({".py"}):
        txt = open(fp, encoding="utf-8").read()
        new = PY_RE.sub(py_sub, txt)
        if new != txt:
            stats["files_changed"] += 1
            if not DRY:
                open(fp, "w", encoding="utf-8", newline="").write(new)

    for fp in iter_files({".xaml"}):
        txt = open(fp, encoding="utf-8").read()
        new = XAML_RE.sub(xaml_sub, txt)
        if new != txt:
            stats["files_changed"] += 1
            if not DRY:
                open(fp, "w", encoding="utf-8", newline="").write(new)

    print("DRY RUN" if DRY else "APPLIED")
    print("Files changed:", stats["files_changed"])
    print("Python replacements:", stats["py_repl"])
    print("XAML replacements:", stats["xaml_repl"])

if __name__ == "__main__":
    main()
