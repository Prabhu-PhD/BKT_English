# -*- coding: utf-8 -*-
"""Apply code-string translations (translations_code.json: german->english,
both DECODED values) to .py sources.

Uses tokenize to find STRING tokens precisely (character-based positions, so
umlauts don't break offsets), decodes each with ast.literal_eval, and if its
value is in the map, rewrites just that literal token -- preserving the original
prefix/quote and correctly re-escaping the English text. Comments, identifiers,
and non-matching strings are never touched.

    python apply_german.py [--dry] [--root <dir>]
"""
import os, io, ast, sys, json, re, tokenize

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DRY = "--dry" in sys.argv
ROOT = REPO
if "--root" in sys.argv:
    ROOT = sys.argv[sys.argv.index("--root") + 1]

INCLUDE = ["features", "bkt", "modules"]
SKIP_DIRS = {".git", "bin", "_translation", "dotnet"}

def make_literal(english, raw_token):
    m = re.match(r"^([A-Za-z]*)('''|\"\"\"|'|\")", raw_token)
    prefix, quote = m.group(1), m.group(2)
    is_raw = "r" in prefix.lower()
    c = english
    if not is_raw:
        c = c.replace("\\", "\\\\")
        if len(quote) == 1:
            c = c.replace(quote, "\\" + quote)
        c = c.replace("\n", "\\n").replace("\t", "\\t").replace("\r", "\\r")
    return prefix + quote + c + quote

def iter_files():
    for base in INCLUDE:
        for dp, dn, fns in os.walk(os.path.join(ROOT, base)):
            dn[:] = [d for d in dn if d not in SKIP_DIRS]
            for fn in fns:
                if fn.endswith(".py"):
                    yield os.path.join(dp, fn)

def main():
    trans = json.load(open(os.path.join(REPO, "_translation", "translations_code.json"), encoding="utf-8"))
    trans = {k: v for k, v in trans.items() if v and v != k}
    files_changed = repls_total = 0
    for fp in iter_files():
        src = open(fp, encoding="utf-8").read()
        try:
            toks = list(tokenize.generate_tokens(io.StringIO(src).readline))
        except Exception:
            continue
        lines = src.splitlines(keepends=True)
        offs = [0]
        for ln in lines:
            offs.append(offs[-1] + len(ln))
        def pos2off(p):
            return offs[p[0] - 1] + p[1]
        repls = []
        for tok in toks:
            if tok.type == tokenize.STRING:
                try:
                    val = ast.literal_eval(tok.string)
                except Exception:
                    continue
                if isinstance(val, str) and val in trans:
                    repls.append((pos2off(tok.start), pos2off(tok.end), make_literal(trans[val], tok.string)))
        if repls:
            files_changed += 1
            repls_total += len(repls)
            if not DRY:
                repls.sort(key=lambda r: r[0], reverse=True)
                for s, e, new in repls:
                    src = src[:s] + new + src[e:]
                open(fp, "w", encoding="utf-8", newline="").write(src)
    print("DRY RUN" if DRY else "APPLIED", "(root=%s)" % ROOT)
    print("files changed:", files_changed)
    print("string replacements:", repls_total)

if __name__ == "__main__":
    main()
