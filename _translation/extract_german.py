# -*- coding: utf-8 -*-
"""Find German string LITERALS still in the code and classify each by context,
so we only translate user-facing text and never break programmatic strings
(dict keys, settings keys, == comparisons, tags).

Context buckets:
  call:<func>  - string is an argument to a function call (func name recorded)
  return       - returned from a function
  compare      - used in a ==/!= comparison        (UNSAFE - logic)
  dictkey      - used as a dict key / subscript key (UNSAFE - lookup)
  fstring      - literal piece of an f-string
  concat       - part of a +-concatenation
  assign/other - everything else
"""
import os, ast, json, re

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INCLUDE = ["features", "bkt", "modules"]
SKIP_DIRS = {".git", "bin", "_translation", "dotnet"}

UMLAUT = re.compile(r'[äöüßÄÖÜ]')
GWORDS = re.compile(
    r'\b(der|die|das|und|oder|mit|für|von|den|dem|des|ein|eine|einen|einem|einer'
    r'|nicht|wird|werden|wurde|wurden|kein|keine|alle|aus|auf|bei|nach|über|unter'
    r'|durch|gegen|ohne|vom|zum|zur|beim|sind|waren|haben|kann|können'
    r'|muss|müssen|soll|sollen|wenn|dann|als|wie|auch|schon|bitte|hier'
    r'|Fehler|Ergebnis|Datei|Ordner|Einstellung|Auswahl|Abbrechen|abgebrochen'
    r'|erstellt|erstellen|gelöscht|löschen|gespeichert|speichern|geladen|laden'
    r'|vorhanden|ausgewählt|markiert|markierte|Achtung|Hinweis|Warnung|Anzahl'
    r'|Folie|Folien|wählen|gewählt|geöffnet|öffnen|schließen|hinzufügen'
    r'|entfernen|ändern|geändert|verfügbar|benötigt|möchten|wollen|bzw|sowie)\b')

def is_german(s):
    return bool(UMLAUT.search(s) or GWORDS.search(s))

def translatable(s):
    t = s.strip()
    return len(t) >= 2 and bool(re.search(r'[A-Za-zÄÖÜäöüß]', t))

def func_name(call):
    f = call.func
    parts = []
    while isinstance(f, ast.Attribute):
        parts.append(f.attr); f = f.value
    if isinstance(f, ast.Name):
        parts.append(f.id)
    return ".".join(reversed(parts)) or "?"

def classify(node, parent, grand):
    # node is the str Constant; parent/grand are AST ancestors
    if isinstance(parent, ast.Compare):
        return "compare"
    if isinstance(parent, ast.Dict) and node in parent.keys:
        return "dictkey"
    if isinstance(parent, ast.Subscript):
        return "dictkey"
    if isinstance(parent, ast.Return):
        return "return"
    if isinstance(parent, ast.Call):
        return "call:" + func_name(parent)
    if isinstance(parent, ast.keyword):
        return "kwarg:" + (parent.arg or "?")
    if isinstance(parent, ast.JoinedStr):
        # f-string piece; classify by what the f-string is used in
        return "fstring/" + classify(parent, grand, None)
    if isinstance(parent, ast.BinOp):
        return "concat/" + classify(parent, grand, None)
    return parent.__class__.__name__.lower()

def main():
    results = {}  # german -> {"ctx": set, "locs": [..]}
    for base in INCLUDE:
        for dp, dn, fns in os.walk(os.path.join(ROOT, base)):
            dn[:] = [d for d in dn if d not in SKIP_DIRS]
            for fn in fns:
                if not fn.endswith(".py"):
                    continue
                fp = os.path.join(dp, fn)
                try:
                    tree = ast.parse(open(fp, encoding="utf-8").read())
                except Exception:
                    continue
                parent = {}
                for p in ast.walk(tree):
                    for c in ast.iter_child_nodes(p):
                        parent[id(c)] = p
                doc_ids = set()
                for node in ast.walk(tree):
                    if isinstance(node, (ast.Module, ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
                        if node.body and isinstance(node.body[0], ast.Expr) \
                           and isinstance(node.body[0].value, ast.Constant) \
                           and isinstance(node.body[0].value.value, str):
                            doc_ids.add(id(node.body[0].value))
                for node in ast.walk(tree):
                    if isinstance(node, ast.Constant) and isinstance(node.value, str) and id(node) not in doc_ids:
                        v = node.value
                        if translatable(v) and is_german(v):
                            p = parent.get(id(node))
                            g = parent.get(id(p)) if p is not None else None
                            ctx = classify(node, p, g)
                            e = results.setdefault(v, {"ctx": set(), "locs": []})
                            e["ctx"].add(ctx)
                            e["locs"].append("{}:{}".format(os.path.relpath(fp, ROOT).replace("\\", "/"), node.lineno))
    # serialize
    ser = {k: {"ctx": sorted(v["ctx"]), "locs": v["locs"]} for k, v in results.items()}
    json.dump(ser, open(os.path.join(ROOT, "_translation", "german_code_strings.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1, sort_keys=True)

    # summary of context buckets
    import collections
    buckets = collections.Counter()
    for v in results.values():
        for c in v["ctx"]:
            buckets[c.split(":")[0].split("/")[0]] += 1
    print("unique German code strings:", len(results))
    print("context buckets:")
    for c, n in buckets.most_common():
        print("  {:5d}  {}".format(n, c))

if __name__ == "__main__":
    main()
