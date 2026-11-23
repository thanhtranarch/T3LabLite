#! python3
# -*- coding: utf-8 -*-
"""audit_ipy_compat.py — the MCP server runs IronPython 2.7: code it can load
must not use Python-3-only syntax or APIs.

Scope: every lib module reachable by static import (top level and inside
functions) from the server's entry points (core/server.py, core/mcp_ipy_host.py,
core/tool_specs.py, core/tools/*.py). A violation there fails at the moment an
agent calls the tool, not at import (2026-11-22: casefold / open(encoding=) in
logic modules the windows run under CPython).

  I1  f-string                          SyntaxError on import
  I2  str.casefold()                    AttributeError -> Snippets._compat.fold()
  I3  open(..., encoding/errors/newline) TypeError -> io.open
  I4  nonlocal / keyword-only args / annotations / walrus / async / yield from /
      `raise ... from` / zero-arg super()           SyntaxError or TypeError
  I5  os.replace, os.makedirs(exist_ok=), pathlib, typing   missing on 2.7
  I6  json.dump(s) without ensure_ascii=False  IronPython's ASCII escaper raises on
      accented text -> core.jsonsafe.dumps (ASCII, CPython bytes unchanged) or
      ensure_ascii=False (text written via io.open(encoding='utf-8') / in memory)
  I7  `except OSError` around open()       2.7 IOError is not an OSError -> EnvironmentError
  I8  open(.., 'w'/'a', encoding='utf-8-sig'/'utf-16')  IronPython writes a BOM on
      every write -> encoding='utf-8' + write the BOM (u'\\ufeff') once
  I9  list comprehension rebinds an `except .. as name` used after it (2.7 leaks
      the loop variable) -> rename the loop variable
  I10 '\\uXXXX' in a plain str without unicode_literals   6 literal chars on 2.7 -> u'...'

A line ending in `# ipy-ok: <reason>` is exempt (a guarded fallback).

Run:  python3 lite_guard/audit_ipy_compat.py [--quiet] [--list]
"""

import ast
import io
import os
import re
import sys
import tokenize

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(ROOT, 'lib')
ENTRY = ['core/server.py', 'core/mcp_ipy_host.py', 'core/tool_specs.py']

# Module (lib-relative path) -> why its CPython-only code is never reached under
# IronPython. Keep this short: fix the code instead when the server can call it.
ALLOW = {
}


def _rel(path):
    return os.path.relpath(path, LIB).replace(os.sep, '/')


def _resolve(dotted):
    """lib-relative .py path of a dotted module name, or None (stdlib / .NET)."""
    base = os.path.join(LIB, *dotted.split('.'))
    for cand in (base + '.py', os.path.join(base, '__init__.py')):
        if os.path.isfile(cand):
            return _rel(cand)
    return None


def _package(rel, level):
    """Dotted package a `from .<level>` import in lib module `rel` starts from."""
    parts = rel[:-3].split('/')[:-1]        # pkg/mod.py and pkg/__init__.py -> pkg
    if level > 1:
        parts = parts[:len(parts) - (level - 1)]
    return '.'.join(parts)


def imports_of(rel):
    try:
        tree = ast.parse(io.open(os.path.join(LIB, rel), encoding='utf-8-sig').read())
    except (SyntaxError, IOError, OSError):
        return set()
    out = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for a in node.names:
                r = _resolve(a.name)
                if r:
                    out.add(r)
        elif isinstance(node, ast.ImportFrom):
            if node.level:
                pkg = _package(rel, node.level)
                mod = '.'.join(p for p in (pkg, node.module) if p)
            else:
                mod = node.module
            if not mod:
                continue
            r = _resolve(mod)
            if r:
                out.add(r)
            for a in node.names:            # from core import tool_specs
                r = _resolve('{}.{}'.format(mod, a.name))
                if r:
                    out.add(r)
    return out


def reachable():
    entries = list(ENTRY)
    tools = os.path.join(LIB, 'core', 'tools')
    if os.path.isdir(tools):
        entries += ['core/tools/' + n for n in sorted(os.listdir(tools)) if n.endswith('.py')]
    seen, todo = set(), [e for e in entries if os.path.isfile(os.path.join(LIB, e))]
    while todo:
        rel = todo.pop()
        if rel in seen:
            continue
        seen.add(rel)
        todo.extend(imports_of(rel) - seen)
    return seen


_PY3_MODULES = ('pathlib', 'typing', 'concurrent', 'asyncio', 'dataclasses', 'enum',
                'secrets', 'statistics', 'zoneinfo', 'graphlib')


def _future(tree, name):
    return any(isinstance(n, ast.ImportFrom) and n.module == '__future__'
               and any(a.name == name for a in n.names) for n in tree.body)


def _json_names(tree):
    """(names bound to the json module, names bound to json.dump / json.dumps)."""
    mods, funcs = set(), set()
    for n in ast.walk(tree):
        if isinstance(n, ast.Import):
            for a in n.names:
                if a.name == 'json':
                    mods.add(a.asname or 'json')
        elif isinstance(n, ast.ImportFrom) and n.module == 'json' and not n.level:
            for a in n.names:
                if a.name in ('dump', 'dumps'):
                    funcs.add(a.asname or a.name)
    return mods, funcs


def _is_json_dump(call, mods, funcs):
    f = call.func
    if isinstance(f, ast.Attribute) and f.attr in ('dump', 'dumps'):
        return isinstance(f.value, ast.Name) and f.value.id in mods
    return isinstance(f, ast.Name) and f.id in funcs


def _ascii_false(call):
    return any(k.arg == 'ensure_ascii' and isinstance(k.value, ast.Constant)
               and k.value.value is False for k in call.keywords)


_OPEN_FUNCS = {'open': (1, 3), 'io.open': (1, 3), 'codecs.open': (1, 2)}   # mode, encoding pos


def _open_kind(call):
    f = call.func
    if isinstance(f, ast.Name) and f.id == 'open':
        return 'open'
    if (isinstance(f, ast.Attribute) and f.attr == 'open' and isinstance(f.value, ast.Name)
            and f.value.id in ('io', 'codecs', '_io')):
        return 'codecs.open' if f.value.id == 'codecs' else 'io.open'
    return None


def _arg(call, pos, name):
    for k in call.keywords:
        if k.arg == name:
            return k.value
    return call.args[pos] if len(call.args) > pos else None


def _const_str(node):
    return node.value if isinstance(node, ast.Constant) and isinstance(node.value, str) else None


_BOM_ENCODINGS = ('utf-8-sig', 'utf_8_sig', 'utf8-sig', 'utf-16', 'utf_16', 'utf16')
_IO_CATCHERS = ('IOError', 'EnvironmentError', 'Exception', 'BaseException')


def _handler_names(handler):
    t = handler.type
    elts = t.elts if isinstance(t, ast.Tuple) else [t] if t is not None else []
    return [e.id if isinstance(e, ast.Name) else getattr(e, 'attr', '') for e in elts]


def _scope_walk(nodes):
    """ast.walk over `nodes` without entering nested functions / classes / lambdas."""
    todo = list(nodes)
    while todo:
        n = todo.pop()
        yield n
        for c in ast.iter_child_nodes(n):
            if not isinstance(c, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef,
                                  ast.Lambda)):
                todo.append(c)


def _comp_leaks(handler):
    """List comprehensions in an `except .. as name` body that rebind name before a use."""
    name = handler.name
    body = list(_scope_walk(handler.body))
    out = []
    for comp in body:
        if not isinstance(comp, ast.ListComp):
            continue
        if not any(isinstance(t, ast.Name) and t.id == name
                   for g in comp.generators for t in ast.walk(g.target)):
            continue
        end = (comp.end_lineno, comp.end_col_offset)
        if any(isinstance(n, ast.Name) and n.id == name and isinstance(n.ctx, ast.Load)
               and (n.lineno, n.col_offset) >= end for n in body):
            out.append(comp)
    return out


_U_ESCAPE = re.compile(r'(?<!\\)(?:\\\\)*\\[uUN]')


def _plain_u_escapes(src):
    """Lines of plain (no u/r/b/f prefix) str literals holding a backslash-u/U/N escape."""
    if '\\u' not in src and '\\U' not in src and '\\N' not in src:
        return []
    out = []
    try:
        for tok in tokenize.generate_tokens(io.StringIO(src).readline):
            if tok.type != tokenize.STRING:
                continue
            text = tok.string
            prefix = text[:len(text) - len(text.lstrip('rRuUbBfF'))].lower()
            if not prefix and _U_ESCAPE.search(text):
                out.append(tok.start[0])
    except (tokenize.TokenError, SyntaxError):
        pass
    return out


def scan(rel):
    """[(line, code, text)] for one lib module."""
    return scan_source(io.open(os.path.join(LIB, rel), encoding='utf-8-sig').read())


def scan_source(src):
    """[(line, code, text)] for one module's source."""
    try:
        tree = ast.parse(src)
    except SyntaxError as ex:
        return [(ex.lineno or 0, 'I4', 'not valid Python 3 either: {}'.format(ex.msg))]
    out = []
    future_print = _future(tree, 'print_function')
    json_mods, json_funcs = _json_names(tree)
    lines = src.splitlines()

    def hit(node, code, text):
        line = getattr(node, 'lineno', 0)
        if 0 < line <= len(lines) and '# ipy-ok:' in lines[line - 1]:
            return
        out.append((line, code, text))

    for node in ast.walk(tree):
        if isinstance(node, ast.JoinedStr):
            hit(node, 'I1', 'f-string')
        elif isinstance(node, ast.Call):
            f = node.func
            if isinstance(f, ast.Attribute) and f.attr == 'casefold':
                hit(node, 'I2', '.casefold() -> Snippets._compat.fold()')
            elif isinstance(f, ast.Name) and f.id == 'open' and any(
                    k.arg in ('encoding', 'errors', 'newline') for k in node.keywords):
                hit(node, 'I3', 'open(..., encoding=) -> io.open')
            elif isinstance(f, ast.Name) and f.id == 'super' and not node.args:
                hit(node, 'I4', 'super() without arguments')
            elif (isinstance(f, ast.Name) and f.id == 'print' and node.keywords
                  and not future_print):
                hit(node, 'I4', 'print(..., kw=) without from __future__ import print_function')
            elif (isinstance(f, ast.Attribute) and f.attr == 'replace'
                  and isinstance(f.value, ast.Name) and f.value.id == 'os'):
                hit(node, 'I5', 'os.replace (2.7: os.rename after removing the target)')
            elif (isinstance(f, ast.Attribute) and f.attr == 'makedirs'
                  and any(k.arg == 'exist_ok' for k in node.keywords)):
                hit(node, 'I5', 'os.makedirs(exist_ok=)')
            elif _is_json_dump(node, json_mods, json_funcs) and not _ascii_false(node):
                hit(node, 'I6', 'json.dump(s) escapes ASCII -> core.jsonsafe.dumps(), '
                                'or ensure_ascii=False for text kept in memory / io.open utf-8')
            kind = _open_kind(node)
            if kind:
                mpos, epos = _OPEN_FUNCS[kind]
                mode = _const_str(_arg(node, mpos, 'mode')) or ''
                enc = (_const_str(_arg(node, epos, 'encoding')) or '').lower()
                if enc in _BOM_ENCODINGS and any(c in mode for c in 'wax'):
                    hit(node, 'I8', "{} writing '{}' (BOM on every write) -> encoding='utf-8' "
                                    "and write u'\\ufeff' once".format(kind, enc))
        elif isinstance(node, ast.Try):
            opens = any(isinstance(n, ast.Call) and _open_kind(n)
                        for b in node.body for n in ast.walk(b))
            for h in node.handlers:
                names = _handler_names(h)
                if (opens and 'OSError' in names
                        and not any(n in _IO_CATCHERS for n in names)):
                    hit(h, 'I7', 'except OSError around open() misses IOError on 2.7 '
                                 '-> except EnvironmentError')
                if h.name:
                    for comp in _comp_leaks(h):
                        hit(comp, 'I9', "list comprehension rebinds '{}' from `except .. as "
                                        "{}` (2.7 leaks it) -> rename the loop variable"
                                        .format(h.name, h.name))
        elif isinstance(node, ast.Nonlocal):
            hit(node, 'I4', 'nonlocal')
        elif isinstance(node, (ast.AsyncFunctionDef, ast.Await, ast.YieldFrom, ast.NamedExpr)):
            hit(node, 'I4', type(node).__name__)
        elif isinstance(node, ast.AnnAssign):
            hit(node, 'I4', 'variable annotation')
        elif isinstance(node, ast.arguments):
            if node.kwonlyargs or node.posonlyargs:
                hit(node, 'I4', 'keyword-only / positional-only arguments')
            if any(a.annotation is not None for a in node.args + node.kwonlyargs):
                hit(node, 'I4', 'argument annotation')
        elif isinstance(node, (ast.FunctionDef,)) and node.returns is not None:
            hit(node, 'I4', 'return annotation')
        elif isinstance(node, ast.Raise) and node.cause is not None:
            hit(node, 'I4', 'raise ... from')
        elif isinstance(node, ast.Import):
            for a in node.names:
                if a.name.split('.')[0] in _PY3_MODULES:
                    hit(node, 'I5', 'import {}'.format(a.name))
        elif isinstance(node, ast.ImportFrom) and node.module and not node.level:
            if node.module.split('.')[0] in _PY3_MODULES:
                hit(node, 'I5', 'from {} import'.format(node.module))
    if not _future(tree, 'unicode_literals'):
        for line in _plain_u_escapes(src):
            hit(type('L', (), {'lineno': line})(), 'I10',
                "'\\u..' in a plain str is 6 literal chars on 2.7 -> u'...' prefix")
    return sorted(out)


def run(quiet=False, listing=False):
    mods = sorted(reachable())
    if listing:
        print('\n'.join(mods))
    bad = 0
    for rel in mods:
        if rel in ALLOW:
            continue
        for line, code, text in scan(rel):
            bad += 1
            print('{}:{}: {} {}'.format(rel, line, code, text))
    if not quiet or bad:
        print('IPY COMPAT: {} module(s) the server can load · {} violation(s)'.format(
            len(mods), bad))
    elif quiet:
        print('IPY COMPAT: clean ({} modules)'.format(len(mods)))
    return 1 if bad else 0


if __name__ == '__main__':
    sys.exit(run('--quiet' in sys.argv, '--list' in sys.argv))
