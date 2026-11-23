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

A line ending in `# ipy-ok: <reason>` is exempt (a guarded fallback).

Run:  python3 lite_guard/audit_ipy_compat.py [--quiet] [--list]
"""

import ast
import io
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(ROOT, 'lib')
ENTRY = ['core/server.py', 'core/mcp_ipy_host.py', 'core/tool_specs.py']
PACKAGES = ('core', 'Snippets', 'Services', 'Intelligence', 'GUI', 'Selection', 'Pattern')

# Module (lib-relative path) -> why its CPython-only code is never reached under
# IronPython. Keep this short: fix the code instead when the server can call it.
ALLOW = {
}


def _rel(path):
    return os.path.relpath(path, LIB).replace(os.sep, '/')


def _resolve(dotted):
    """lib-relative .py path of a dotted module name, or None (stdlib / .NET)."""
    parts = dotted.split('.')
    if parts[0] not in PACKAGES:
        return None
    base = os.path.join(LIB, *parts)
    for cand in (base + '.py', os.path.join(base, '__init__.py')):
        if os.path.isfile(cand):
            return _rel(cand)
    return None


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
        elif isinstance(node, ast.ImportFrom) and node.module and not node.level:
            r = _resolve(node.module)
            if r:
                out.add(r)
            for a in node.names:            # from core import tool_specs
                r = _resolve('{}.{}'.format(node.module, a.name))
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


def _future_print(tree):
    return any(isinstance(n, ast.ImportFrom) and n.module == '__future__'
               and any(a.name == 'print_function' for a in n.names) for n in tree.body)


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
    future_print = _future_print(tree)
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
