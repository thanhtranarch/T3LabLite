# -*- coding: utf-8 -*-
"""Two crash classes no unit test sees, because nothing here runs Revit.

1. A name read before the function's own local import of it: Python makes
   the name local for the whole function, so the read raises
   UnboundLocalError (2026-11-18: server.py Execute, every MCP call).
2. Python-3-only code in a module the IronPython 2.7 engine imports
   (startup.py, the MCP server and what they reach): SyntaxError or
   NameError at the first call (2026-11-18: `raise ... from` in
   family_schema.py killed the famigen_* MCP tools).

Run: python lite_guard/test_engine_safety.py
"""
import ast
import io
import os
import sys
import tokenize

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
EXT = REPO
LIB = os.path.join(EXT, 'lib')
# Entry points IronPython runs; everything they import is checked.
IPY_ROOTS = ('startup.py', 'lib/core/mcp_ipy_host.py', 'lib/assistant_pane.py',
             'lib/ribbon_memory.py')
# (file under lib/, message prefix): reviewed, not reached under IronPython.
IPY_ALLOWED = ()

PY3_NAMES = {
    'FileNotFoundError', 'PermissionError', 'NotADirectoryError', 'IsADirectoryError',
    'FileExistsError', 'ConnectionError', 'ConnectionResetError', 'ConnectionRefusedError',
    'ConnectionAbortedError', 'BrokenPipeError', 'TimeoutError', 'ModuleNotFoundError',
    'RecursionError', 'InterruptedError', 'ProcessLookupError', 'BlockingIOError'}
PY3_ATTRS = {
    ('time', 'monotonic'), ('time', 'perf_counter'), ('time', 'time_ns'), ('shutil', 'which'),
    ('subprocess', 'run'), ('subprocess', 'DEVNULL'), ('os', 'replace'), ('os', 'scandir'),
    ('math', 'inf'), ('math', 'isclose'), ('re', 'fullmatch'), ('shlex', 'quote'),
    ('json', 'JSONDecodeError'), ('datetime', 'timezone'), ('itertools', 'zip_longest'),
    ('functools', 'lru_cache'), ('contextlib', 'suppress'), ('tempfile', 'TemporaryDirectory')}
PY3_MODULES = {
    'urllib.request', 'urllib.parse', 'urllib.error', 'http.server', 'http.client', 'queue',
    'socketserver', 'configparser', 'html', 'pathlib', 'enum', 'concurrent.futures', 'typing',
    'sqlite3', 'builtins', 'winreg', 'dataclasses', 'asyncio', 'importlib.util'}
PY3_METHODS = {'removeprefix', 'removesuffix', 'casefold', 'fromisoformat'}
PY3_KWARGS = {'makedirs': {'exist_ok'}, 'Thread': {'daemon'}, 'max': {'default'},
              'min': {'default'}, 'split': {'maxsplit'}, 'rsplit': {'maxsplit'},
              'check_output': {'timeout'}}


def _read(path):
    with io.open(path, encoding='utf-8-sig') as f:
        return f.read()


def _py_files(root):
    for folder, _dirs, files in os.walk(root):
        if '__pycache__' in folder or os.sep + '.git' in folder:
            continue
        for name in files:
            if name.endswith('.py'):
                yield os.path.join(folder, name)


# ── 1 · read before a local import ───────────────────────────────────────
def _own_scope(fn):
    stack = list(fn.body)
    while stack:
        node = stack.pop()
        yield node
        for child in ast.iter_child_nodes(node):
            if not isinstance(child, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef,
                                      ast.Lambda)):
                stack.append(child)


def early_reads(source):
    """[(line, function, name, import line)] for names read above their local import."""
    out = []
    for fn in ast.walk(ast.parse(source)):
        if not isinstance(fn, (ast.FunctionDef, ast.AsyncFunctionDef)):
            continue
        imported, reads = {}, {}
        args = fn.args
        declared = set(a.arg for a in args.args + args.kwonlyargs)      # parameters are bound
        declared.update(a.arg for a in (args.vararg, args.kwarg) if a)
        for node in _own_scope(fn):
            if isinstance(node, (ast.Global, ast.Nonlocal)):
                declared.update(node.names)
            elif isinstance(node, (ast.Import, ast.ImportFrom)):
                for alias in node.names:
                    name = (alias.asname or alias.name).split('.')[0]
                    imported[name] = min(imported.get(name, node.lineno), node.lineno)
            elif isinstance(node, ast.Name) and isinstance(node.ctx, ast.Load):
                reads.setdefault(node.id, []).append(node.lineno)
        for name, line in imported.items():
            early = [r for r in reads.get(name, []) if r < line]
            if early and name not in declared:
                out.append((min(early), fn.name, name, line))
    return sorted(out)


# ── 2 · Python 3 only code under IronPython ──────────────────────────────
def _resolve(module, importer):
    parts = module.split('.')
    for base in (LIB, os.path.dirname(importer)):
        path = os.path.join(base, *parts)
        if os.path.isfile(path + '.py'):
            return path + '.py'
        if os.path.isfile(os.path.join(path, '__init__.py')):
            return os.path.join(path, '__init__.py')
    return None


def ipy_closure(roots):
    seen, todo = set(), list(roots)
    while todo:
        path = todo.pop()
        if path in seen:
            continue
        seen.add(path)
        try:
            tree = ast.parse(_read(path))
        except SyntaxError:
            continue
        for node in ast.walk(tree):
            names = []
            if isinstance(node, ast.Import):
                names = [a.name for a in node.names]
            elif isinstance(node, ast.ImportFrom) and node.module and not node.level:
                names = [node.module] + [node.module + '.' + a.name for a in node.names]
            for name in names:
                found = _resolve(name, path)
                if found and found not in seen:
                    todo.append(found)
    return sorted(seen)


def py2_traps(source):
    """[(line, what)]: code IronPython 2.7 cannot compile or run."""
    tree = ast.parse(source)
    parent = {}
    for node in ast.walk(tree):
        for child in ast.iter_child_nodes(node):
            parent[child] = node

    def in_try(node):
        node = parent.get(node)
        while node is not None:
            if isinstance(node, ast.Try):
                return True
            node = parent.get(node)
        return False

    future = set()
    for node in tree.body:
        if isinstance(node, ast.ImportFrom) and node.module == '__future__':
            future.update(a.name for a in node.names)
    out = []
    for node in ast.walk(tree):
        line = getattr(node, 'lineno', 0)
        if isinstance(node, ast.JoinedStr):
            out.append((line, 'f-string'))
        elif isinstance(node, (ast.Nonlocal, ast.AsyncFunctionDef, ast.Await, ast.YieldFrom,
                               ast.NamedExpr, ast.AnnAssign)):
            out.append((line, type(node).__name__))
        elif isinstance(node, (ast.FunctionDef, ast.Lambda)):
            args = node.args
            if args.kwonlyargs or getattr(args, 'posonlyargs', None):
                out.append((line, 'keyword-only arguments'))
            if any(a.annotation for a in args.args) or getattr(node, 'returns', None):
                out.append((line, 'annotation'))
        elif isinstance(node, ast.ClassDef) and node.keywords:
            out.append((line, 'class keywords'))
        elif isinstance(node, ast.Raise) and node.cause is not None:
            out.append((line, 'raise ... from'))
        elif isinstance(node, ast.Starred) and not isinstance(parent.get(node), ast.Call):
            out.append((line, 'starred outside a call'))
        elif isinstance(node, ast.Dict) and any(k is None for k in node.keys):
            out.append((line, '{**x}'))
        elif isinstance(node, ast.Name) and isinstance(node.ctx, ast.Load) \
                and node.id in PY3_NAMES:
            out.append((line, 'py3-only name ' + node.id))
        elif isinstance(node, ast.Attribute) and isinstance(node.value, ast.Name) \
                and (node.value.id, node.attr) in PY3_ATTRS and not in_try(node):
            out.append((line, 'py3-only {}.{}'.format(node.value.id, node.attr)))
        if isinstance(node, ast.Attribute) and node.attr in PY3_METHODS \
                and isinstance(parent.get(node), ast.Call) and not in_try(node):
            out.append((line, 'py3-only method .{}()'.format(node.attr)))
        if isinstance(node, ast.Call):
            func = node.func
            name = func.id if isinstance(func, ast.Name) else getattr(func, 'attr', None)
            kwargs = set(k.arg for k in node.keywords if k.arg)
            if isinstance(func, ast.Name) and name == 'super' and not node.args:
                out.append((line, 'super() without arguments'))
            elif isinstance(func, ast.Name) and name == 'open' \
                    and kwargs & {'encoding', 'newline', 'errors'}:
                out.append((line, 'builtin open(encoding=): use io.open'))
            elif name in PY3_KWARGS and kwargs & PY3_KWARGS[name] and not in_try(node):
                out.append((line, '{}({}=)'.format(name, sorted(kwargs & PY3_KWARGS[name])[0])))
            elif isinstance(func, ast.Name) and name == 'print' \
                    and 'print_function' not in future and (kwargs or len(node.args) != 1):
                out.append((line, 'print() needs __future__ print_function'))
        if isinstance(node, (ast.Import, ast.ImportFrom)) and not in_try(node):
            modules = [a.name for a in node.names] if isinstance(node, ast.Import) \
                else [node.module or '']
            for module in modules:
                if module in PY3_MODULES:
                    out.append((line, 'py3-only module ' + module))
    for token in tokenize.generate_tokens(io.StringIO(source).readline):
        if token.type == tokenize.NUMBER and '_' in token.string:
            out.append((token.start[0], 'numeric literal with _'))
    return sorted(set(out))


# ── run ──────────────────────────────────────────────────────────────────
FAILED = []


def check(label, ok, detail=''):
    print('  {} {}{}'.format('ok  ' if ok else 'FAIL', label, ('\n       ' + detail) if detail and not ok else ''))
    if not ok:
        FAILED.append(label)


def main():
    print('checkers catch what they are for')
    bad = "def f(task):\n    x = mod.go(task)\n    try:\n        from core import mod\n    except Exception:\n        pass\n    return x\n"
    check('a read above the local import is found', early_reads(bad) == [(2, 'f', 'mod', 4)])
    good = ("import mod\ndef f(task):\n    return mod.go(task)\ndef g():\n    from core import mod\n"
            "    return mod\ndef h(cfg=None):\n    if cfg is None:\n        from core import cfg\n    return cfg\n")
    check('module import / import-then-use is clean', early_reads(good) == [])
    traps = dict((what, line) for line, what in py2_traps(
        "def f(a):\n    try:\n        pass\n    except Exception as ex:\n        raise ValueError(a) from ex\n"
        "x = f'{1}'\nopen('p', encoding='utf-8')\n"))
    check('raise ... from, f-string, open(encoding=) are found',
          set(traps) == {'raise ... from', 'f-string', 'builtin open(encoding=): use io.open'})
    check('io.open and a guarded fallback are clean', py2_traps(
        "import io\ntry:\n    import queue\nexcept ImportError:\n    import Queue as queue\n"
        "io.open('p', encoding='utf-8')\n") == [])

    print('no name is read above its own local import')
    hits = []
    for path in _py_files(EXT):
        try:
            found = early_reads(_read(path))
        except SyntaxError:
            continue
        for line, fn, name, at in found:
            hits.append('{}:{} {}() reads `{}` above its import at line {}'.format(
                os.path.relpath(path, REPO), line, fn, name, at))
    check('extension is clean', not hits, '\n       '.join(hits))

    print('modules IronPython imports hold no Python 3 only code')
    roots = [os.path.join(EXT, r) for r in IPY_ROOTS]
    check('entry points exist', all(os.path.isfile(r) for r in roots),
          ', '.join(r for r in roots if not os.path.isfile(r)))
    closure = ipy_closure([r for r in roots if os.path.isfile(r)])
    check('closure reaches the MCP server', os.path.join(LIB, 'core', 'server.py') in closure)
    hits = []
    for path in closure:
        rel = os.path.relpath(path, LIB).replace(os.sep, '/')
        try:
            found = py2_traps(_read(path))
        except SyntaxError as ex:
            found = [(ex.lineno or 0, 'does not parse: {}'.format(ex.msg))]
        for line, what in found:
            if any(rel == f and what.startswith(w) for f, w in IPY_ALLOWED):
                continue
            hits.append('{}:{} {}'.format(rel, line, what))
    check('{} modules are clean'.format(len(closure)), not hits, '\n       '.join(hits))

    if FAILED:
        print('\n{} check(s) failed.'.format(len(FAILED)))
        return 1
    print('\nAll engine safety checks passed.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
