#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Lite guard -- catches ribbon tools and usage tracking lost when folders are
copied over from t3lab_dev.

T3LabLite is split from t3lab_dev, and files are often brought across by
copying whole folders. That has silently dropped Lite-only code before:
5751bd2 overwrote lib/_cpython_bootstrap.py and lib/core/server.py with the
dev copies, and the T3Lab Space dashboard stopped getting ribbon and MCP rows
for a release.

Checks (any error exits 1, warnings never fail):
  markers  every Lite-only hook listed in manifest.json is still in its file
  tools    every ribbon folder recorded in manifest.json still exists, and
           every button still has its script.py / bundle.yaml
  imports  T3Lab.tab scripts and lib modules only import lib modules that
           exist
  xaml     every *.xaml file named in Python code exists

Usage:
  python lite_guard/guard.py                check the working tree (CI)
  python lite_guard/guard.py --staged       check what is about to be committed
  python lite_guard/guard.py --update       re-record the ribbon tool list after
                                            adding or removing a tool on purpose
  python lite_guard/guard.py --install-hook run the --staged check on every commit

Plain CPython 3, standard library only -- no Revit or pyRevit needed.
"""
from __future__ import print_function

import argparse
import ast
import io
import json
import os
import re
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MANIFEST = os.path.join(ROOT, 'lite_guard', 'manifest.json')
HOOKS_PATH = 'lite_guard/githooks'

TAB_DIR = 'T3Lab.tab'
LIB_DIR = 'lib'

# pyRevit bundle folder types (folder name suffix)
CONTAINER_EXTS = ('.tab', '.panel', '.stack', '.pulldown',
                  '.splitbutton', '.splitpushbutton')
SCRIPT_EXTS = ('.pushbutton', '.smartbutton', '.panelbutton')
BUNDLE_EXTS = ('.urlbutton', '.linkbutton', '.invokebutton')
COMPONENT_EXTS = CONTAINER_EXTS + SCRIPT_EXTS + BUNDLE_EXTS + (
    '.content', '.nobutton')

SKIP_DIRS = ('.git', '__pycache__')

# Tool folders hold no tracking code: a ribbon click is recorded because every
# T3Lab.tab script.py calls _cpython_bootstrap.init_cpython_paths(). A tool
# folder replaced from t3lab_dev must keep that call.
RIBBON_TRACKING_RE = re.compile(r'init_cpython_paths\(')


# ── file sources ─────────────────────────────────────────────────────────────

def _git(args, data=None):
    proc = subprocess.Popen(['git'] + args, cwd=ROOT, stdin=subprocess.PIPE,
                            stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    out, err = proc.communicate(data)
    if proc.returncode != 0:
        raise RuntimeError('git {} failed: {}'.format(
            ' '.join(args), err.decode('utf-8', 'replace').strip()))
    return out


class WorkTree(object):
    """Files as they are on disk."""

    def __init__(self, root):
        self.root = root
        self._files = None

    def files(self):
        if self._files is None:
            found = set()
            for dirpath, dirnames, filenames in os.walk(self.root):
                dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS]
                rel = os.path.relpath(dirpath, self.root)
                for name in filenames:
                    path = name if rel == '.' else os.path.join(rel, name)
                    found.add(path.replace(os.sep, '/'))
            self._files = found
        return self._files

    def read(self, path):
        try:
            with io.open(os.path.join(self.root, path), 'rb') as fh:
                return fh.read()
        except (IOError, OSError):
            return None


class IndexTree(object):
    """Files as `git commit` is about to record them (the index).

    Checking the index rather than the disk matters when only part of a copy
    is staged: a bad file can be staged while the disk copy is already fixed.
    """

    def __init__(self):
        out = _git(['ls-files', '-z'])
        self._files = set(p for p in out.decode('utf-8').split('\0') if p)
        self._cache = {}

    def files(self):
        return self._files

    def read(self, path):
        if path not in self._files:
            return None
        if path not in self._cache:
            self._cache[path] = _git(['show', ':' + path])
        return self._cache[path]

    def prefetch(self, paths):
        """Read many files with one `git cat-file --batch` process."""
        paths = [p for p in paths if p in self._files and p not in self._cache]
        if not paths:
            return
        out = _git(['cat-file', '--batch'],
                   ''.join(':{}\n'.format(p) for p in paths).encode('utf-8'))
        pos = 0
        for path in paths:
            end = out.index(b'\n', pos)
            header = out[pos:end].split()
            pos = end + 1
            if header[-1] == b'missing':
                continue
            size = int(header[-1])
            self._cache[path] = out[pos:pos + size]
            pos += size + 1


def _text(raw):
    return raw.decode('utf-8', 'replace') if raw is not None else None


# ── report ───────────────────────────────────────────────────────────────────

class Report(object):

    def __init__(self):
        self.errors = []
        self.warnings = []

    def error(self, check, msg):
        self.errors.append((check, msg))

    def warn(self, check, msg):
        self.warnings.append((check, msg))

    def print_all(self, out=sys.stdout):
        for label, items in (('WARNING', self.warnings), ('ERROR', self.errors)):
            for check, msg in items:
                print('  {:<7} [{}] {}'.format(label, check, msg), file=out)


# ── manifest ─────────────────────────────────────────────────────────────────

def load_manifest():
    with io.open(MANIFEST, 'r', encoding='utf-8') as fh:
        return json.load(fh)


def save_manifest(manifest):
    with io.open(MANIFEST, 'w', encoding='utf-8', newline='\n') as fh:
        fh.write(json.dumps(manifest, indent=2, ensure_ascii=False) + '\n')


# ── check: markers ───────────────────────────────────────────────────────────

def marker_ok(marker, text):
    """True when `text` still carries the hook `marker` describes."""
    if text is None:
        return False
    found = len(re.findall(marker['pattern'], text, re.M))
    return found >= marker.get('min_count', 1)


def check_markers(tree, manifest, report):
    for marker in manifest.get('markers', []):
        path = marker['file']
        text = _text(tree.read(path))
        if text is None:
            report.error('markers', '{} is missing -- {}'.format(
                path, marker['why']))
        elif not marker_ok(marker, text):
            report.error('markers', '{} lost /{}/ -- {}'.format(
                path, marker['pattern'], marker['why']))


# ── check: ribbon tools ──────────────────────────────────────────────────────

def ribbon_components(tree):
    """Every pyRevit bundle folder under T3Lab.tab (repo-relative)."""
    found = set()
    for path in tree.files():
        if not path.startswith(TAB_DIR + '/'):
            continue
        parts = path.split('/')[:-1]
        for i, part in enumerate(parts):
            if part.endswith(COMPONENT_EXTS):
                found.add('/'.join(parts[:i + 1]))
    return found


def _bundle_name(folder):
    return folder.rsplit('/', 1)[-1].rsplit('.', 1)[0]


def read_layout(text):
    """The `layout:` list of a bundle.yaml, or None when it has none.

    A tiny reader for the block-list form this repo uses -- no PyYAML needed.
    """
    if text is None:
        return None
    items = None
    for line in text.splitlines():
        stripped = line.strip()
        if items is None:
            if re.match(r'^layout\s*:\s*$', line):
                items = []
            continue
        if not stripped or stripped.startswith('#'):
            continue
        if not stripped.startswith('- ') and not line.startswith((' ', '\t')):
            break
        if stripped.startswith('- '):
            name = stripped[2:].strip().strip('"\'')
            # separators (---) and slideouts (>>>) are not bundles
            if name and not set(name) <= set('->'):
                items.append(name)
    return items


def check_tools(tree, manifest, report):
    present = ribbon_components(tree)
    recorded = set(manifest.get('tools', []))
    files = tree.files()

    for folder in sorted(recorded - present):
        report.error('tools', 'ribbon folder is gone: {}'.format(folder))
    for folder in sorted(present - recorded):
        report.warn('tools', 'new ribbon folder, not recorded yet (run '
                             '--update if it is meant to ship): {}'.format(folder))

    for folder in sorted(present):
        children = set(p[len(folder) + 1:] for p in files
                       if p.startswith(folder + '/') and
                       '/' not in p[len(folder) + 1:])
        if folder.endswith(SCRIPT_EXTS):
            if not any(c.startswith('script.') for c in children):
                report.error('tools', 'button has no script.py: {}'.format(folder))
            elif 'script.py' in children:
                text = _text(tree.read(folder + '/script.py'))
                if not RIBBON_TRACKING_RE.search(text):
                    report.error('tools', '{}/script.py does not call '
                                          '_cpython_bootstrap.init_cpython_paths(), '
                                          'so clicks on it are not tracked'.format(folder))
        elif folder.endswith(BUNDLE_EXTS):
            if 'bundle.yaml' not in children:
                report.error('tools', 'button has no bundle.yaml: {}'.format(folder))

        if not folder.endswith(CONTAINER_EXTS):
            continue
        layout = read_layout(_text(tree.read(folder + '/bundle.yaml')))
        if layout is None:
            continue
        sub = set(_bundle_name(c) for c in present
                  if c.startswith(folder + '/') and
                  '/' not in c[len(folder) + 1:])
        for name in layout:
            if name not in sub:
                report.warn('tools', '{}/bundle.yaml lists "{}" but there is '
                                     'no such folder'.format(folder, name))
        for name in sorted(sub - set(layout)):
            report.warn('tools', '"{}" is not in the layout of '
                                 '{}/bundle.yaml'.format(name, folder))


# ── check: imports ───────────────────────────────────────────────────────────

def lib_modules(tree):
    """Dotted names importable from lib/ (modules, packages, namespace dirs)."""
    mods = set()
    prefix = LIB_DIR + '/'
    for path in tree.files():
        if not path.startswith(prefix):
            continue
        parts = path[len(prefix):].split('/')
        for i in range(1, len(parts)):
            mods.add('.'.join(parts[:i]))
        if parts[-1].endswith('.py'):
            name = parts[-1][:-3]
            if name != '__init__':
                mods.add('.'.join(parts[:-1] + [name]))
    return mods


def _imports(tree_ast):
    """(lineno, level, module, in_try) for every import in a module."""
    found = []

    def visit(node, in_try):
        for child in ast.iter_child_nodes(node):
            child_try = in_try or isinstance(child, ast.Try)
            if isinstance(child, ast.Import):
                for alias in child.names:
                    found.append((child.lineno, 0, alias.name, in_try))
            elif isinstance(child, ast.ImportFrom):
                if child.module:
                    found.append((child.lineno, child.level or 0,
                                  child.module, in_try))
            visit(child, child_try)

    visit(tree_ast, False)
    return found


def check_imports(tree, manifest, report):
    ignored = set((i['file'], i['module'])
                  for i in manifest.get('ignore_imports', []))
    mods = lib_modules(tree)
    tops = set(m.split('.')[0] for m in mods)
    py_files = sorted(p for p in tree.files() if p.endswith('.py') and (
        p.startswith(TAB_DIR + '/') or p.startswith(LIB_DIR + '/')))
    if hasattr(tree, 'prefetch'):
        tree.prefetch(py_files)

    for path in py_files:
        raw = tree.read(path)
        try:
            parsed = ast.parse(raw, filename=path)
        except (SyntaxError, ValueError) as exc:
            report.warn('imports', 'cannot parse {} ({}), imports not '
                                   'checked'.format(path, exc.__class__.__name__))
            continue

        package = None
        if path.startswith(LIB_DIR + '/'):
            package = path[len(LIB_DIR) + 1:].split('/')[:-1]

        for lineno, level, module, in_try in _imports(parsed):
            if level:
                if package is None or level - 1 > len(package):
                    continue
                base = package[:len(package) - (level - 1)]
                target = '.'.join(base + [module])
            else:
                if module.split('.')[0] not in tops:
                    continue      # stdlib, .NET, pyRevit or a sibling file
                target = module
            if target in mods or (path, target) in ignored:
                continue
            msg = '{}:{} imports {}, which is not in lib/'.format(
                path, lineno, target)
            if in_try:
                report.warn('imports', msg + ' (inside try, so it may be optional)')
            else:
                report.error('imports', msg)


# ── check: xaml ──────────────────────────────────────────────────────────────

# A literal file name only: "ManaSheets.xaml" or ".../Tools/ManaSheets.xaml".
# Leaves out "%s.xaml", the bare ".xaml" extension and the System.Xaml assembly.
_XAML_RE = re.compile(r'''["'](?:[^"'\r\n]*[\\/])?([\w][\w \-]*\.xaml)["']''')


def check_xaml(tree, report):
    files = tree.files()
    have = set(p.rsplit('/', 1)[-1].lower() for p in files)
    for path in sorted(files):
        if not path.endswith('.py') or not (
                path.startswith((TAB_DIR + '/', LIB_DIR + '/')) or '/' not in path):
            continue
        text = _text(tree.read(path))
        for lineno, line in enumerate(text.splitlines(), 1):
            for match in _XAML_RE.finditer(line):
                name = match.group(1)
                if name.lower() not in have:
                    report.error('xaml', '{}:{} needs {}, which is missing'.format(
                        path, lineno, name))


# ── staged deletions (information only) ─────────────────────────────────────

def staged_deletions():
    out = _git(['diff', '--cached', '--name-only', '--diff-filter=D', '-z'])
    return [p for p in out.decode('utf-8').split('\0') if p]


# ── entry points ─────────────────────────────────────────────────────────────

def run_checks(tree, manifest):
    report = Report()
    check_markers(tree, manifest, report)
    check_tools(tree, manifest, report)
    check_imports(tree, manifest, report)
    check_xaml(tree, report)
    return report


def cmd_check(staged):
    manifest = load_manifest()
    tree = IndexTree() if staged else WorkTree(ROOT)
    report = run_checks(tree, manifest)

    where = 'staged files' if staged else 'working tree'
    print('lite_guard: checking {}'.format(where))

    if staged:
        deleted = staged_deletions()
        if deleted:
            print('  This commit deletes {} file(s):'.format(len(deleted)))
            for path in deleted[:30]:
                print('    - ' + path)
            if len(deleted) > 30:
                print('    ... and {} more'.format(len(deleted) - 30))

    report.print_all()
    if report.errors:
        print('lite_guard: FAILED -- {} error(s). Something Lite needs was '
              'removed or overwritten.'.format(len(report.errors)))
        print('  Restore it (git checkout HEAD -- <file>, then merge the '
              't3lab_dev change in by hand),')
        print('  or, if the removal is on purpose, update '
              'lite_guard/manifest.json.')
        return 1
    print('lite_guard: OK ({} warning(s))'.format(len(report.warnings)))
    return 0


def cmd_update():
    manifest = load_manifest()
    before = set(manifest.get('tools', []))
    after = ribbon_components(WorkTree(ROOT))
    for folder in sorted(after - before):
        print('  + ' + folder)
    for folder in sorted(before - after):
        print('  - ' + folder)
    manifest['tools'] = sorted(after)
    save_manifest(manifest)
    print('lite_guard: recorded {} ribbon folders in manifest.json'.format(
        len(after)))
    return 0


def cmd_install_hook():
    _git(['config', 'core.hooksPath', HOOKS_PATH])
    print('lite_guard: every commit in this clone now runs the check '
          '(core.hooksPath = {})'.format(HOOKS_PATH))
    return 0


def main(argv=None):
    # Windows consoles default to a code page that cannot print every path.
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            stream.reconfigure(errors='replace')

    parser = argparse.ArgumentParser(
        description='Catch ribbon tools and usage tracking lost when copying '
                    'from t3lab_dev.')
    group = parser.add_mutually_exclusive_group()
    group.add_argument('--staged', action='store_true',
                       help='check the git index instead of the working tree')
    group.add_argument('--update', action='store_true',
                       help='re-record the ribbon folder list in manifest.json')
    group.add_argument('--install-hook', action='store_true',
                       help='run the check on every commit in this clone')
    args = parser.parse_args(argv)

    if args.update:
        return cmd_update()
    if args.install_hook:
        return cmd_install_hook()
    return cmd_check(args.staged)


if __name__ == '__main__':
    sys.exit(main())
