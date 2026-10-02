#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Copy files from a t3lab_dev checkout into T3LabLite without losing Lite code.

A safer replacement for copying folders by hand in Explorer:
  - dry run by default: prints what would change, writes nothing
  - never deletes: files only Lite has are listed as KEEP and left alone
  - never copies the paths listed under "lite_only" in manifest.json
  - BLOCKED: a dev file that would remove a Lite-only hook (usage tracking,
    the Lite update repo) or drop a tool from a ribbon layout is not copied;
    merge it by hand, or pass --force once you have
  - after --apply, runs lite_guard/guard.py on the result

Usage:
  python lite_guard/sync_from_dev.py D:/t3lab_dev lib/GUI
  python lite_guard/sync_from_dev.py D:/t3lab_dev lib/GUI lib/core/server.py --apply
  python lite_guard/sync_from_dev.py D:/t3lab_dev "T3Lab.tab/Modeling & Datum.panel/FamiTransfer.pushbutton" --apply

Paths are relative to the repository root (the same in both repos).
"""
from __future__ import print_function

import argparse
import io
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import guard  # noqa: E402

ROOT = guard.ROOT
MAX_KEEP_LINES = 40


def _read(path):
    try:
        with io.open(path, 'rb') as fh:
            return fh.read()
    except (IOError, OSError):
        return None


def _same(a, b):
    """Equal bytes, or equal text once CRLF/LF differences are ignored."""
    if a == b:
        return True
    if a is None or b is None or b'\0' in a[:8192] or b'\0' in b[:8192]:
        return False
    return a.replace(b'\r\n', b'\n') == b.replace(b'\r\n', b'\n')


def _walk(base, rel):
    """Repo-relative posix paths of every file under base/rel."""
    full = os.path.join(base, rel)
    if os.path.isfile(full):
        return [rel.replace(os.sep, '/')]
    found = []
    for dirpath, dirnames, filenames in os.walk(full):
        dirnames[:] = [d for d in dirnames if d not in guard.SKIP_DIRS]
        for name in filenames:
            if name.endswith(('.pyc', '.pyo')):
                continue
            path = os.path.relpath(os.path.join(dirpath, name), base)
            found.append(path.replace(os.sep, '/'))
    return found


def _is_lite_only(path, lite_only):
    return any(path == lo or path.startswith(lo.rstrip('/') + '/')
               for lo in lite_only)


def _block_reasons(path, lite_raw, dev_raw, markers):
    """Why copying the dev file over the Lite one would lose something."""
    reasons = []
    lite_text, dev_text = guard._text(lite_raw), guard._text(dev_raw)

    for marker in markers:
        if marker['file'] != path:
            continue
        if guard.marker_ok(marker, lite_text) and not guard.marker_ok(marker, dev_text):
            reasons.append('drops ' + marker['why'])

    folder, name = path.rsplit('/', 1) if '/' in path else ('', path)
    if name == 'bundle.yaml' and folder.endswith(guard.CONTAINER_EXTS):
        lite_layout = guard.read_layout(lite_text) or []
        dev_layout = guard.read_layout(dev_text) or []
        dropped = [n for n in lite_layout if n not in dev_layout]
        if dropped:
            reasons.append('drops from the ribbon layout: ' + ', '.join(dropped))
    return reasons


def _normalize_rel(dev_dir, path):
    """Accept repo-relative paths, or absolute ones inside dev_dir."""
    if os.path.isabs(path):
        path = os.path.relpath(path, dev_dir)
    path = path.replace('\\', '/').strip('/')
    return '.' if path in ('', '.') else path


def plan(dev_dir, paths, manifest):
    lite_only = manifest.get('lite_only', [])
    markers = manifest.get('markers', [])
    rows = {'NEW': [], 'CHANGED': [], 'BLOCKED': [], 'KEEP': [], 'SKIP': []}
    unchanged = 0
    seen = set()

    for rel in paths:
        if not os.path.exists(os.path.join(dev_dir, rel)):
            raise SystemExit('sync_from_dev: "{}" is not in {}'.format(rel, dev_dir))

        dev_files = set(_walk(dev_dir, rel))
        for path in sorted(dev_files):
            if path in seen:
                continue
            seen.add(path)
            if _is_lite_only(path, lite_only):
                rows['SKIP'].append((path, 'lite_only in manifest.json'))
                continue
            dev_raw = _read(os.path.join(dev_dir, path))
            lite_raw = _read(os.path.join(ROOT, path))
            if lite_raw is None:
                rows['NEW'].append((path, ''))
            elif _same(lite_raw, dev_raw):
                unchanged += 1
            else:
                reasons = _block_reasons(path, lite_raw, dev_raw, markers)
                if reasons:
                    rows['BLOCKED'].append((path, '; '.join(reasons)))
                else:
                    rows['CHANGED'].append((path, ''))

        if os.path.isdir(os.path.join(ROOT, rel)):
            for path in sorted(_walk(ROOT, rel)):
                if path not in dev_files and path not in seen:
                    seen.add(path)
                    rows['KEEP'].append((path, 'only in Lite, left as is'))
    return rows, unchanged


def print_plan(rows, unchanged, dev_dir):
    for kind in ('NEW', 'CHANGED', 'BLOCKED', 'SKIP'):
        for path, note in rows[kind]:
            print('  {:<8} {}{}'.format(kind, path, '  -- ' + note if note else ''))
            if kind == 'BLOCKED':
                print('           merge by hand:  git diff --no-index "{}" "{}"'.format(
                    path, os.path.join(dev_dir, path)))
    keep = rows['KEEP']
    for path, note in keep[:MAX_KEEP_LINES]:
        print('  {:<8} {}  -- {}'.format('KEEP', path, note))
    if len(keep) > MAX_KEEP_LINES:
        print('  KEEP     ... and {} more Lite-only files'.format(
            len(keep) - MAX_KEEP_LINES))
    print('  {} unchanged'.format(unchanged))


def apply(rows, dev_dir, force):
    kinds = ('NEW', 'CHANGED', 'BLOCKED') if force else ('NEW', 'CHANGED')
    count = 0
    for kind in kinds:
        for path, _ in rows[kind]:
            target = os.path.join(ROOT, path)
            parent = os.path.dirname(target)
            if parent and not os.path.isdir(parent):
                os.makedirs(parent)
            shutil.copy2(os.path.join(dev_dir, path), target)
            count += 1
    return count


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            stream.reconfigure(errors='replace')

    parser = argparse.ArgumentParser(
        description='Copy files from t3lab_dev into T3LabLite without losing '
                    'Lite-only code. Dry run unless --apply.')
    parser.add_argument('dev_dir', help='root folder of the t3lab_dev checkout')
    parser.add_argument('paths', nargs='+',
                        help='files or folders to bring over, relative to the '
                             'repo root (use . for everything)')
    parser.add_argument('--apply', action='store_true',
                        help='write NEW and CHANGED files (default: dry run)')
    parser.add_argument('--force', action='store_true',
                        help='with --apply, also overwrite BLOCKED files')
    args = parser.parse_args(argv)

    dev_dir = os.path.abspath(args.dev_dir)
    if not os.path.isdir(dev_dir):
        raise SystemExit('sync_from_dev: {} is not a folder'.format(dev_dir))
    if os.path.normcase(dev_dir) == os.path.normcase(ROOT):
        raise SystemExit('sync_from_dev: dev_dir is this repo -- point it at t3lab_dev')

    manifest = guard.load_manifest()
    paths = [_normalize_rel(dev_dir, p) for p in args.paths]
    rows, unchanged = plan(dev_dir, paths, manifest)

    mode = 'apply' if args.apply else 'dry run -- nothing written, add --apply'
    print('sync_from_dev: {} -> {}  ({})'.format(dev_dir, ROOT, mode))
    print_plan(rows, unchanged, dev_dir)

    if not args.apply:
        return 1 if rows['BLOCKED'] else 0

    written = apply(rows, dev_dir, args.force)
    print('sync_from_dev: wrote {} file(s)'.format(written))
    if rows['BLOCKED'] and not args.force:
        print('sync_from_dev: {} BLOCKED file(s) not copied -- merge them by '
              'hand'.format(len(rows['BLOCKED'])))
    print()
    return guard.cmd_check(staged=False)


if __name__ == '__main__':
    sys.exit(main())
