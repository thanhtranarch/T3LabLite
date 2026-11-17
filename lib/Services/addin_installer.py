# -*- coding: utf-8 -*-
"""Install a T3Lab add-in for another Autodesk app from its GitHub release, so
MCP Control's Connect works without a manual setup (owner, 2026-11-17).

Packages (repo thanhtranarch/t3lab_CADmcp, scripts/package.py there):
T3LabCAD (AutoCAD and Civil 3D 2022-2027) and T3LabMax (3ds Max 2022-2027).
Each release zip holds
    <Package>.bundle/...  -> %APPDATA%\\Autodesk\\ApplicationPlugins\\<Package>.bundle
    <Package>/...         -> %APPDATA%\\<Package>  (bridge.py, tools_snapshot.json)
    manifest.json         {"name": "<Package>", "version": ...}
A file the host app holds open (AutoCAD's loader DLL) is skipped and reported:
closing the app and connecting again updates it.

mcp_paths.json "<package>_addin_url" (e.g. "T3LabMax_addin_url") points at
another zip: an https URL or a local file (offline installs).

Engine-agnostic (CPython 3 + IronPython 2.7). Test: lite_guard/test_addin_installer.py.
"""
from __future__ import unicode_literals

import io
import json
import os
import shutil
import tempfile
import zipfile

_RELEASE = 'https://github.com/thanhtranarch/t3lab_CADmcp/releases/latest/download/'
PACKAGES = {
    'T3LabCAD': {
        'url': _RELEASE + 'T3LabCAD.zip',
        'bundle': 'T3LabCAD.bundle',
        'apps': 'AutoCAD and Civil 3D 2022-2027',
        'host': 'AutoCAD / Civil 3D',
        'old': ('T3LabCAD.Core.dll', 'T3LabCAD.Core.pdb',
                'T3LabCAD.Loader.dll', 'T3LabCAD.Loader.pdb'),   # pre-1.1 layout
    },
    'T3LabMax': {
        'url': _RELEASE + 'T3LabMax.zip',
        'bundle': 'T3LabMax.bundle',
        'apps': '3ds Max 2022-2027',
        'host': '3ds Max',
        'old': (),
    },
}
# Folders inside Program Files\Autodesk, per app: (prefix, subfolder that must exist or '').
_PRODUCTS = {
    'autocad': (('AutoCAD 20', ''),),
    'civil3d': (('AutoCAD 20', 'C3D'), ('Civil 3D 20', '')),
    'navisworks': (('Navisworks Manage 20', ''), ('Navisworks Simulate 20', ''),
                   ('Navisworks Freedom 20', '')),
    '3dsmax': (('3ds Max 20', ''),),
}
MIN_YEAR = 2022
TIMEOUT_S = 60


class InstallError(Exception):
    pass


def appdata():
    return os.environ.get('APPDATA') or os.path.join(os.path.expanduser('~'), 'AppData', 'Roaming')


def plugins_dir():
    return os.path.join(appdata(), 'Autodesk', 'ApplicationPlugins')


def installed_years(app, roots=None):
    """Years of `app` installed under Program Files\\Autodesk (>= MIN_YEAR), sorted."""
    if roots is None:
        roots = [os.path.join(os.environ.get(v, ''), 'Autodesk')
                 for v in ('ProgramW6432', 'ProgramFiles') if os.environ.get(v)]
    years = set()
    for root in roots:
        try:
            names = os.listdir(root)
        except OSError:
            continue
        for prefix, sub in _PRODUCTS.get(app, ()):
            for name in names:
                if not name.startswith(prefix):
                    continue
                tail = name[len(prefix) - 2:][:4]
                if not tail.isdigit() or int(tail) < MIN_YEAR:
                    continue
                if sub and not os.path.isdir(os.path.join(root, name, sub)):
                    continue
                years.add(int(tail))
    return sorted(years)


def _safe_members(z, package, bundle):
    """{zip name: (kind, relative path)} for files we install; rejects anything else."""
    out = {}
    for info in z.infolist():
        name = info.filename.replace('\\', '/')
        if name.endswith('/'):
            continue
        parts = name.split('/')
        if name.startswith('/') or '..' in parts or ':' in parts[0]:
            raise InstallError('Unsafe path in the add-in zip: {}'.format(name))
        if parts[0] == bundle and len(parts) > 1:
            out[info.filename] = ('bundle', '/'.join(parts[1:]))
        elif parts[0] == package and len(parts) > 1:
            out[info.filename] = ('data', '/'.join(parts[1:]))
        elif name != 'manifest.json':
            raise InstallError('Unexpected file in the add-in zip: {}'.format(name))
    return out


def _fetch(source, dest):
    """Copy a local zip, or download an https one, to dest."""
    if os.path.isfile(source):
        shutil.copyfile(source, dest)
        return
    if not source.lower().startswith('https://'):
        raise InstallError('Add-in source must be an https URL or a file: {}'.format(source))
    try:
        from urllib.request import Request, urlopen
    except ImportError:                       # IronPython 2.7
        from urllib2 import Request, urlopen
    req = Request(source, headers={'User-Agent': 'T3Lab-MCP-Control'})
    try:
        resp = urlopen(req, timeout=TIMEOUT_S)
        with io.open(dest, 'wb') as f:
            shutil.copyfileobj(resp, f)
    except Exception as ex:
        code = getattr(ex, 'code', None)
        if code == 404:
            raise InstallError('No add-in release published yet ({}).'.format(source))
        raise InstallError('Could not download the add-in: {}'.format(ex))


def install(package, source=None, roots=None):
    """Download and install `package`. Returns a dict:
    version, bundle_dir, data_dir, written [...], locked [...]. Raises InstallError."""
    spec = PACKAGES.get(package)
    if spec is None:
        raise InstallError('No T3Lab add-in package named {}.'.format(package))
    source = source or spec['url']
    bundle_dir = os.path.join(plugins_dir(), spec['bundle'])
    data_dir = os.path.join(appdata(), package)
    tmp = tempfile.mkdtemp(prefix='t3lab_addin_')
    try:
        zpath = os.path.join(tmp, package + '.zip')
        _fetch(source, zpath)
        try:
            z = zipfile.ZipFile(zpath)
        except zipfile.BadZipfile:
            raise InstallError('The downloaded add-in is not a zip file.')
        with z:
            try:
                manifest = json.loads(z.read('manifest.json').decode('utf-8'))
            except KeyError:
                raise InstallError('The add-in zip has no manifest.json.')
            if manifest.get('name') != package:
                raise InstallError('The zip is {}, not {}.'.format(manifest.get('name'), package))
            members = _safe_members(z, package, spec['bundle'])
            if not any(k == 'bundle' and r == 'PackageContents.xml' for k, r in members.values()):
                raise InstallError('The add-in zip has no PackageContents.xml.')
            if not any(k == 'data' and r == 'bridge.py' for k, r in members.values()):
                raise InstallError('The add-in zip has no bridge.py.')
            written, locked = [], []
            for name, (kind, rel) in sorted(members.items()):
                target = os.path.join(bundle_dir if kind == 'bundle' else data_dir,
                                      *rel.split('/'))
                folder = os.path.dirname(target)
                if not os.path.isdir(folder):
                    os.makedirs(folder)
                data = z.read(name)
                try:
                    with io.open(target, 'rb') as f:
                        if f.read() == data:
                            continue
                except (IOError, OSError):
                    pass
                try:
                    with io.open(target, 'wb') as f:
                        f.write(data)
                    written.append(rel)
                except (IOError, OSError):
                    locked.append(rel)          # AutoCAD holds the loader DLL open
        for old in spec['old']:
            try:
                os.remove(os.path.join(bundle_dir, 'Contents', old))
            except OSError:
                pass
        return {'version': manifest.get('version', ''), 'bundle_dir': bundle_dir,
                'data_dir': data_dir, 'written': written, 'locked': locked,
                'apps': spec['apps'], 'host': spec['host']}
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def summary(result):
    """One line for MCP Control's status text."""
    text = 'Installed the T3Lab add-in {} for {}.'.format(
        result.get('version') or '', result.get('apps') or '')
    host = result.get('host') or 'the app'
    if result.get('locked'):
        text += ' {} file(s) are in use: close {} and click Connect again to update them.'.format(
            len(result['locked']), host)
    else:
        text += ' Restart {} if it is open.'.format(host)
    return text.replace('  ', ' ')
