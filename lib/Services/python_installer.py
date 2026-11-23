# -*- coding: utf-8 -*-
"""A Python 3 for the MCP bridge on a PC that has none (owner, 2026-11-23:
nobody installs or configures anything by hand).

AI apps start %APPDATA%\\T3LabAI\\bridge.py with a CPython 3; Revit's own
engines cannot run it. When no interpreter is found, the newest stable
official build is unpacked to
    %APPDATA%\\T3LabAI\\python\\<version>\\tools\\python.exe
from the Python Software Foundation's NuGet package ("python", x64;
"pythonarm64"; "pythonx86"). It is a plain zip: no installer runs, no admin
right, nothing on PATH, no entry in Windows' app list -- a private copy only
T3Lab's bridge uses. Deleting the folder removes it.

mcp_paths.json "auto_install_python": false turns it off;
"python_package_url" points at another .nupkg (an https URL or a local file,
for offline PCs).

Engine-agnostic (CPython 3 + IronPython 2.7). Test: lite_guard/test_python_installer.py.
"""
from __future__ import unicode_literals

import io
import json
import os
import shutil
import tempfile
import zipfile

_FLAT = 'https://api.nuget.org/v3-flatcontainer/'
# PROCESSOR_ARCHITECTURE -> NuGet package id.
_PACKAGES = {'AMD64': 'python', 'ARM64': 'pythonarm64', 'X86': 'pythonx86'}
TIMEOUT_S = 120
EXE = 'python.exe'


class InstallError(Exception):
    pass


def appdata():
    return os.environ.get('APPDATA') or os.path.join(os.path.expanduser('~'), 'AppData', 'Roaming')


def private_dir():
    return os.path.join(appdata(), 'T3LabAI', 'python')


def supported():
    """Only Windows has the NuGet builds (and only Windows runs Revit)."""
    return os.name == 'nt'


def package_id():
    arch = (os.environ.get('PROCESSOR_ARCHITEW6432')
            or os.environ.get('PROCESSOR_ARCHITECTURE') or 'AMD64').upper()
    return _PACKAGES.get(arch, 'python')


def _version_key(text):
    return tuple(int(p) for p in text.split('.'))


def stable_versions(versions):
    """Final releases only ('3.15.2', not '3.16.0-a2'), oldest first."""
    out = []
    for v in versions:
        try:
            _version_key(v)
        except (ValueError, AttributeError):
            continue
        out.append(v)
    return sorted(out, key=_version_key)


def find(root=None):
    """python.exe of the newest private copy, or ''."""
    root = root or private_dir()
    try:
        names = stable_versions(os.listdir(root))
    except OSError:
        return ''
    for name in reversed(names):
        exe = os.path.join(root, name, 'tools', EXE)
        if os.path.isfile(exe):
            return exe
    return ''


def _download(url, dest):
    """https -> file. .NET first: IronPython 2.7's own ssl cannot reach every
    modern server, and inside Revit both engines have clr."""
    if not url.lower().startswith('https://'):
        raise InstallError('Python source must be an https URL or a file: {}'.format(url))
    try:
        import clr  # noqa: F401
        from System.Net import WebClient, ServicePointManager, SecurityProtocolType
    except Exception:
        WebClient = None
    if WebClient is not None:
        try:
            ServicePointManager.SecurityProtocol = (
                ServicePointManager.SecurityProtocol | SecurityProtocolType.Tls12)
        except Exception:
            pass
        client = WebClient()
        try:
            client.Headers.Add('User-Agent', 'T3Lab-MCP-Control')
            client.DownloadFile(url, dest)
            return
        except Exception as ex:
            raise InstallError('Could not download Python: {}'.format(ex))
        finally:
            client.Dispose()
    try:
        from urllib.request import Request, urlopen
    except ImportError:                       # IronPython 2.7 outside Revit
        from urllib2 import Request, urlopen
    try:
        resp = urlopen(Request(url, headers={'User-Agent': 'T3Lab-MCP-Control'}),
                       timeout=TIMEOUT_S)
        with io.open(dest, 'wb') as f:
            shutil.copyfileobj(resp, f)
    except Exception as ex:
        raise InstallError('Could not download Python: {}'.format(ex))


def latest_version(package=None, tmp=None):
    """Newest final release NuGet lists for the package."""
    package = package or package_id()
    own = tmp is None
    tmp = tmp or tempfile.mkdtemp(prefix='t3lab_py_')
    try:
        index = os.path.join(tmp, 'index.json')
        _download('{}{}/index.json'.format(_FLAT, package), index)
        with io.open(index, 'r', encoding='utf-8-sig') as f:
            versions = stable_versions(json.load(f).get('versions') or [])
    except InstallError:
        raise
    except Exception as ex:
        raise InstallError('Could not read the Python version list: {}'.format(ex))
    finally:
        if own:
            shutil.rmtree(tmp, ignore_errors=True)
    if not versions:
        raise InstallError('No Python release found on NuGet.')
    return versions[-1]


def _nuspec_version(z):
    import re
    for name in z.namelist():
        if name.lower().endswith('.nuspec') and '/' not in name:
            m = re.search(r'<version>\s*([^<\s]+)\s*</version>', z.read(name).decode('utf-8', 'ignore'))
            if m:
                return m.group(1)
    return ''


def install(source=None, root=None):
    """Unpack the newest Python (or `source`, a .nupkg file or https URL) into
    the private folder. Returns python.exe's path. Raises InstallError."""
    root = root or private_dir()
    tmp = tempfile.mkdtemp(prefix='t3lab_py_')
    try:
        package = package_id()
        pkg = os.path.join(tmp, 'python.nupkg')
        if source and os.path.isfile(source):
            shutil.copyfile(source, pkg)
        elif source:
            _download(source, pkg)
        else:
            version = latest_version(package, tmp)
            _download('{0}{1}/{2}/{1}.{2}.nupkg'.format(_FLAT, package, version), pkg)
        try:
            z = zipfile.ZipFile(pkg)
        except zipfile.BadZipfile:
            raise InstallError('The downloaded Python package is not a zip file.')
        with z:
            version = _nuspec_version(z)
            if not stable_versions([version]):
                raise InstallError('The Python package has no usable version ({}).'.format(version))
            members = []
            for info in z.infolist():
                name = info.filename.replace('\\', '/')
                parts = name.split('/')
                if name.endswith('/') or parts[0] != 'tools':
                    continue
                if '..' in parts or ':' in name or name.startswith('/'):
                    raise InstallError('Unsafe path in the Python package: {}'.format(name))
                members.append((info, parts))
            if not any(parts == ['tools', EXE] for _i, parts in members):
                raise InstallError('The Python package has no tools/python.exe.')
            target = os.path.join(root, version)
            staging = target + '.t3lab-tmp'
            shutil.rmtree(staging, ignore_errors=True)
            for info, parts in members:
                dest = os.path.join(staging, *parts)
                folder = os.path.dirname(dest)
                if not os.path.isdir(folder):
                    os.makedirs(folder)
                with io.open(dest, 'wb') as f:
                    f.write(z.read(info))
        # The folder appears complete or not at all: find() never sees half a copy.
        try:
            shutil.rmtree(target, ignore_errors=True)
            os.rename(staging, target)
        except OSError as ex:
            shutil.rmtree(staging, ignore_errors=True)
            raise InstallError('Could not put Python in {}: {}'.format(target, ex))
        return os.path.join(target, 'tools', EXE)
    except InstallError:
        raise
    except Exception as ex:
        raise InstallError('Python was not installed: {}'.format(ex))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
