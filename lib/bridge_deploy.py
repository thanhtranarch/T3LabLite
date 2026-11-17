# -*- coding: utf-8 -*-
"""Copy core/bridge.py and the files it reads next to itself to
%APPDATA%\\T3LabAI, where every MCP client config points.

Two callers:
  - MCPService.deploy_bridge (MCP Control, Connect): always deploys and takes
    ownership of the deployed bridge.
  - startup.py at every Revit start: refresh() updates an already deployed
    bridge, so a new tool (e.g. t3lab_hub) reaches the clients without
    opening MCP Control. It never deploys a first time, and never overwrites a
    bridge another extension on this machine deployed (T3Lab and T3Lab Lite
    share the folder): the owner is the extension whose MCP Control or
    Connect deployed last (bridge_source.txt).

A running client keeps the bridge it started with until it restarts.
Engine-agnostic (IronPython 2.7 + CPython 3). Test: lite_guard/test_bridge_deploy.py.
"""
from __future__ import unicode_literals

import io
import os

BRIDGE = 'bridge.py'
# (lib/core name, deployed name). skills.json -> t3lab_capabilities;
# mcp_providers.py -> external providers; hub_store.py -> t3lab_hub. A missing
# source is skipped: the bridge runs without it, feature off.
COMPANIONS = (('skills_registry.json', 'skills.json'),
              ('mcp_providers.py', 'mcp_providers.py'),
              ('hub_store.py', 'hub_store.py'))
SOURCE_FILE = 'bridge_source.txt'


def core_dir():
    """lib/core of this extension."""
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'core')


def target_dir():
    """%APPDATA%\\T3LabAI (= core.paths.settings_dir), created on demand."""
    base = os.environ.get('APPDATA') or os.path.expanduser('~')
    return os.path.join(base, 'T3LabAI')


def _read(path):
    try:
        with io.open(path, 'rb') as f:
            return f.read()
    except (IOError, OSError):
        return None


def _copy(src, dst):
    """True when dst was written (missing or different), False when equal."""
    data = _read(src)
    if data is None:
        raise IOError('missing source: {}'.format(src))
    if _read(dst) == data:
        return False
    tmp = dst + '.tmp'
    with io.open(tmp, 'wb') as f:
        f.write(data)
    if os.path.exists(dst):
        os.remove(dst)          # IronPython 2.7 has no os.replace
    os.rename(tmp, dst)
    return True


def _same_dir(a, b):
    return os.path.normcase(os.path.abspath(a or '')) == os.path.normcase(os.path.abspath(b or ''))


def owner(target=None):
    """lib/core path that last deployed the bridge in `target`, or ''."""
    data = _read(os.path.join(target or target_dir(), SOURCE_FILE))
    return data.decode('utf-8').strip() if data else ''


def deploy(target=None, source=None):
    """Copy the bridge and its companions; take ownership.

    Returns (bridge path or None, error or None, names written).
    """
    target, source = target or target_dir(), source or core_dir()
    written = []
    try:
        if not os.path.isdir(target):
            os.makedirs(target)
        dst = os.path.join(target, BRIDGE)
        if _copy(os.path.join(source, BRIDGE), dst):
            written.append(BRIDGE)
    except Exception as ex:
        return None, '{}'.format(ex), written
    for src_name, dst_name in COMPANIONS:
        try:
            if _copy(os.path.join(source, src_name), os.path.join(target, dst_name)):
                written.append(dst_name)
        except Exception:
            pass
    try:
        if not _same_dir(owner(target), source):
            with io.open(os.path.join(target, SOURCE_FILE), 'w', encoding='utf-8') as f:
                f.write('{}\n'.format(source))
    except Exception:
        pass
    return dst, None, written


def refresh(target=None, source=None):
    """Revit start: update a bridge this extension owns. (state, detail) with
    state in skipped | current | updated | failed. Never raises."""
    target, source = target or target_dir(), source or core_dir()
    try:
        if not os.path.isfile(os.path.join(target, BRIDGE)):
            return 'skipped', 'no bridge deployed yet (MCP Control deploys it)'
        current = owner(target)
        if current and not _same_dir(current, source):
            return 'skipped', 'deployed by another extension: {}'.format(current)
        _path, error, written = deploy(target, source)
        if error:
            return 'failed', error
        return ('updated', ', '.join(written)) if written else ('current', '')
    except Exception as ex:
        return 'failed', '{}'.format(ex)
