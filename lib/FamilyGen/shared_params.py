# -*- coding: utf-8 -*-
"""Revit shared parameter files, read as text (pure Python, no Revit).

The Studio runs outside Revit's API context, so it reads the file Revit uses
(``Application.SharedParametersFilename``, captured when the Studio opens) as
plain text: tab-separated ``*PARAM`` rows. The builder adds the same
definitions through the API (``FamilyManager.AddParameter(ExternalDefinition,
...)``, FamilyGen/builder.py). docs/famigen-param-profiles.md; tests:
lite_guard/test_famigen_studio.py (SharedParamTests).
"""
import io
import os
import re

GUID = re.compile(r'^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')

# DATATYPE in the file (old names, or a spec id on newer Revit) -> schema type.
_KINDS = (('length', 'length'), ('number', 'number'), ('integer', 'integer'), ('int', 'integer'),
          ('text', 'text'), ('string', 'text'), ('material', 'material'))


def kind_of(datatype):
    """'LENGTH' / 'autodesk.spec.aec:length-2.0.0' -> 'length'; None when unknown."""
    low = u'{}'.format(datatype or u'').strip().lower()
    stem = low.split(':')[-1].split('-')[0] if ':' in low else low
    for key, kind in _KINDS:
        if stem == key:
            return kind
    return None


def parse(text):
    """[{'guid', 'name', 'datatype', 'kind', 'group', 'description'}] from a file's text."""
    groups = {}
    header = {}
    out = []
    for line in (text or u'').splitlines():
        if not line or line.startswith('#'):
            continue
        cells = line.split('\t')
        tag = cells[0]
        if tag.startswith('*'):
            header[tag[1:]] = [c.strip().upper() for c in cells[1:]]
            continue
        cols = header.get(tag)
        if not cols:
            continue
        row = dict(zip(cols, cells[1:]))
        if tag == 'GROUP':
            groups[row.get('ID', '')] = row.get('NAME', '')
        elif tag == 'PARAM' and GUID.match(row.get('GUID', '').strip()) and row.get('NAME', '').strip():
            out.append({'guid': row['GUID'].strip().lower(), 'name': row['NAME'].strip(),
                        'datatype': row.get('DATATYPE', '').strip(),
                        'kind': kind_of(row.get('DATATYPE')),
                        'group': row.get('GROUP', '').strip(),
                        'description': row.get('DESCRIPTION', '').strip()})
    for item in out:
        item['group'] = groups.get(item['group'], item['group'])
    return out


def read(path):
    """parse() of the file at ``path`` ([] when it is missing or unreadable).

    Revit writes these files as UTF-16 (with a BOM) or UTF-8; both are tried."""
    if not path or not os.path.isfile(path):
        return []
    with io.open(path, 'rb') as handle:
        raw = handle.read()
    for encoding in ('utf-16', 'utf-8-sig', 'cp1252') if raw[:2] in (b'\xff\xfe', b'\xfe\xff') \
            else ('utf-8-sig', 'utf-16', 'cp1252'):
        try:
            text = raw.decode(encoding)
        except (UnicodeDecodeError, UnicodeError):
            continue
        if '*PARAM' in text or 'PARAM\t' in text:
            return parse(text)
    return []


def find(definitions, name=None, guid=None, kind=None):
    """The definition with ``guid``, else with ``name`` (case-insensitive) and a
    compatible kind; None when there is none."""
    if guid:
        for item in definitions or ():
            if item['guid'] == u'{}'.format(guid).strip().lower():
                return item
    low = u'{}'.format(name or u'').strip().lower()
    for item in definitions or ():
        if item['name'].lower() == low and (kind is None or item['kind'] in (None, kind)):
            return item
    return None


class Cache(object):
    """Definitions of one file, re-read when the file changes on disk."""

    def __init__(self, path=None):
        self.path = path
        self._stamp = None
        self._defs = []

    def definitions(self):
        try:
            stamp = os.path.getmtime(self.path) if self.path else None
        except OSError:
            stamp = None
        if stamp != self._stamp:
            self._stamp = stamp
            try:
                self._defs = read(self.path) if stamp is not None else []
            except (IOError, OSError):
                self._defs = []
        return self._defs
