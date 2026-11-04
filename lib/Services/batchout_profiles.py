# -*- coding: utf-8 -*-
"""BatchOut profile files: read, validate, list and name them on disk.

Revit-free; tested by `dev/test_batchout_report.py`.
"""

import json
import os
import re

PROFILE_EXT = '.json'
# Bookkeeping files in the profiles folder (latest setup, crash history) start with '_'.
_RESERVED_PREFIX = '_'


def read_profile_file(path):
    """Profile dict from a JSON file; ValueError when it is not a BatchOut profile."""
    with open(path, 'r', encoding='utf-8-sig') as f:
        try:
            data = json.load(f)
        except ValueError as ex:
            raise ValueError('not valid JSON ({})'.format(ex))
    if not isinstance(data, dict) or not str(data.get('Name') or '').strip():
        raise ValueError('not a BatchOut profile (no "Name" field)')
    return data


def safe_filename(name):
    """Windows-safe file stem for a profile name."""
    stem = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', str(name or '')).strip(' .')
    stem = stem.lstrip(_RESERVED_PREFIX)
    return stem[:80] or 'Profile'


def unique_profile_path(folder, name):
    """Free path in `folder` for `name`: 'X.json', 'X (2).json', ..."""
    stem = safe_filename(name)
    candidate = os.path.join(folder, stem + PROFILE_EXT)
    counter = 2
    while os.path.exists(candidate):
        candidate = os.path.join(folder, '{} ({}){}'.format(stem, counter, PROFILE_EXT))
        counter += 1
    return candidate


def unique_profile_name(existing_names, name):
    """`name`, or 'name (2)', 'name (3)', ... when the picker already shows it."""
    taken = set(str(n).strip().lower() for n in existing_names)
    base = str(name or '').strip() or 'Profile'
    candidate, counter = base, 2
    while candidate.lower() in taken:
        candidate = '{} ({})'.format(base, counter)
        counter += 1
    return candidate


def list_profiles(folder):
    """[(name, path, data)] of valid profiles in `folder`, sorted by name; bad files skipped."""
    found = []
    if not folder or not os.path.isdir(folder):
        return found
    for filename in os.listdir(folder):
        if filename.startswith(_RESERVED_PREFIX) or not filename.lower().endswith(PROFILE_EXT):
            continue
        path = os.path.join(folder, filename)
        try:
            data = read_profile_file(path)
        except (OSError, ValueError):
            continue
        found.append((str(data['Name']).strip(), path, data))
    found.sort(key=lambda entry: entry[0].lower())
    return found


def write_profile_file(path, data):
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    return path
