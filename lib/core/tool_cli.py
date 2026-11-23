# -*- coding: utf-8 -*-
"""Agent-native CLI over the T3Lab tool manifest + the argument contract shared
by the bridge (MCP / CLI) and the T3Lab Assistant.

Pattern from HKUDS/CLI-Anything (Apache-2.0, ideas only, no code): `--help`-style
discovery instead of loading every schema, typed `--flag=value` arguments,
token-lean output, and "never trust exit 0" checks of exported files.
Stdlib only: deployed next to bridge.py (lib/bridge_deploy.py) and imported
by Intelligence/agent_loop.py. Doc: docs/agents/cli.md. Test: dev/test_tool_cli.py.
"""

from __future__ import unicode_literals

import difflib
import fnmatch
import glob
import io
import json
import os
import re
import time

PIN_ARGS = ('revit_document', 'revit_port')     # bridge pins, stripped before Revit
# Flags the bridge CLI owns; a tool parameter with one of these names goes in --json.
GLOBAL_FLAGS = ('json', 'json-file', 'agent', 'tier', 'turbo', 'port', 'out', 'pick')
NEAR_MISS = 0.8          # difflib ratio for "did you mean"


# ─── Schema ───

def schema_of(tools, name):
    """inputSchema of `name` from MCP (inputSchema), Anthropic (input_schema)
    or OpenAI/Ollama (function.parameters) tool lists; None if unknown."""
    for t in tools or []:
        if not isinstance(t, dict):
            continue
        fn = t.get('function') if isinstance(t.get('function'), dict) else None
        tname = t.get('name') or (fn or {}).get('name')
        if tname != name:
            continue
        for schema in (t.get('inputSchema'), t.get('input_schema'),
                       (fn or {}).get('parameters')):
            if isinstance(schema, dict):
                return schema
        return {}
    return None


def _props(schema):
    props = (schema or {}).get('properties')
    return props if isinstance(props, dict) else {}


def _type_of(spec):
    t = (spec or {}).get('type')
    if isinstance(t, list):
        t = next((x for x in t if x != 'null'), None)
    return t or ''


def params_line(schema, limit=16):
    """'a (required, string); b (optional, integer, enum: x|y)' for error hints."""
    required = set((schema or {}).get('required') or [])
    parts = []
    for pname, spec in [x for x in _props(schema).items() if x[0] not in PIN_ARGS][:limit]:
        spec = spec if isinstance(spec, dict) else {}
        bits = [_type_of(spec) or 'any']
        if spec.get('enum'):
            bits.append('enum: ' + '|'.join('{}'.format(v) for v in spec['enum'][:8]))
        parts.append('{} ({}, {})'.format(pname, 'required' if pname in required
                                          else 'optional', ', '.join(bits)))
    return '; '.join(parts)


def arg_problems(schema, args, strict=False, ignore=PIN_ARGS):
    """What would make this call wrong before it reaches Revit.

    Always: missing required parameters, and unknown names that are a near
    miss of a real one (a typo the tool would silently ignore). strict also
    reports every other unknown name (CLI flags). Schemas with no properties
    or additionalProperties are not checked for unknown names.
    """
    props = _props(schema)
    if not props:
        return []
    args = args if isinstance(args, dict) else {}
    out = []
    for req in (schema.get('required') or []):
        if req in props and args.get(req) in (None, ''):
            out.append({'kind': 'missing', 'param': req})
    if schema.get('additionalProperties') not in (None, False):
        return out
    free = [p for p in props if p not in args]
    for key in args:
        if key in props or key in ignore or key.startswith('_'):
            continue
        near = difflib.get_close_matches(key, free, n=1, cutoff=NEAR_MISS)
        if near:
            out.append({'kind': 'unknown', 'param': key, 'did_you_mean': near[0]})
        elif strict:
            out.append({'kind': 'unknown', 'param': key})
    return out


def problems_text(name, problems):
    """One sentence an agent can act on."""
    bits = []
    for p in problems:
        if p['kind'] == 'missing':
            bits.append('missing required "{}"'.format(p['param']))
        elif p.get('did_you_mean'):
            bits.append('unknown "{}" (did you mean "{}"?)'.format(p['param'], p['did_you_mean']))
        else:
            bits.append('unknown "{}"'.format(p['param']))
    return 'Arguments for {}: {}. Nothing ran.'.format(name, '; '.join(bits))


# ─── Flags -> arguments ───

_TRUE = ('1', 'true', 'yes', 'on')
_FALSE = ('0', 'false', 'no', 'off')


def _coerce(raw, spec):
    """String from the command line -> the schema's type. ValueError if not."""
    t = _type_of(spec)
    if t == 'integer':
        return int(raw)
    if t == 'number':
        f = float(raw)
        return int(f) if f.is_integer() and '.' not in raw and 'e' not in raw.lower() else f
    if t == 'boolean':
        low = raw.strip().lower()
        if low in _TRUE:
            return True
        if low in _FALSE:
            return False
        raise ValueError('expected true/false')
    if t == 'array':
        if raw.lstrip().startswith('['):
            value = json.loads(raw)
            if not isinstance(value, list):
                raise ValueError('expected a list')
            return value
        item = (spec or {}).get('items') if isinstance((spec or {}).get('items'), dict) else {}
        return [_coerce(x.strip(), item) for x in raw.split(',') if x.strip()]
    if t == 'object':
        value = json.loads(raw)
        if not isinstance(value, dict):
            raise ValueError('expected a JSON object')
        return value
    if not t and raw[:1] in ('[', '{'):
        return json.loads(raw)
    return raw


def parse_flags(tokens, schema):
    """--name=value / --flag / --no-flag after the tool name -> (args, errors).

    Global flags (GLOBAL_FLAGS) are skipped. Repeating an array flag appends.
    Values are typed by the schema and checked against enum.
    """
    props = _props(schema)
    args, errors = {}, []
    skip_next = False
    for tok in tokens:
        if skip_next:
            skip_next = False
            continue
        if not tok.startswith('--') or tok == '--':
            continue
        body = tok[2:]
        key, has_value, raw = body.partition('=')
        if key in GLOBAL_FLAGS:
            skip_next = key == 'json' and not has_value
            continue
        if not has_value:
            if key.startswith('no-') and _type_of(props.get(key[3:])) == 'boolean':
                key, raw = key[3:], 'false'
            else:
                raw = 'true'
        key_py = key if key in props else key.replace('-', '_')
        spec = props.get(key_py)
        if spec is None:
            args[key_py] = raw          # reported by arg_problems(strict=True)
            continue
        try:
            value = _coerce(raw, spec if isinstance(spec, dict) else {})
        except ValueError as ex:
            errors.append('--{}: {} ({})'.format(key, ex, _type_of(spec) or 'value'))
            continue
        enum = spec.get('enum') if isinstance(spec, dict) else None
        if enum and value not in enum:
            errors.append('--{}: "{}" not in {}'.format(key, raw, '|'.join(
                '{}'.format(v) for v in enum)))
            continue
        if _type_of(spec) == 'array' and isinstance(args.get(key_py), list):
            args[key_py].extend(value)
        else:
            args[key_py] = value
    return args, errors


# ─── Discovery ───

def _first_sentence(text, width=110):
    text = ' '.join('{}'.format(text or '').split())
    m = re.match(r'(.+?[.!?])(\s|$)', text)
    text = m.group(1) if m else text
    return text if len(text) <= width else text[:width - 1].rstrip() + '…'


_WORD = re.compile(r'[a-z0-9]+')


def search(tools, query='', limit=15):
    """Ranked 'name - first sentence' lines; no query = every name, one line."""
    named = [t for t in tools or [] if isinstance(t, dict) and t.get('name')]
    words = _WORD.findall((query or '').lower())
    if not words:
        return {'count': len(named), 'tools': ', '.join(sorted(t['name'] for t in named)),
                'next': 'tools <words> to search, help <tool> for parameters'}
    scored = []
    for t in named:
        name = t['name'].lower()
        desc = '{}'.format(t.get('description') or '').lower()
        score = 0
        for w in words:
            stem = w.rstrip('s') if len(w) > 3 else w
            if stem in name:
                score += 3
            if stem in desc:
                score += 1
        if score:
            scored.append((-score, t['name'], t))
    scored.sort(key=lambda x: (x[0], x[1]))
    hits = ['{} — {}'.format(n, _first_sentence(t.get('description')))
            for _s, n, t in scored[:max(1, int(limit))]]
    return {'query': query, 'matches': len(scored), 'tools': hits,
            'next': 'help <tool> for parameters'}


def _flag_hint(spec):
    t = _type_of(spec)
    if t == 'array':
        item = _type_of((spec or {}).get('items') if isinstance((spec or {}).get('items'), dict)
                        else {}) or 'value'
        return '<{0},{0},...>'.format(item)
    if t == 'boolean':
        return ''
    return '<{}>'.format(t or 'value')


def help_text(tool, prog='bridge.py'):
    """--help page for one tool: usage, description, one line per parameter."""
    name = tool.get('name')
    schema = tool.get('inputSchema') or tool.get('input_schema') or {}
    required = set(schema.get('required') or [])
    lines = ['usage: {} call {} [--param=value ...] [--json=\'{{...}}\'] '
             '[--out=compact|pretty|tsv] [--pick=a,b]'.format(prog, name), '']
    desc = ' '.join('{}'.format(tool.get('description') or '').split())
    if desc:
        lines += [desc, '']
    ann = tool.get('annotations') if isinstance(tool.get('annotations'), dict) else {}
    if ann.get('readOnlyHint'):
        lines += ['read-only', '']
    props = [(p, s) for p, s in _props(schema).items() if p not in PIN_ARGS]
    if not props:
        lines.append('no parameters')
    else:
        lines.append('parameters:')
        flags = []
        for p, s in props:
            hint = _flag_hint(s if isinstance(s, dict) else {})
            flags.append('--{}'.format(p) + ('=' + hint if hint else ''))
        width = max(len(f) for f in flags)
        for (p, s), flag in zip(props, flags):
            s = s if isinstance(s, dict) else {}
            extra = []
            if s.get('enum'):
                extra.append('one of ' + '|'.join('{}'.format(v) for v in s['enum']))
            if 'default' in s:
                extra.append('default {}'.format(json.dumps(s['default'], ensure_ascii=False)))
            text = ' '.join('{}'.format(s.get('description') or '').split())
            text = '; '.join([x for x in [text] + extra if x])
            lines.append('  {}  {}  {}'.format(flag.ljust(width + 2),
                                               'required' if p in required else 'optional',
                                               text).rstrip())
    if any(p in GLOBAL_FLAGS for p, _s in props):
        lines += ['', 'parameters named like a global flag go in --json']
    return '\n'.join(lines)


# ─── Output ───

def _rows(value):
    return (isinstance(value, list) and value
            and all(isinstance(r, dict) for r in value))


def _main_list(payload):
    """(key, rows): the payload itself if it is a list of rows, else its
    longest list-of-rows value."""
    if _rows(payload):
        return None, payload
    best = (None, None)
    if isinstance(payload, dict):
        for k, v in payload.items():
            if _rows(v) and len(v) > len(best[1] or []):
                best = (k, v)
    return best


def pick(payload, fields):
    """Keep only `fields` in every row of every list of rows (depth 1)."""
    keep = [f.strip() for f in fields if f.strip()]
    if not keep:
        return payload

    def cut(rows):
        return [dict((k, r[k]) for k in keep if k in r) for r in rows]
    if _rows(payload):
        return cut(payload)
    if isinstance(payload, dict):
        return dict((k, cut(v) if _rows(v) else v) for k, v in payload.items())
    return payload


def _cell(value):
    if value is None:
        return ''
    if isinstance(value, (dict, list)):
        value = json.dumps(value, ensure_ascii=False, separators=(',', ':'))
    return '{}'.format(value).replace('\t', ' ').replace('\r', ' ').replace('\n', ' ')


def render(payload, fmt='pretty'):
    """pretty (indent 2) | compact (one line) | tsv (header once, one row per
    line: no repeated keys; the other top-level keys as '# key: value')."""
    if fmt == 'compact':
        return json.dumps(payload, ensure_ascii=False, separators=(',', ':'))
    if fmt == 'tsv':
        key, rows = _main_list(payload)
        if rows:
            cols = []
            for r in rows:
                cols += [c for c in r if c not in cols]
            lines = []
            if isinstance(payload, dict):
                lines += ['# {}: {}'.format(k, _cell(v)) for k, v in payload.items()
                          if k != key]
            lines.append('\t'.join(cols))
            lines += ['\t'.join(_cell(r.get(c)) for c in cols) for r in rows]
            return '\n'.join(lines)
        return json.dumps(payload, ensure_ascii=False, separators=(',', ':'))
    return json.dumps(payload, ensure_ascii=False, indent=2)


# ─── Output verification ("never trust exit 0") ───

_OLE = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
_ZIP = b'PK\x03\x04'
# extension -> (kind, check(head bytes) -> bool); None = existence + size only.
SIGNATURES = {
    'pdf': ('PDF', lambda h: h.startswith(b'%PDF-')),
    'dwg': ('DWG', lambda h: h.startswith(b'AC1')),        # AC1009 .. AC1032
    'dxf': ('DXF', lambda h: b'SECTION' in h[:256] or h.startswith(b'AutoCAD Binary DXF')),
    'dgn': ('DGN', lambda h: h.startswith(_OLE) or h[:2] in (b'\x08\x09', b'\xc8\x09')),
    'ifc': ('IFC', lambda h: h.lstrip(b'\xef\xbb\xbf \r\n\t').startswith(b'ISO-10303-21')),
    'ifczip': ('IFC zip', lambda h: h.startswith(_ZIP)),
    'ifcxml': ('IFC XML', lambda h: h.lstrip(b'\xef\xbb\xbf \r\n\t').startswith(b'<')),
    'rvt': ('Revit', lambda h: h.startswith(_OLE)),
    'rfa': ('Revit family', lambda h: h.startswith(_OLE)),
    'rte': ('Revit template', lambda h: h.startswith(_OLE)),
    'rft': ('Revit family template', lambda h: h.startswith(_OLE)),
    'xlsx': ('Excel', lambda h: h.startswith(_ZIP)),
    'docx': ('Word', lambda h: h.startswith(_ZIP)),
    'zip': ('ZIP', lambda h: h.startswith(_ZIP)),
    'png': ('PNG', lambda h: h.startswith(b'\x89PNG\r\n\x1a\n')),
    'jpg': ('JPEG', lambda h: h.startswith(b'\xff\xd8\xff')),
    'jpeg': ('JPEG', lambda h: h.startswith(b'\xff\xd8\xff')),
    'tif': ('TIFF', lambda h: h[:4] in (b'II*\x00', b'MM\x00*')),
    'tiff': ('TIFF', lambda h: h[:4] in (b'II*\x00', b'MM\x00*')),
    'bmp': ('BMP', lambda h: h.startswith(b'BM')),
    'xml': ('XML', lambda h: h.lstrip(b'\xef\xbb\xbf \r\n\t').startswith(b'<')),
    'html': ('HTML', lambda h: b'<' in h[:512]),
    'json': ('JSON', lambda h: h.lstrip(b'\xef\xbb\xbf \r\n\t')[:1] in (b'{', b'[')),
    'csv': ('CSV', None), 'txt': ('text', None), 'nwc': ('Navisworks cache', None),
    'fbx': ('FBX', None), 'gbxml': ('gbXML', lambda h: b'<' in h[:512]),
}


def verify_file(path, min_bytes=1):
    """{path, ok, bytes, kind, problem?} for one exported file."""
    out = {'path': path, 'ok': False}
    if not os.path.isfile(path):
        out['problem'] = 'missing'
        return out
    size = os.path.getsize(path)
    out['bytes'] = size
    ext = os.path.splitext(path)[1].lower().lstrip('.')
    kind, check = SIGNATURES.get(ext, (ext.upper() or 'file', None))
    out['kind'] = kind
    if size < max(1, int(min_bytes)):
        out['problem'] = 'empty' if size == 0 else 'smaller than {} bytes'.format(min_bytes)
        return out
    if check is not None:
        with open(path, 'rb') as f:
            head = f.read(512)
        if not check(head):
            out['problem'] = 'content is not {} (wrong header)'.format(kind)
            return out
    elif ext not in SIGNATURES:
        out['note'] = 'type not checked'
    out['ok'] = True
    return out


def verify(paths, exts=None, expect=None, since_min=None, min_bytes=1, now=None,
           max_files=None):
    """Check exported files. `paths`: files, folders (every file directly in
    them) or glob patterns (Windows shells do not expand them). `exts`
    ('pdf,dwg') and `since_min` (written in the last N minutes) filter what a
    folder or pattern found; `expect` = the count the tool reported. ok only
    when every file passes and the count matches."""
    now = time.time() if now is None else now
    files, missing = [], []
    for p in paths:
        if any(ch in p for ch in '*?['):
            hits = sorted(glob.glob(p))
            if not hits:
                missing.append(p)
            files += hits
        elif os.path.isdir(p):
            files += sorted(os.path.join(p, n) for n in os.listdir(p)
                            if os.path.isfile(os.path.join(p, n)))
        else:
            files.append(p)
    if exts:
        files = _match_ext(files, exts)
    if since_min is not None:
        cutoff = now - float(since_min) * 60
        files = [f for f in files if not os.path.isfile(f) or os.path.getmtime(f) >= cutoff]
    capped = max_files is not None and len(files) > max_files
    if capped:
        files = files[:max_files]
    seen, checked = set(), []
    for f in files:
        key = os.path.normcase(os.path.abspath(f))
        if key not in seen:
            seen.add(key)
            checked.append(verify_file(f, min_bytes))
    bad = [r for r in checked if not r['ok']]
    out = {'checked': len(checked), 'passed': len(checked) - len(bad), 'failed': len(bad),
           'by_kind': {}, 'problems': bad}
    for r in checked:
        if r['ok']:
            out['by_kind'][r['kind']] = out['by_kind'].get(r['kind'], 0) + 1
    if missing:
        out['no_match'] = missing
    if capped:
        out['capped_at'] = max_files
    if expect is not None:
        out['expected'] = int(expect)
        if len(checked) != int(expect):
            out['count_problem'] = '{} files, the tool reported {}'.format(len(checked), expect)
    out['ok'] = bool(checked) and not bad and not missing and 'count_problem' not in out
    return out


def _match_ext(paths, exts):
    pats = ['*.' + e.strip().lstrip('.').lower() for e in exts.split(',') if e.strip()]
    return [p for p in paths if any(fnmatch.fnmatch(p.lower(), pat) for pat in pats)]


# ─── Skill for agents with a shell (Claude Code, Codex, OpenClaw) ───

SKILL_NAME = 't3lab-cli'
SKILL_MD = '''---
name: t3lab-cli
description: Drive Revit through the T3Lab bridge CLI from a shell (python bridge.py ...). Use for Revit model queries, sheets, views, schedules, exports and edits when the T3Lab MCP tools are not loaded, or to save tokens on long lists. Discover with `tools`/`help`, call with typed flags, verify exported files.
---

# T3Lab CLI

`B` = `python -I "%APPDATA%\\T3LabAI\\bridge.py"` (deployed by MCP Control in Revit).
Revit must be open with T3Lab loaded; `B instances` lists the open windows.

## 1. Discover, do not guess
- `B tools <words>` - ranked tool names + one line each (works with Revit closed, from the cache).
- `B help <tool>` - parameters, types, required, enum values. Read it before the first call of a tool.
- `B capabilities` - your tier, skills, the open documents.

## 2. Call with typed flags
`B call <tool> --name=value --list=a,b --flag --no-flag`
- Values are typed from the schema; a typo or a missing required parameter is refused before Revit runs (exit 2, `bad_arguments`, "did you mean").
- Nested objects: `--json-file=args.json` (PowerShell strips quotes of inline JSON).
- Pin the document on edits: `--revit_document="<title>"`.

## 3. Keep output small
- `--out=tsv` for lists (header once, a row per line), `--out=compact` for the rest.
- `--pick=id,name,number` keeps only these columns.
- Cut results carry `result_id`: `B call t3lab_result --result_id=<id> --key=<list> --offset=<n>`,
  or `--action=export --format=csv` to write the whole list to a file instead of reading it.

## 4. Never trust "success" alone
After any export (PDF, DWG, IFC, NWC, images, Excel):
`B verify "<folder or pattern>" --ext=pdf --since=30 --expect=<count the tool reported>`
It checks each file exists, is not empty and has the right header (%PDF, AC10xx DWG, ISO-10303-21 IFC...).
Report `ok`, the counts and every problem; never say "exported N sheets" without it.

## 5. Exit codes
0 ok - 1 tool or Revit error (`error_code`, `retryable`) - 2 bad arguments, nothing ran.
`revit_busy` / `revit_unavailable`: nothing ran, retry later; never repeat a write that reported success.
'''


def write_skill(folder):
    """<folder>/t3lab-cli/SKILL.md; returns the path."""
    target = os.path.join(folder, SKILL_NAME)
    if not os.path.isdir(target):
        os.makedirs(target)
    path = os.path.join(target, 'SKILL.md')
    with io.open(path, 'w', encoding='utf-8') as f:
        f.write(SKILL_MD)
    return path
