# -*- coding: utf-8 -*-
"""Recipes: a team-saved sequence of tool calls, run locally as one tool.

Payload (what a reviewer publishes, locked by sha256 in tool_versions):
    {"kind": "recipe", "name", "description", "input_schema",
     "is_write", "is_destructive",
     "plan": {"steps": [{"tool": "set_parameter", "args": {...}}, ...]},
     "inputs_schema": {"type": "object", "properties": {...}, "required": [...]}}

Templates inside step args:
    {{input.level}}                    an input value
    {{steps[0].result.ids}}            a field of an earlier step's result
    {{steps[1].result.items[0].id}}    dotted path, list index allowed
A string that is exactly one template keeps the value's type (a list stays a
list); a template inside a longer string is formatted into it.

Every step runs through the caller's execute_tool — in the Assistant that is
the gated _exec_tool, so ToolGate and the audit log see each call as usual.
A published recipe runs only when the sha256 of its payload matches the one
the reviewer approved (jsonb_text() reproduces PostgreSQL's jsonb::text).
"""
from __future__ import unicode_literals

import hashlib
import json
import re

RECIPE_PREFIX = 'recipe_'
# Never inside a recipe: no recursion, no catalog search, no UI launcher.
FORBIDDEN_STEP_TOOLS = frozenset(['find_tools', 'remember_fact', 'open_t3lab_tool',
                                  'search_past_chats'])
MAX_STEPS = 25

_TEMPLATE = re.compile(r'\{\{\s*([^{}]+?)\s*\}\}')
_PATH_TOKEN = re.compile(r'([A-Za-z_][A-Za-z0-9_]*)|\[(\d+)\]|\.(\d+)')

try:
    _text_types = (str, unicode)  # noqa: F821  (IronPython 2.7)
except NameError:
    _text_types = (str,)


class RecipeError(Exception):
    pass


# ── canonical text + hash (must equal PostgreSQL jsonb::text) ──────────────

def _key_order(k):
    b = k.encode('utf-8')
    return (len(b), b)


def jsonb_text(value):
    """PostgreSQL's jsonb output: keys by (byte length, bytes), ', ' and ': '."""
    if isinstance(value, dict):
        items = sorted(value.items(), key=lambda kv: _key_order(kv[0]))
        return '{' + ', '.join(json.dumps(k, ensure_ascii=False) + ': ' + jsonb_text(v)
                               for k, v in items) + '}'
    if isinstance(value, (list, tuple)):
        return '[' + ', '.join(jsonb_text(v) for v in value) + ']'
    if value is True:
        return 'true'
    if value is False:
        return 'false'
    if value is None:
        return 'null'
    if isinstance(value, _text_types):
        return json.dumps(value, ensure_ascii=False)
    return '{}'.format(value)          # int / float / Decimal as written


def payload_sha(payload):
    return hashlib.sha256(jsonb_text(payload).encode('utf-8')).hexdigest()


def verified_payload(row):
    """The published payload when its hash matches, else None."""
    payload, sha = row.get('payload'), row.get('sha256')
    if not isinstance(payload, dict) or not sha:
        return None
    return payload if payload_sha(payload) == sha else None


# ── templates ──────────────────────────────────────────────────────────────

def _resolve(path, inputs, results):
    tokens = []
    for name, idx, didx in _PATH_TOKEN.findall(path):
        tokens.append(name if name else int(idx or didx))
    if not tokens:
        raise RecipeError('Empty template {{}}.')
    head = tokens.pop(0)
    if head == 'input':
        cur = inputs
    elif head == 'steps':
        cur = results
    else:
        raise RecipeError('Unknown template root "{}" (use input. or steps[i].)'.format(head))
    for t in tokens:
        try:
            cur = cur[t]
        except (KeyError, IndexError, TypeError):
            raise RecipeError('Template {{{{{}}}}} has no value at "{}".'.format(path, t))
    return cur


def render(value, inputs, results=()):
    """Substitute templates in args (recursively). results[i] = {"result": res_i}."""
    if isinstance(value, dict):
        return dict((k, render(v, inputs, results)) for k, v in value.items())
    if isinstance(value, list):
        return [render(v, inputs, results) for v in value]
    if isinstance(value, _text_types):
        whole = _TEMPLATE.match(value)
        if whole and whole.end() == len(value):
            return _resolve(whole.group(1), inputs, list(results))

        def sub(m):
            v = _resolve(m.group(1), inputs, list(results))
            return v if isinstance(v, _text_types) else json.dumps(v, ensure_ascii=False)
        return _TEMPLATE.sub(sub, value)
    return value


# ── run ────────────────────────────────────────────────────────────────────

def _check_inputs(schema, inputs):
    schema = schema or {}
    missing = [k for k in (schema.get('required') or []) if inputs.get(k) in (None, '')]
    if missing:
        raise RecipeError('Missing input: {}.'.format(', '.join(missing)))
    out = dict(inputs)
    for k, spec in (schema.get('properties') or {}).items():
        if k not in out and isinstance(spec, dict) and 'default' in spec:
            out[k] = spec['default']
    return out


def run(row, inputs, execute_tool, confirm=None, require_verified=True):
    """Run a recipe row (catalog row or a draft with recipe.plan).

    confirm(summary) -> bool is asked ONCE before a write recipe starts.
    Returns {"success", "steps": [...], "result"} or {"error", "step", "steps"}.
    """
    payload = verified_payload(row)
    if payload is None:
        if require_verified:
            return {"error": u"Recipe {} is not published or its content does not match "
                             u"what the reviewer approved; it was not run.".format(row.get('slug'))}
        payload = {'plan': (row.get('recipe') or {}).get('plan') or row.get('plan'),
                   'inputs_schema': (row.get('recipe') or {}).get('inputs_schema')
                   or row.get('input_schema'),
                   'is_write': row.get('is_write', True)}
    steps = ((payload.get('plan') or {}).get('steps')) or []
    if not steps:
        return {"error": u"Recipe {} has no steps.".format(row.get('slug'))}
    if len(steps) > MAX_STEPS:
        return {"error": u"Recipe has {} steps; the limit is {}.".format(len(steps), MAX_STEPS)}
    bad = [s.get('tool') for s in steps
           if s.get('tool') in FORBIDDEN_STEP_TOOLS
           or (s.get('tool') or '').startswith(RECIPE_PREFIX)]
    if bad:
        return {"error": u"Recipe uses tools a recipe may not call: {}.".format(', '.join(bad))}
    try:
        inputs = _check_inputs(payload.get('inputs_schema'), dict(inputs or {}))
    except RecipeError as ex:
        return {"error": u'{}'.format(ex)}
    if payload.get('is_write', True) and confirm is not None:
        summary = u'Run recipe "{}" ({} steps: {})?'.format(
            row.get('name') or row.get('slug'), len(steps),
            u', '.join(s.get('tool') for s in steps))
        if not confirm(summary):
            return {"cancelled": True, "note": u"Recipe not run: the user said no."}
    done, last = [], None
    for i, step in enumerate(steps):
        try:
            args = render(step.get('args') or {}, inputs, done)
        except RecipeError as ex:
            return {"error": u'Step {} ({}): {}'.format(i + 1, step.get('tool'), ex),
                    "step": i, "steps": _brief(done, steps)}
        res = execute_tool(step.get('tool'), args)
        if not isinstance(res, dict):
            res = {"result": res}
        done.append({"result": res})
        last = res
        if res.get('error') or res.get('cancelled'):
            return {"error": u'Step {} ({}) stopped the recipe: {}'.format(
                        i + 1, step.get('tool'), res.get('error') or 'cancelled'),
                    "step": i, "steps": _brief(done, steps)}
    return {"success": True, "steps": _brief(done, steps), "result": last}


def _brief(done, steps):
    return [{"tool": steps[i].get('tool'), "ok": not (d['result'].get('error'))}
            for i, d in enumerate(done)]


# ── recipes as tools ───────────────────────────────────────────────────────

def recipe_tools(rows):
    """Registry-shape tools for the published, verified recipes in `rows`."""
    out = []
    for row in rows or []:
        if row.get('kind') != 'recipe' or row.get('status') != 'published':
            continue
        payload = verified_payload(row)
        if payload is None:
            continue
        n = len(((payload.get('plan') or {}).get('steps')) or [])
        title = row.get('title_vi') or ''
        desc = u'{}{} (team recipe, {} steps, runs locally)'.format(
            (title + u' — ') if title else u'', row.get('description') or row.get('name'), n)
        out.append({'name': RECIPE_PREFIX + row['slug'], 'description': desc,
                    'inputSchema': payload.get('inputs_schema')
                    or {'type': 'object', 'properties': {}}})
    return out


# ── "Save as tool" ─────────────────────────────────────────────────────────

_SECRET_PATTERNS = [
    (re.compile(r'\bsk-[A-Za-z0-9_\-]{10,}'), 'an API key'),
    (re.compile(r'\bsb_(?:secret|publishable)_[A-Za-z0-9_\-]{8,}'), 'a Supabase key'),
    (re.compile(r'\beyJ[A-Za-z0-9_\-]{8,}\.[A-Za-z0-9_\-]{8,}\.[A-Za-z0-9_\-]{8,}'), 'a token'),
    (re.compile(r'^[A-Za-z]:[\\/]'), 'an absolute file path'),
    (re.compile(r'^\\\\[^\\]+\\'), 'a network path'),
    (re.compile(r'[^@\s"\']+@[^@\s"\']+\.[A-Za-z]{2,}'), 'an e-mail address'),
]


def scan_secrets(value, where='args'):
    """[(where, what)] for values that must not be published in a recipe."""
    found = []
    if isinstance(value, dict):
        for k, v in value.items():
            found.extend(scan_secrets(v, '{}.{}'.format(where, k)))
    elif isinstance(value, list):
        for i, v in enumerate(value):
            found.extend(scan_secrets(v, '{}[{}]'.format(where, i)))
    elif isinstance(value, _text_types):
        if _TEMPLATE.search(value) and _TEMPLATE.sub('', value).strip() == '':
            return found
        for rx, what in _SECRET_PATTERNS:
            if rx.search(value):
                found.append((where, what))
                break
    return found


def _input_name(key, taken):
    base = re.sub(r'[^a-z0-9_]+', '_', (key or 'value').lower()).strip('_') or 'value'
    name, n = base, 2
    while name in taken:
        name, n = '{}_{}'.format(base, n), n + 1
    return name


def build_from_calls(calls, user_text=u''):
    """Turn a turn's successful calls into {steps, inputs_schema}.

    calls: [(tool, args, ok)]. A string argument the user typed in their
    message becomes an input ({{input.x}}), so the recipe works for the next
    level / category / value too.
    """
    text = (user_text or u'').lower()
    steps, props, taken = [], {}, set()
    for tool, args, ok in calls or []:
        if not ok or not tool or tool in FORBIDDEN_STEP_TOOLS or tool.startswith(RECIPE_PREFIX):
            continue
        new_args = {}
        for k, v in (args or {}).items():
            if isinstance(v, _text_types) and len(v.strip()) >= 2 and v.lower() in text:
                existing = [n for n, p in props.items() if p.get('example') == v]
                name = existing[0] if existing else _input_name(k, taken)
                if not existing:
                    taken.add(name)
                    props[name] = {'type': 'string', 'description': k, 'example': v}
                new_args[k] = '{{input.%s}}' % name
            else:
                new_args[k] = v
        steps.append({'tool': tool, 'args': new_args})
    schema = {'type': 'object',
              'properties': dict((n, {'type': 'string', 'description': p['description'],
                                      'default': p['example']}) for n, p in props.items()),
              'required': []}
    return {'steps': steps, 'inputs_schema': schema}
