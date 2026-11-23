# -*- coding: utf-8 -*-
"""One declaration per T3Lab tool: the single source for the MCP registry, the
dispatch, tiers, confirmation, Assistant placement, read-back and Tool Hub titles.

A tool lives in `core/tools/<domain>.py` as a ToolSpec in that module's SPECS
list; every module in that folder is collected here. Consumers: core/server.py
(registry, dispatch, _WRITE_TOOLS, is_destructive), Intelligence/tool_schema.py,
Intelligence/agents/specialists.py, core/write_verify.py, Tool Hub titles,
dev/sync_toolhub.py. Domain modules import Revit only inside functions, so this
loads outside Revit and under IronPython 2.7.
Doc: docs/agents/tool-specs.md. Test: dev/test_tool_specs.py.
"""
from __future__ import unicode_literals

import importlib
import os
import sys

EFFECTS = ('read', 'write', 'file')
# Where the Assistant sees the tool (Intelligence/agents/specialists.py subsets,
# tool_schema ESSENTIAL / LOCAL_HIDDEN); 'external' = MCP clients only.
PLACEMENTS = ('read', 'modify', 'export', 'modeling', 'qa', 'multidoc', 'essential', 'hidden',
              'external')
VERIFY_MODES = ('ids', 'must_change', 'self', 'files')
_PACKAGE = 'core.tools'


class ToolSpec(object):
    """What the system needs to know about one tool.

    run(doc, args) -> dict runs on Revit's main thread (or off it for reads,
    main_thread=False). destructive: bool, or (field, values) — destructive
    when args[field] is one of values — or callable(args) -> bool. dry_run
    (default True when the schema has it) never counts as destructive.
    """

    def __init__(self, name, description, run, properties=None, required=(),
                 effect='write', destructive=False, placement=('modify',),
                 verify=(), title_vi='', keywords=(), ribbon=None, main_thread=None,
                 needs_doc=True, background_doc=False, requires=()):
        self.name = name
        self.description = description
        self.run = run
        self.properties = dict(properties or {})
        self.required = list(required)
        self.effect = effect
        self.destructive = destructive
        self.placement = (placement,) if isinstance(placement, str) else tuple(placement)
        self.verify = tuple(verify)
        self.title_vi = title_vi
        self.keywords = list(keywords)
        # One button folder or several ('CropSync.pushbutton', 'DatumSync.pushbutton').
        self.ribbon = (ribbon,) if isinstance(ribbon, str) else tuple(ribbon or ())
        self.main_thread = effect != 'read' if main_thread is None else bool(main_thread)
        self.needs_doc = needs_doc
        # Uses only `doc`: a call pinned to an open, inactive document runs there
        # (server._BACKGROUND_DOC_TOOLS; proved by dev/test_mcp_background_doc.py).
        self.background_doc = bool(background_doc)
        # lib-relative files the tool needs ('Snippets/doc_transfer.py'). A spec
        # whose files are absent is not registered: T3Lab Lite ships fewer tools.
        self.requires = (requires,) if isinstance(requires, str) else tuple(requires or ())

    def missing_requirements(self):
        lib = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        return [r for r in self.requires
                if not os.path.exists(os.path.join(lib, r.replace('/', os.sep)))]

    def schema(self):
        return {'type': 'object', 'properties': dict(self.properties),
                'required': list(self.required)}

    def registry_entry(self):
        return {'name': self.name, 'description': self.description,
                'inputSchema': self.schema()}

    def handle_keys(self):
        """Id arguments that accept "@set" handles: *_id / *_ids whose type or
        item type allows a string (server._expand_id_handles expands them)."""
        out = []
        for name, prop in self.properties.items():
            if not (name.endswith('_id') or name.endswith('_ids')):
                continue
            prop = prop if isinstance(prop, dict) else {}
            kind = (prop.get('items') or {}).get('type') if prop.get('type') == 'array' \
                else prop.get('type')
            kinds = kind if isinstance(kind, list) else [kind]
            if 'string' in kinds:
                out.append(name)
        return sorted(out)

    def is_destructive(self, args):
        args = args if isinstance(args, dict) else {}
        if args.get('dry_run') or not self.destructive:
            return False
        if callable(self.destructive):
            return bool(self.destructive(args))
        if isinstance(self.destructive, tuple):
            field, values = self.destructive
            return '{}'.format(args.get(field) or '').lower() in values
        return True

    def problems(self):
        """Declaration mistakes; [] when the spec is complete."""
        out = []
        if self.effect not in EFFECTS:
            out.append('effect must be one of {}'.format(EFFECTS))
        bad = [p for p in self.placement if p not in PLACEMENTS]
        if bad or not self.placement:
            out.append('placement must be from {}'.format(PLACEMENTS))
        if [v for v in self.verify if v not in VERIFY_MODES]:
            out.append('verify must be from {}'.format(VERIFY_MODES))
        if self.effect != 'read' and not self.verify:
            out.append('a write/file tool needs a read-back mode (verify)')
        if self.effect != 'read' and 'dry_run' not in self.properties:
            out.append('a write/file tool needs a dry_run property (agents preview first)')
        if not callable(self.run):
            out.append('run is not callable')
        if not self.title_vi:
            out.append('title_vi is empty (Tool Hub search in Vietnamese)')
        if [r for r in self.required if r not in self.properties]:
            out.append('required names a missing property')
        if isinstance(self.destructive, tuple):
            field = self.destructive[0]
            enum = (self.properties.get(field) or {}).get('enum') or []
            if [v for v in self.destructive[1] if v not in enum]:
                out.append('destructive values are not in the {} enum'.format(field))
        if len(self.description) > 600:
            out.append('description over 600 chars (token budget)')
        return out


_CACHE = {}
LOAD_ERRORS = []
SKIPPED = {}            # {tool: [missing files]} — declared but not shipped here


def _domain_names():
    folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'tools')
    try:
        names = os.listdir(folder)
    except OSError:
        return []
    return sorted(n[:-3] for n in names if n.endswith('.py') and not n.startswith('_'))


def all_specs():
    """{name: ToolSpec} from every core/tools/<domain>.py, loaded once.
    A module that fails to import is skipped and recorded in LOAD_ERRORS."""
    if 'specs' in _CACHE:
        return _CACHE['specs']
    specs = {}
    for domain in _domain_names():
        try:
            module = importlib.import_module('{}.{}'.format(_PACKAGE, domain))
        except Exception as ex:
            LOAD_ERRORS.append('{}: {}'.format(domain, ex))
            continue
        for spec in getattr(module, 'SPECS', ()):
            if spec.missing_requirements():
                SKIPPED[spec.name] = spec.missing_requirements()
                continue
            if spec.name in specs:
                LOAD_ERRORS.append('{}: duplicate tool {}'.format(domain, spec.name))
                continue
            spec.domain = domain
            specs[spec.name] = spec
    _CACHE['specs'] = specs
    return specs


def get(name):
    return all_specs().get(name)


def names(effect=None, placement=None, verify=None, main_thread=None, background_doc=None):
    out = []
    for s in all_specs().values():
        if background_doc is not None and s.background_doc != background_doc:
            continue
        if effect is not None and s.effect != effect:
            continue
        if placement is not None and placement not in s.placement:
            continue
        if verify is not None and verify not in s.verify:
            continue
        if main_thread is not None and s.main_thread != main_thread:
            continue
        out.append(s.name)
    return frozenset(out)


def registry_entries():
    return dict((n, s.registry_entry()) for n, s in all_specs().items())


def is_destructive(name, args):
    """True / False for a spec tool, None when `name` is not one."""
    spec = get(name)
    return None if spec is None else spec.is_destructive(args)


def titles():
    """Tool Hub title_vi + keywords of every spec tool."""
    return dict((n, {'title_vi': s.title_vi, 'keywords': list(s.keywords)})
                for n, s in all_specs().items())


def prompt_lines():
    """One line per spec tool with an `operation` enum, for the Assistant's
    system prompt (verbs -> tool; destructive ops named). Stable order."""
    out = []
    for name in sorted(all_specs()):
        spec = all_specs()[name]
        ops = (spec.properties.get('operation') or {}).get('enum')
        if not ops:
            continue
        line = '   - `{}` ({}): {}'.format(name, spec.title_vi, ' / '.join(ops))
        if isinstance(spec.destructive, tuple) and spec.destructive[0] == 'operation':
            line += ' — confirm first: {}'.format(', '.join(spec.destructive[1]))
        out.append(line)
    return out


def by_ribbon():
    """{'<Tool>.pushbutton': [tool names]} for buttons with a headless tool."""
    out = {}
    for n, s in all_specs().items():
        for button in s.ribbon:
            out.setdefault(button, []).append(n)
    return out


def run(name, doc, args):
    """Run a spec tool. Errors come back as {'error', 'tool'}, never raised:
    a .NET exception escaping Revit's ExternalEvent kills the session."""
    spec = get(name)
    if spec is None:
        return {'error': 'Unknown tool {}.'.format(name), 'tool': name}
    if spec.needs_doc and doc is None:
        return {'error': 'No active document. Open a project in Revit first.', 'tool': name}
    try:
        result = spec.run(doc, dict(args or {}))
    except Exception as ex:
        text = '{}'.format(getattr(ex, 'Message', None) or ex).strip()
        return {'error': text.splitlines()[0] if text else ex.__class__.__name__,
                'tool': name}
    if not isinstance(result, dict):
        result = {'result': result}
    return result


def reload():
    """Forget the cache (tests)."""
    _CACHE.clear()
    SKIPPED.clear()
    del LOAD_ERRORS[:]
    for key in [k for k in sys.modules if k.startswith(_PACKAGE + '.')]:
        del sys.modules[key]
