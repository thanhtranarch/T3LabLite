# -*- coding: utf-8 -*-
"""FamiGen Studio message router (pure Python: no Revit, no WPF, no HTTP).

The Studio page (``lib/FamilyGen/studio/``) talks to Python with JSON
messages: the page sends ``{"type": ..., ...}`` through
``window.chrome.webview.postMessage`` and pyRevit 7.1+ hands each one to
``StudioController.handle``; the controller answers through ``post(dict)``,
which the host maps to ``output.post_message``. Every side effect goes through
an adapter the host injects, so the whole protocol is unit-tested
(``lite_guard/test_famigen_studio.py``):

* ``http``   - ``start(url, headers, body_text, callback)`` / ``cancel()``;
               ``callback(status, raw)`` runs later on the UI thread.
* ``revit``  - file pickers, clipboard, ``create_family(schema, folder, load,
               callback)`` (ExternalEvent when modeless), ``close()``.
* ``claude`` - ``api_key()``, ``models()``, ``default_model()``.
* ``prefs``  - ``get(key, default)``, ``set(key, value)``, ``save()``.
* ``store``  - optional ``ProposalStore`` (Claude results are kept there too, so
               the MCP tool ``famigen_create_family`` can build them by id).

The protocol is listed in docs/famigen-studio-spec.md section 6.
"""
import json
import os
import re
import time

from FamilyGen import claude_designer as cd
from FamilyGen import flex
from FamilyGen import param_profiles
from FamilyGen import guidance
from FamilyGen import preview_mesh
from FamilyGen import studio_link
from FamilyGen import templates as family_templates
from Intelligence.family_schema import (
    AUTO_CATEGORY, DETAIL_LEVELS, SUPPORTED_CATEGORIES, build_system_prompt, color_hex,
    expand_schema, material_parameter_name, parse_color, plane_table, schema_summary,
    type_names, validate_family_schema,
)

STUDIO_VERSION = '1.0'
MIN_PYREVIT = (7, 0)
MESSAGING_PYREVIT = (7, 1)
PREVIEW_DECIMALS = 1
MAX_JSON_CHARS = 2000000
THUMB_MAX_BYTES = 2 * 1024 * 1024
MAX_CHECKS = cd.MAX_CHECKS   # render-vs-reference rounds per user message
# The progress strip above the panes, in order. States: todo · active · done ·
# wait (Claude asked a question) · warn (problems left) · error · skip.
PROGRESS_STEPS = (('brief', 'Brief'), ('draft', 'Claude drafts'), ('check', 'FamiGen checks'),
                  ('review', '3D review'), ('compare', 'Reference check'),
                  ('create', 'Create family'))
PROGRESS_IDLE = u'Describe a family or pick a template to start.'

def type_table(schema):
    """Parameters × types grid for the Studio's Parameters tab (pure)."""
    if not isinstance(schema, dict):
        return {'columns': [], 'rows': []}
    names = type_names(schema)
    types = [t for t in schema.get('types') or [] if isinstance(t, dict)]
    values_by_type = {}
    for ftype in types:
        if isinstance(ftype.get('name'), str):
            vals = ftype.get('values') if isinstance(ftype.get('values'), dict) else {}
            values_by_type[ftype['name'].strip()] = dict(
                (k.strip().lower(), v) for k, v in vals.items() if isinstance(k, str))
    rows = []
    for param in schema.get('parameters') or []:
        if isinstance(param, dict) and isinstance(param.get('name'), str):
            rows.append({'name': param['name'], 'kind': param.get('type', 'length'),
                         'instance': bool(param.get('instance', False)),
                         'default': param.get('value')})
    for mat in schema.get('materials') or []:
        if isinstance(mat, dict) and isinstance(mat.get('name'), str):
            rows.append({'name': material_parameter_name(mat), 'kind': 'material',
                         'instance': False, 'default': mat['name']})
    for row in rows:
        cells = []
        for tname in names:
            val = values_by_type.get(tname, {}).get(row['name'].strip().lower())
            cells.append({'value': row['default'] if val is None else val,
                          'override': val is not None})
        row['cells'] = cells
    return {'columns': names, 'rows': rows}


def type_material_map(schema):
    """{type name: {material: material shown}} for the preview's type switch."""
    out = {}
    if not isinstance(schema, dict):
        return out
    params = {}
    for mat in schema.get('materials') or []:
        if isinstance(mat, dict) and isinstance(mat.get('name'), str):
            params[material_parameter_name(mat).lower()] = mat['name']
    for ftype in schema.get('types') or []:
        if not isinstance(ftype, dict) or not isinstance(ftype.get('name'), str):
            continue
        swap = {}
        for key, value in (ftype.get('values') or {}).items() if isinstance(ftype.get('values'), dict) else []:
            if isinstance(key, str) and isinstance(value, str) and key.strip().lower() in params:
                swap[params[key.strip().lower()]] = value.strip()
        out[ftype['name'].strip()] = swap
    return out


_PATH = re.compile(r'^\$\.(geometry|materials|parameters|types|reference_planes|dimensions)\[(\d+)\](\.[A-Za-z_]+)?(?:\S*)?:\s*')
_ROOT = {'$.family_name': 'Family name', '$.family_category': 'Category', '$.geometry': 'Parts',
         '$.materials': 'Materials', '$.parameters': 'Parameters', '$.types': 'Types', '$': 'Family'}
_KIND = {'geometry': 'Part', 'materials': 'Material', 'parameters': 'Parameter', 'types': 'Type',
         'reference_planes': 'Plane', 'dimensions': 'Dimension'}


def friendly_issue(schema, text):
    """'$.geometry[2].material: unknown …' -> "Part 'Leg' (material): unknown …".

    For people, not for Claude: Claude still gets the exact JSON paths.
    """
    text = u'{}'.format(text)
    match = _PATH.match(text)
    if match and isinstance(schema, dict):
        kind, index, field = match.group(1), int(match.group(2)), match.group(3)
        items = schema.get(kind) if isinstance(schema.get(kind), list) else []
        item = items[index] if index < len(items) and isinstance(items[index], dict) else {}
        name = item.get('id') or item.get('name') or item.get('parameter') or u'#{}'.format(index + 1)
        where = u"{} '{}'".format(_KIND[kind], name)
        if field:
            where += u' ({})'.format(field[1:].replace('_', ' '))
        return where + u': ' + text[match.end():]
    for path, label in sorted(_ROOT.items(), key=lambda kv: -len(kv[0])):
        if text.startswith(path + ':'):
            return label + text[len(path):]
    return text


def friendly_model(model_id):
    """'claude-opus-5-5' -> 'Claude Opus 5.5' (ids come from the live list)."""
    text = u'{}'.format(model_id or u'')
    match = re.match(r'^claude-([a-z]+)-(\d+)(?:-(\d{1,2}))?(?:-\d{8})?$', text)
    if not match:
        return text
    version = match.group(2) + (u'.' + match.group(3) if match.group(3) else u'')
    return u'Claude {} {}'.format(match.group(1).capitalize(), version)


def constraint_overlay(schema, expanded, model):
    """Planes and labelled dimensions to draw, plus lock warnings (pure).

    A lock only takes in Revit when the locked face lies on its plane, so a
    part whose bounding box misses the plane is reported here, before Revit.
    """
    table = plane_table(schema)
    custom = set(p['name'].strip().lower() for p in schema.get('reference_planes') or []
                 if isinstance(p, dict) and isinstance(p.get('name'), str))
    planes = [{'name': v[0], 'axis': v[1], 'offset': v[2], 'template': k not in custom}
              for k, v in sorted(table.items(), key=lambda kv: (kv[1][1], kv[1][2]))]
    dims = []
    for dim in schema.get('dimensions') or []:
        if not isinstance(dim, dict):
            continue
        a = table.get(u'{}'.format(dim.get('from') or '').strip().lower())
        b = table.get(u'{}'.format(dim.get('to') or '').strip().lower())
        if a and b and a[1] == b[1]:
            dims.append({'parameter': dim.get('parameter'), 'axis': a[1],
                         'from': a[2], 'to': b[2], 'value': round(abs(b[2] - a[2]), 1)})
    warnings = []
    if model is None:
        return planes, dims, warnings
    boxes = {}
    for mesh in model.meshes:
        if mesh.positions:
            lo = [min(p[k] for p in mesh.positions) for k in range(3)]
            hi = [max(p[k] for p in mesh.positions) for k in range(3)]
            if mesh.index in boxes:
                old_lo, old_hi = boxes[mesh.index]
                lo = [min(lo[k], old_lo[k]) for k in range(3)]
                hi = [max(hi[k], old_hi[k]) for k in range(3)]
            boxes[mesh.index] = (lo, hi)
    for index, geom in enumerate(expanded.get('geometry') or []):
        if not isinstance(geom, dict) or not isinstance(geom.get('locks'), dict) or index not in boxes:
            continue
        lo, hi = boxes[index]
        for key, name in sorted(geom['locks'].items()):
            plane = table.get(u'{}'.format(name).strip().lower())
            if not plane or key[0] not in 'xyz' or plane[1] != key[0]:
                continue
            k = 'xyz'.index(key[0])
            face = lo[k] if key.endswith('_min') else hi[k]
            if abs(face - plane[2]) > 1.0:
                warnings.append(u'{}: locks.{} -> {!r} at {} mm, but that face is at {} mm; '
                                u'Revit will not lock it'.format(
                                    geom.get('id') or '#{}'.format(index + 1), key, plane[0],
                                    round(plane[2], 1), round(face, 1)))
    return planes, dims, warnings


def preview_payload(json_text, type_name=None):
    """Validate + tessellate the editor JSON into what the 3D page draws.

    ``type_name``: the family type shown; its length values flex the preview
    through the labelled dimensions (FamilyGen/flex.py), as Revit would."""
    payload = {'ok': False, 'parse_error': None, 'errors': [], 'warnings': [], 'parts': [],
               'planes': [], 'dimensions': [],
               'materials': [], 'summary': {}, 'size_mm': [0, 0, 0], 'bbox': None,
               'table': {'columns': [], 'rows': []}, 'type_materials': {}, 'triangles': 0}
    text = (json_text or u'').strip()
    if not text:
        return payload
    if len(text) > MAX_JSON_CHARS:
        payload['parse_error'] = u'The JSON is larger than {} characters.'.format(MAX_JSON_CHARS)
        return payload
    try:
        schema = json.loads(text)
    except ValueError as ex:
        payload['parse_error'] = u'{}'.format(ex)
        return payload
    if not isinstance(schema, dict):
        payload['errors'] = [u'$: must be a JSON object with a geometry array']
        return payload
    errors, warnings = validate_family_schema(schema)
    try:
        model = preview_mesh.build_preview(schema)
    except Exception as ex:          # the preview never blocks the editor
        model = None
        warnings = list(warnings) + [u'Preview failed: {}'.format(ex)]
    expanded = expand_schema(schema)
    geometry = expanded.get('geometry') or []
    values = flex.parameter_values(schema, type_name)
    moved = {}
    if model is not None:
        try:
            moved = flex.flex_model(model, expanded, schema, values)
        except Exception as ex:      # a flex bug must not hide the model
            warnings = list(warnings) + [u'Resizing the preview failed: {}'.format(ex)]
            moved = {}
    parts = []
    if model is not None:
        warnings = list(warnings) + list(model.warnings)
        for mesh in model.meshes:
            if not mesh.triangles:
                continue
            geom = geometry[mesh.index] if 0 <= mesh.index < len(geometry) else {}
            geom = geom if isinstance(geom, dict) else {}
            levels = geom.get('detail_levels')
            if not isinstance(levels, list) or not levels:
                levels = list(DETAIL_LEVELS)
            flat = []
            for p in mesh.positions:
                flat.extend(round(float(c), PREVIEW_DECIMALS) for c in p)
            idx = []
            for tri in mesh.triangles:
                idx.extend(tri)
            parts.append({'index': mesh.index, 'label': mesh.label, 'kind': mesh.kind,
                          'material': mesh.material, 'solid': bool(mesh.is_solid),
                          'subcategory': geom.get('subcategory') or u'',
                          'levels': [lv for lv in levels if lv in DETAIL_LEVELS],
                          'pos': flat, 'idx': idx})
        if model.bbox_min is not None:
            payload['bbox'] = [list(model.bbox_min), list(model.bbox_max)]
            payload['size_mm'] = [round(v, 1) for v in model.size_mm()]
        payload['triangles'] = model.triangle_count
    counts = model.material_counts() if model is not None else {}
    materials = []
    for mat in schema.get('materials') or []:
        if isinstance(mat, dict) and isinstance(mat.get('name'), str):
            rgb = parse_color(mat.get('color'))
            materials.append({'name': mat['name'],
                              'color': color_hex(rgb) if rgb else None,
                              'transparency': mat.get('transparency') or 0,
                              'parameter': material_parameter_name(mat),
                              'parts': counts.get(mat['name'], 0)})
    planes, dims, lock_warnings = constraint_overlay(
        flex.flexed_schema_planes(schema, values) if moved else schema, expanded, model)
    warnings = list(warnings) + lock_warnings
    payload.update({
        'issues': {'errors': [friendly_issue(schema, e) for e in errors],
                   'warnings': [friendly_issue(schema, w) for w in warnings]},
        'planes': planes, 'dimensions': dims,
        'ok': not errors, 'errors': errors, 'warnings': warnings, 'parts': parts,
        'materials': materials, 'summary': schema_summary(schema),
        'table': type_table(schema), 'type_materials': type_material_map(schema),
        'preview_type': type_name or u'', 'flexed': sorted(moved),
        'param_fit': param_profiles.evaluate(schema),
    })
    return payload


class StudioController(object):
    def __init__(self, post, http, revit, claude, prefs, store=None,
                 overlay_reader=None, pyrevit_version=None):
        self.post = post
        self.http = http
        self.revit = revit
        self.claude = claude
        self.prefs = prefs
        self.store = store
        self.read_overlay = overlay_reader or guidance.read_overlay
        self.pyrevit_version = pyrevit_version
        self.category = prefs.get('category', AUTO_CATEGORY)
        if self.category not in SUPPORTED_CATEGORIES + (AUTO_CATEGORY,):
            self.category = AUTO_CATEGORY
        self.check_refs = bool(prefs.get('check_refs', True))
        self._ref_camera = None
        self._checks_left = 0
        self._snapshot_token = 0
        self.effort = prefs.get('effort', cd.DEFAULT_EFFORT)
        self.model = prefs.get('model') or None
        self.session = None
        self.attachments = []
        self._att_seq = 0
        self._request_id = 0
        self.busy = False
        self.current_json = u''
        self.engine_pref = prefs.get('engine', 'auto')      # auto | claude | mcp
        self.mcp_client = None
        self._last_beat = 0.0
        self._selected_part = None
        self._proposal_id = None
        self._steps = {}
        self._progress_text = PROGRESS_IDLE
        self.preview_type = None        # family type shown (and flexed) in the 3D view
        self._mcp_waiting = []          # inbox ids sent over MCP and not read yet
        self._mcp_working = False       # the AI app read a message, no family back yet
        self._mcp_seen_call = 0.0
        try:
            self._seen = set(store.ids()) if store is not None else set()
        except Exception:
            self._seen = set()

    # ── dispatch ─────────────────────────────────────────────────────────
    def handle(self, message):
        """Route one page message. Never raises: errors go back to the page."""
        if not isinstance(message, dict):
            return
        kind = message.get('type')
        handler = getattr(self, '_on_' + u'{}'.format(kind), None) if kind else None
        if handler is None:
            self._status(u'Unknown request from the page: {}'.format(kind), 'warning')
            return
        try:
            handler(message)
        except Exception as ex:      # a bug must not leave the page busy forever
            self._set_busy(False)
            self.post({'type': 'error', 'text': u'{}: {}'.format(kind, ex)})

    def _status(self, text, tone='info'):
        self.post({'type': 'status', 'text': text, 'tone': tone})

    def _set_busy(self, on, text=u''):
        self.busy = bool(on)
        self.post({'type': 'busy', 'on': self.busy, 'text': text})

    # ── progress strip ───────────────────────────────────────────────────
    def _progress_payload(self):
        return {'type': 'progress', 'text': self._progress_text,
                'steps': [{'id': sid, 'label': label, 'state': self._steps.get(sid, 'todo')}
                          for sid, label in PROGRESS_STEPS]}

    def _progress(self, step, state='active', text=u'', skip=(), done=()):
        """Move the strip to ``step``; open steps before it are done, later ones reopen.

        ``step`` None resets the strip. ``skip`` / ``done`` set other steps outright
        (the reference check runs after the 3D review, or not at all)."""
        ids = [sid for sid, _ in PROGRESS_STEPS]
        if step is None:
            self._steps = {}
        else:
            at = ids.index(step)
            for i, sid in enumerate(ids):
                cur = self._steps.get(sid, 'todo')
                if i < at and cur in ('todo', 'active', 'wait'):
                    self._steps[sid] = 'done'
                elif i > at and cur in ('active', 'wait', 'error'):
                    self._steps[sid] = 'todo'
            self._steps[step] = state
        for sid in skip:
            self._steps[sid] = 'skip'
        for sid in done:
            self._steps[sid] = 'done'
        self._progress_text = text or (PROGRESS_IDLE if step is None else u'')
        self.post(self._progress_payload())

    def _progress_fail(self, text):
        """Mark the running step failed (Claude error, Stop, Revit refused)."""
        ids = [sid for sid, _ in PROGRESS_STEPS]
        running = [sid for sid in ids if self._steps.get(sid) in ('active', 'wait')]
        if running:
            self._progress(running[-1], 'error', text)

    def _progress_review(self, text, compared=False, problems=False):
        """A family is on screen: everything before the 3D review is done."""
        if problems:
            self._steps['check'] = 'warn'
        self._progress('review', 'active', text,
                       skip=() if compared else ('compare',), done=('compare',) if compared else ())

    # ── lifecycle ────────────────────────────────────────────────────────
    def _claude_state(self):
        try:
            key = self.claude.api_key()
        except Exception:
            key = None
        if not key:
            return {'ready': False, 'models': [], 'model': None,
                    'reason': u'Add a Claude API key in Support > LLMs Setting, then reopen '
                              u'FamiGen Studio, or work from an AI app connected over MCP.'}
        try:
            models = [m for m in (self.claude.models() or []) if u'claude' in u'{}'.format(m)]
        except Exception:
            models = []
        if not models:
            return {'ready': False, 'models': [], 'model': None,
                    'reason': u'Claude did not answer with a model list. Check the API key '
                              u'and the internet connection.'}
        if self.model not in models:
            self.model = self.claude.default_model() or models[0]
            if self.model not in models:
                self.model = models[0]
        return {'ready': True, 'models': models, 'model': self.model, 'reason': u'',
                'labels': [friendly_model(m) for m in models]}

    def _on_ready(self, message):
        self.post({
            'type': 'init',
            'studio_version': STUDIO_VERSION,
            'pyrevit': self.pyrevit_version,
            # False: the host kept its own title bar, so the page hides its own.
            'frameless': bool(getattr(self.revit, 'frameless', True)),
            'categories': [AUTO_CATEGORY] + list(SUPPORTED_CATEGORIES),
            'check_refs': self.check_refs,
            'engine': self._engine_info(),
            'category': self.category,
            'claude': self._claude_state(),
            'efforts': list(cd.EFFORT_LEVELS),
            'effort': self.effort,
            'folder': self.prefs.get('folder', u''),
            'load': bool(self.prefs.get('load', False)),
            'templates': family_templates.templates(),
            'template_groups': family_templates.groups(),
            'into_family': bool(getattr(self.revit, 'into_family', False)),
            'limits': {'attachments': cd.MAX_ATTACHMENTS},
            'progress': self._progress_payload(),
        })

    def _remember(self, key, value):
        self.prefs.set(key, value)
        try:
            self.prefs.save()
        except Exception:
            pass

    def _on_set_category(self, message):
        category = message.get('category')
        if category not in SUPPORTED_CATEGORIES and category != AUTO_CATEGORY:
            self._status(u'Unsupported category: {}'.format(category), 'warning')
            return
        if category != self.category:
            self.category = category
            self._remember('category', category)
            if self.session is not None:
                self._reset_session(u'Category changed to {}: new conversation.'.format(category))

    def _on_set_model(self, message):
        model = message.get('model')
        if model:
            self.model = model
            self._remember('model', model)
            if self.session is not None:
                self.session.model = model       # the API drops unreadable thinking blocks

    def _on_set_effort(self, message):
        effort = message.get('effort')
        if effort in cd.EFFORT_LEVELS:
            self.effort = effort
            self._remember('effort', effort)
            if self.session is not None:
                self.session.effort = effort

    def _reset_session(self, note=None):
        self._request_id += 1
        if self.busy:
            try:
                self.http.cancel()
            except Exception:
                pass
            self._set_busy(False)
        self.session = None
        self._ref_camera = None
        self.attachments = []
        self._post_attachments()
        self.post({'type': 'chat_reset'})
        self._progress(None)
        if note:
            self.post({'type': 'chat', 'role': 'system', 'text': note})

    def _on_new_chat(self, message):
        self._reset_session(u'New conversation. The current model stays.')

    def _on_window(self, message):
        action = message.get('action')
        if action in ('minimize', 'maximize', 'drag'):
            try:
                self.revit.window_state(action)
            except Exception:
                pass

    def _on_close(self, message):
        self._reset_session()
        self.revit.close()

    # ── attachments ──────────────────────────────────────────────────────
    @staticmethod
    def _thumb(att):
        if att['kind'] == 'image' and att['size'] <= THUMB_MAX_BYTES:
            return u'data:{};base64,{}'.format(att['media_type'], att['data'])
        return None

    def _add_attachment(self, record):
        if len(self.attachments) >= cd.MAX_ATTACHMENTS:
            raise ValueError(u'At most {} reference files per message.'.format(cd.MAX_ATTACHMENTS))
        self._att_seq += 1
        record['id'] = 'a{}'.format(self._att_seq)
        self.attachments.append(record)

    def _on_attach_data(self, message):
        try:
            record = cd.attachment_from_base64(message.get('name') or u'pasted image',
                                               message.get('media_type') or u'',
                                               message.get('data') or u'')
            self._add_attachment(record)
        except ValueError as ex:
            self._status(u'{}'.format(ex), 'warning')
        self._post_attachments()

    def _on_set_check(self, message):
        self.check_refs = bool(message.get('on'))
        self._remember('check_refs', self.check_refs)

    def _post_attachments(self):
        items = [{'id': a['id'], 'name': a['name'], 'kind': a['kind'], 'thumb': self._thumb(a),
                  'size': u'{:.1f} MB'.format(a['size'] / 1048576.0) if a['size'] >= 1048576
                  else u'{} KB'.format(max(1, a['size'] // 1024))} for a in self.attachments]
        self.post({'type': 'attachments', 'items': items})

    def _on_attach(self, message):
        paths = self.revit.pick_files() or []
        problems = []
        for path in paths:
            try:
                self._add_attachment(cd.load_attachment(path))
            except ValueError as ex:
                problems.append(u'{}'.format(ex))
        self._post_attachments()
        if problems:
            self._status(u' '.join(problems), 'warning')

    def _on_remove_attachment(self, message):
        self.attachments = [a for a in self.attachments if a['id'] != message.get('id')]
        self._post_attachments()

    # ── Claude ───────────────────────────────────────────────────────────
    def _on_send(self, message):
        if self.busy:
            return
        text = (message.get('text') or u'').strip()
        if not text:
            self._status(u'Describe the family, or the change you want.', 'warning')
            return
        if self.engine() == 'mcp':
            self._send_to_app(text, message)
            return
        state = self._claude_state()
        if not state['ready']:
            self.post({'type': 'chat', 'role': 'system', 'text': state['reason']})
            return
        if self.session is None or self.session.category != self.category:
            auto = self.category == AUTO_CATEGORY
            self.session = cd.ClaudeFamilySession(
                self.category, self.model,
                overlay=u'' if auto else (self.read_overlay(self.category) or u''),
                effort=self.effort, guide_reader=self.read_overlay if auto else None)
        attachments = list(self.attachments)
        current = message.get('json') if message.get('json') is not None else self.current_json
        body = self.session.begin_turn(text, current, attachments)
        self.attachments = []
        self._post_attachments()
        self.post({'type': 'chat', 'role': 'user', 'text': text,
                   'files': [a['name'] for a in attachments],
                   'thumbs': [t for t in (self._thumb(a) for a in attachments) if t]})
        self._checks_left = MAX_CHECKS
        self._set_busy(True, u'Claude is designing the family…')
        self._progress(None)
        self._progress('draft', 'active', u'Claude is designing the family…')
        self._send(body)

    def _send_to_app(self, text, message):
        """MCP mode: the AI app reads the message with famigen_get_studio_state."""
        if message.get('json') is not None:
            self.current_json = message.get('json') or self.current_json
        files = []
        folder = os.path.join(studio_link.folder(), 'inbox_files')
        for att in self.attachments:
            try:
                if not os.path.isdir(folder):
                    os.makedirs(folder)
                path = os.path.join(folder, u'{}_{}'.format(int(time.time()), att['name']))
                if att['kind'] == 'text':
                    with open(path, 'wb') as handle:
                        handle.write(att['data'].encode('utf-8'))
                else:
                    import base64
                    with open(path, 'wb') as handle:
                        handle.write(base64.b64decode(att['data']))
                files.append(path)
            except Exception:
                continue
        thumbs = [t for t in (self._thumb(a) for a in self.attachments) if t]
        self.attachments = []
        self._post_attachments()
        self._publish_state()
        item = studio_link.post_message(text, files)
        app = self.mcp_client or u'your AI app'
        mcp_id = item.get('id') if isinstance(item, dict) else None
        if mcp_id:
            self._mcp_waiting.append(mcp_id)
        self.post({'type': 'chat', 'role': 'user', 'text': text,
                   'files': [os.path.basename(f) for f in files], 'thumbs': thumbs,
                   'mcp_id': mcp_id, 'mcp_state': u'Sent to {} · not read yet'.format(app)})
        if len(self._mcp_waiting) == 1:
            self.post({'type': 'chat', 'role': 'system',
                       'text': u'Sent over MCP. In {}, say "continue in FamiGen": the message '
                               u'shows Read when {} picks it up.'.format(app, app)})
        self._progress(None)
        self._progress('draft', 'wait', u'Sent to {}: waiting for it to read the message…'.format(app))

    def _send(self, body):
        self._request_id += 1
        rid = self._request_id
        session = self.session
        headers = session.headers(self.claude.api_key())
        self.http.start(cd.API_URL, headers, json.dumps(body, ensure_ascii=False),
                        lambda status, raw: self._on_http(rid, session, status, raw))

    def _on_cancel(self, message):
        if not self.busy:
            return
        self._request_id += 1
        try:
            self.http.cancel()
        except Exception:
            pass
        self._set_busy(False)
        # The unanswered user turn stays in the transcript; start clean so the
        # next message is not sent after a user turn with no reply.
        self.session = None
        self._progress_fail(u'Stopped.')
        self.post({'type': 'chat', 'role': 'system',
                   'text': u'Stopped. The next message starts a new conversation.'})

    def _on_http(self, rid, session, status, raw):
        if rid != self._request_id or session is not self.session:
            return                                   # stale (cancelled / reset)
        outcome = session.handle_http(status, raw)
        if outcome.analysis:
            self._show_analysis(outcome.analysis)
        if outcome.kind == 'continue':
            if outcome.errors and outcome.errors[0].startswith(u'Not connected'):
                text = u'Claude is joining loose parts so the model reads as one piece…'
            elif outcome.errors:
                text = u'Claude is fixing {} problem(s) FamiGen found (round {})…'.format(
                    len(outcome.errors), outcome.repairs)
            else:
                text = u'Claude is continuing…'
            self._set_busy(True, text)
            if outcome.errors:
                self._progress('check', 'active', text)
            self._send(outcome.request)
            return
        self._set_busy(False)
        self.post(dict({'type': 'usage'}, **self._usage(session)))
        if outcome.kind == 'error':
            self._progress_fail(u'Claude did not finish: see the chat.')
            self.post({'type': 'chat', 'role': 'error', 'text': outcome.message})
            return
        if outcome.kind == 'reply':
            self._progress('draft', 'wait', u'Claude asked a question: answer in the chat.')
            self.post({'type': 'chat', 'role': 'assistant', 'text': outcome.text})
            return
        # 'schema'
        text = json.dumps(outcome.schema, indent=2, ensure_ascii=False)
        proposal_id = None
        if self.store is not None and outcome.schema is not None:
            try:
                proposal_id = self.store.save(outcome.schema, source='claude',
                                              note=outcome.summary[:200])
            except Exception:
                proposal_id = None
        meta = []
        if outcome.repairs:
            meta.append(u'{} repair round(s)'.format(outcome.repairs))
        if outcome.errors:
            meta.append(u'{} problem(s) left'.format(len(outcome.errors)))
        self.post({'type': 'chat', 'role': 'assistant',
                   'text': outcome.summary or u'Family ready for review.',
                   'assumptions': outcome.assumptions, 'meta': u' · '.join(meta),
                   'proposal_id': proposal_id})
        self.post({'type': 'schema', 'json': text, 'source': 'claude',
                   'proposal_id': proposal_id})
        self.current_json = text
        self._proposal_id = proposal_id
        if proposal_id:
            self._seen.add(proposal_id)
        if session.auto and session.detected_category:
            self.post({'type': 'detected', 'category': session.detected_category})
        self._preview(text)
        compared = self._checks_left < MAX_CHECKS
        if (not outcome.errors and self.check_refs and session.has_references
                and self._checks_left > 0):
            self._checks_left -= 1
            self._snapshot_token += 1
            self._set_busy(True, u'Rendering the model to compare it with your reference…')
            self._progress('compare', 'active',
                           u'Rendering the model to compare it with your reference…')
            self.post({'type': 'snapshot_request', 'token': self._snapshot_token,
                       'views': ['ref', 'iso'] if self._ref_camera else ['iso', 'front'],
                       'camera': self._ref_camera})
            return
        if outcome.errors:
            self._status(u'Claude could not fix every problem; see Checks, then '
                         u'ask Claude again.', 'warning')
            self._progress_review(u'{} problem(s) left: see Checks.'.format(len(outcome.errors)),
                                  compared, problems=True)
        else:
            self._status(u'Review the model, then Create Family.', 'success')
            self._progress_review(u'Review the model, then Create Family.', compared)

    def _on_snapshot(self, message):
        if message.get('token') != self._snapshot_token or self.session is None:
            return
        self._snapshot_token += 1          # one answer per request; a late duplicate is ignored
        images = [u'{}'.format(i).split(u',', 1)[-1] for i in message.get('images') or [] if i]
        if not images:
            self._set_busy(False)
            self._status(u'The 3D view could not be captured; the reference check was skipped.',
                         'warning')
            self._progress_review(u'Reference check skipped: review the model.')
            return
        views = [u'the reference camera' if v == 'ref' else v for v in message.get('views') or []]
        body = self.session.begin_check(images, views=u', '.join(views))
        self.post({'type': 'chat', 'role': 'system',
                   'text': u'Claude is comparing the model with your reference…',
                   'thumbs': [u'data:image/png;base64,' + images[0]]})
        self._set_busy(True, u'Claude is comparing the model with your reference…')
        self._progress('compare', 'active', u'Claude is comparing the model with your reference…')
        self._send(body)

    def _show_analysis(self, analysis):
        """Claude's reading of the reference, as a chat card + the 'Ref' camera."""
        camera = analysis.get('camera') if isinstance(analysis.get('camera'), dict) else None
        if camera:
            try:
                self._ref_camera = {'azimuth_deg': float(camera.get('azimuth_deg', -135)),
                                    'elevation_deg': float(camera.get('elevation_deg', 30))}
                self.post(dict({'type': 'ref_camera'}, **self._ref_camera))
            except (TypeError, ValueError):
                self._ref_camera = None
        size = analysis.get('overall_size_mm') or []
        lines = []
        if isinstance(analysis.get('concept'), str) and analysis['concept'].strip():
            lines.append(u'Concept: {}'.format(analysis['concept'].strip()))
        for label, key in ((u'Levels', 'levels'), (u'Connects', 'connections'), (u'Hollow', 'hollow')):
            items = [u'{}'.format(v) for v in (analysis.get(key) or []) if v][:12]
            if items:
                lines.append(u'{}: {}'.format(label, u'; '.join(items)))
        for comp in (analysis.get('components') or [])[:40]:
            if isinstance(comp, dict):
                bits = [u'{}'.format(comp.get('name') or u'?')]
                if comp.get('count') and comp.get('count') > 1:
                    bits[0] += u' ×{}'.format(comp['count'])
                bits += [u'{}'.format(comp[k]) for k in ('form', 'size_mm', 'color') if comp.get(k)]
                lines.append(u' · '.join(bits))
        self.post({'type': 'chat', 'role': 'analysis',
                   'text': u'What Claude sees: {}{}'.format(
                       analysis.get('view_type') or u'reference',
                       u', about {} mm'.format(u' × '.join(u'{:,.0f}'.format(v) for v in size[:3]))
                       if size else u''),
                   'assumptions': lines,
                   'meta': analysis.get('scale_basis') or u''})

    def _usage(self, session):
        usage = session.usage_summary()
        usd = usage.get('usd')
        return {'model': usage.get('model'), 'requests': usage.get('requests'),
                'input': usage.get('input_tokens'), 'output': usage.get('output_tokens'),
                'cache_read': usage.get('cache_read_input_tokens'),
                'cache_write': usage.get('cache_creation_input_tokens'),
                'usd': None if usd is None else round(usd, 4)}

    # ── JSON editor + preview ────────────────────────────────────────────
    def _preview(self, text):
        payload = preview_payload(text, self.preview_type)
        payload['type'] = 'preview'
        self.post(payload)
        self._publish_state()

    def _on_set_type(self, message):
        """The page shows another family type: flex the preview to its sizes."""
        name = message.get('name') or None
        if name != self.preview_type:
            self.preview_type = name
            if self.current_json:
                self._preview(self.current_json)

    def _on_set_param(self, message):
        """Edit one length / number value in the Types grid and re-flex the preview.

        ``type_name`` names the family type column; empty = the parameter's default
        (a family without types). The schema keeps its geometry at the defaults:
        Revit flexes it to each type's values, and the preview does the same."""
        try:
            schema = json.loads(self.current_json)
        except (TypeError, ValueError):
            self._status(u'Nothing to edit yet: describe a family first.', 'warning')
            return
        name = u'{}'.format(message.get('param') or u'').strip()
        param = next((p for p in schema.get('parameters') or []
                      if isinstance(p, dict) and u'{}'.format(p.get('name') or u'').strip().lower()
                      == name.lower()), None)
        if param is None:
            self._status(u'"{}" is not a family parameter.'.format(name), 'warning')
            return
        kind = param.get('type', 'length')
        try:
            value = float(message.get('value'))
        except (TypeError, ValueError):
            self._status(u'{}: type a number.'.format(name), 'warning')
            return
        if kind == 'integer':
            value = int(round(value))
        if kind == 'length' and value <= 0:
            self._status(u'{}: a length must be more than 0 mm.'.format(name), 'warning')
            return
        type_name = u'{}'.format(message.get('type_name') or u'').strip()
        target = next((t for t in schema.get('types') or [] if isinstance(t, dict)
                       and u'{}'.format(t.get('name') or u'').strip() == type_name), None)
        if target is not None:
            values = target.get('values') if isinstance(target.get('values'), dict) else {}
            for key in [k for k in values if u'{}'.format(k).strip().lower() == name.lower()]:
                del values[key]
            values[param['name']] = value
            target['values'] = values
            self.preview_type = type_name
        else:
            param['value'] = value
        text = json.dumps(schema, indent=2, ensure_ascii=False)
        self.current_json = text
        self.post({'type': 'schema', 'json': text, 'source': 'edit',
                   'proposal_id': self._proposal_id})
        payload = preview_payload(text, self.preview_type)
        payload['type'] = 'preview'
        self.post(payload)
        self._publish_state()
        driven = any(u'{}'.format(d.get('parameter') or u'').strip().lower() == name.lower()
                     for d in schema.get('dimensions') or [] if isinstance(d, dict))
        shown = u'{:,.0f} mm'.format(value) if kind == 'length' else u'{}'.format(value)
        if driven:
            self._status(u'{} = {}: the model resized.'.format(param['name'], shown), 'success')
        else:
            self._status(u'{} = {} saved, but it drives no labelled dimension, so the model does '
                         u'not change. Ask Claude to make it parametric.'.format(param['name'], shown),
                         'warning')

    def _on_make_parametric(self, message):
        """Ask the AI for the category's missing sizes (Sizes tab > Parameter fit)."""
        try:
            schema = json.loads(self.current_json)
        except (TypeError, ValueError):
            self._status(u'Describe a family first.', 'warning')
            return
        text = param_profiles.brief(schema) if isinstance(schema, dict) else u''
        if not text:
            self._status(u'Every standard size of this category already resizes the family.',
                         'success')
            return
        self._on_send({'text': text, 'json': self.current_json})

    def _on_json_changed(self, message):
        self.current_json = message.get('json') or u''
        self._preview(self.current_json)

    def _on_open_json(self, message):
        path = self.revit.pick_json_file()
        if not path:
            return
        try:
            with open(path, 'rb') as handle:
                text = handle.read().decode('utf-8-sig')
            schema = json.loads(text)
        except (IOError, OSError, ValueError, UnicodeDecodeError) as ex:
            self._status(u'{} cannot be read: {}'.format(os.path.basename(path), ex), 'danger')
            return
        if isinstance(schema, dict) and schema.get('family_category') in SUPPORTED_CATEGORIES:
            self._follow_category(schema['family_category'])
        text = json.dumps(schema, indent=2, ensure_ascii=False)
        self.current_json = text
        self.post({'type': 'schema', 'json': text, 'source': 'file'})
        self._preview(text)
        self._progress_review(u'Opened {}: review it, then Create Family.'.format(
            os.path.basename(path)))
        self._status(u'Opened {}.'.format(os.path.basename(path)), 'success')

    def _on_save_json(self, message):
        text = message.get('json') or self.current_json
        name = u'family'
        try:
            name = json.loads(text).get('family_name') or name
        except Exception:
            pass
        path = self.revit.save_json_file(name, text)
        if path:
            self._status(u'Saved {}.'.format(os.path.basename(path)), 'success')

    def _on_copy_prompt(self, message):
        overlay = u'' if self.category == AUTO_CATEGORY else (self.read_overlay(self.category) or u'')
        text = build_system_prompt(self.category, overlay)
        description = (message.get('text') or u'').strip()
        if description:
            text += u'\nFamily description:\n' + description
        self.revit.copy_text(text)
        self._status(u'Prompt for {} copied. Paste it into any AI chat.'.format(self.category),
                     'success')

    # ── AI app over MCP (FamilyGen/studio_link.py) ───────────────────────
    def engine(self):
        """'mcp' when an AI app drives FamiGen (auto) or the user chose it, else 'claude'."""
        if self.engine_pref == 'mcp' or (self.engine_pref == 'auto' and self.mcp_client):
            return 'mcp'
        return 'claude'

    def _engine_info(self):
        return {'pref': self.engine_pref, 'effective': self.engine(),
                'client': self.mcp_client or u''}

    def _on_set_engine(self, message):
        if message.get('engine') in ('auto', 'claude', 'mcp'):
            before = self.engine()
            self.engine_pref = message['engine']
            self._remember('engine', self.engine_pref)
            self._announce_engine(before)

    def _announce_engine(self, before):
        self.post(dict({'type': 'engine'}, **self._engine_info()))
        if self.engine() == before:
            return
        if self.engine() == 'mcp':
            app = self.mcp_client or u'your AI app'
            text = (u'Connected to {} over MCP. Messages you send here go to {}; in {} say '
                    u'"continue in FamiGen". Its families appear here automatically, and the '
                    u'built-in Claude uses no tokens meanwhile.').format(app, app, app)
        else:
            text = u'Back to the built-in Claude: messages go straight to Claude.'
        self.post({'type': 'chat', 'role': 'system', 'text': text})

    def _track_mcp(self):
        """Message receipts and the AI app's last call, for MCP mode."""
        app = self.mcp_client or u'your AI app'
        if self._mcp_waiting:
            try:
                read = studio_link.read_ids(self._mcp_waiting)
            except Exception:
                read = set()
            for mcp_id in [i for i in self._mcp_waiting if i in read]:
                self._mcp_waiting.remove(mcp_id)
                self.post({'type': 'mcp_status', 'id': mcp_id, 'state': 'read',
                           'text': u'Read by {}'.format(app)})
                self._mcp_working = True
                self._progress('draft', 'active', u'{} read your message and is working…'.format(app))
        try:
            call = studio_link.last_activity()
        except Exception:
            call = None
        if call and call['time'] > self._mcp_seen_call:
            first = not self._mcp_seen_call
            self._mcp_seen_call = call['time']
            if self._mcp_working and not first and call['tool']:
                self._progress('draft', 'active', u'{} called {}…'.format(call['client'], call['tool']))

    def tick(self, now=None):
        """Called about once a second by the host: heartbeat, AI app, new proposals."""
        now = time.time() if now is None else now
        self._track_mcp()
        if now - self._last_beat >= 2.0:
            self._last_beat = now
            try:
                studio_link.beat({'category': self.category})
            except Exception:
                pass
        client = studio_link.active_client()
        if client != self.mcp_client:
            before = self.engine()
            self.mcp_client = client
            self._announce_engine(before)
        if self.store is None or self.busy:
            return
        try:
            ids = self.store.ids()
        except Exception:
            return
        fresh = [i for i in ids if i not in self._seen]
        for proposal_id in fresh:
            self._seen.add(proposal_id)
        for proposal_id in fresh[-1:]:                  # only the newest is shown
            record = self.store.load(proposal_id) or {}
            if record.get('source') != 'mcp' or not isinstance(record.get('schema'), dict):
                continue
            schema = record['schema']
            if schema.get('family_category') in SUPPORTED_CATEGORIES:
                self._follow_category(schema['family_category'])
            text = json.dumps(schema, indent=2, ensure_ascii=False)
            self.current_json = text
            self._proposal_id = proposal_id
            self.post({'type': 'schema', 'json': text, 'source': 'mcp', 'proposal_id': proposal_id})
            self._preview(text)
            self._mcp_working = False
            self.post({'type': 'mcp_status', 'id': None, 'state': 'answered',
                       'text': u'Answered by {}'.format(self.mcp_client or u'the AI app')})
            self._progress_review(u'New family from {}: review it, then Create Family.'.format(
                self.mcp_client or u'the AI app'))
            self.post({'type': 'chat', 'role': 'assistant',
                       'text': record.get('note') or u'Here is the family.',
                       'meta': u'from {} over MCP'.format(self.mcp_client or u'an AI app')})
            self._status(u'New family from {}. Review it, then Create Family.'.format(
                self.mcp_client or u'the AI app'), 'success')

    def _publish_state(self):
        try:
            family = json.loads(self.current_json) if self.current_json else None
        except ValueError:
            family = None
        try:
            studio_link.write_state({'category': self.category, 'family': family,
                                     'proposal_id': self._proposal_id,
                                     'selected_part': self._selected_part})
        except Exception:
            pass

    def _on_select(self, message):
        self._selected_part = message.get('label') or None
        self._publish_state()

    def close_link(self):
        try:
            studio_link.leave()
        except Exception:
            pass

    def _follow_category(self, category):
        """A loaded JSON brings its category: Auto just shows it, a fixed pick follows."""
        if self.category == AUTO_CATEGORY:
            self.post({'type': 'detected', 'category': category})
        elif category != self.category:
            self.category = category
            self._remember('category', category)
            self.session = None
            self.post({'type': 'category', 'category': category})

    # ── My families (history) ────────────────────────────────────────────
    def _on_history(self, message):
        items = []
        if self.store is not None:
            try:
                items = self.store.recent(int(message.get('limit') or 30))
            except Exception as ex:
                self._status(u'History cannot be read: {}'.format(ex), 'warning')
        self.post({'type': 'history', 'items': items})

    def _on_load_history(self, message):
        record = self.store.load(message.get('id')) if self.store is not None else None
        schema = record.get('schema') if isinstance(record, dict) else None
        if not isinstance(schema, dict):
            self._status(u'That family is no longer in the history.', 'warning')
            return
        if schema.get('family_category') in SUPPORTED_CATEGORIES:
            self._follow_category(schema['family_category'])
        text = json.dumps(schema, indent=2, ensure_ascii=False)
        self.current_json = text
        self.post({'type': 'schema', 'json': text, 'source': 'history',
                   'proposal_id': message.get('id')})
        self._preview(text)
        self._progress_review(u'Opened from My families: review it, then Create Family.')
        self.post({'type': 'chat', 'role': 'system',
                   'text': u'Opened "{}" from My families. Ask Claude for changes, or Create '
                           u'Family.'.format(schema.get('family_name') or message.get('id'))})

    # ── Revit ────────────────────────────────────────────────────────────
    def _on_pick_folder(self, message):
        folder = self.revit.pick_folder(self.prefs.get('folder', u''))
        if folder:
            self._remember('folder', folder)
            self.post({'type': 'folder', 'path': folder})

    def _on_set_load(self, message):
        self._remember('load', bool(message.get('load')))

    def _on_create(self, message):
        if self.busy:
            return
        text = message.get('json') or self.current_json
        try:
            schema = json.loads(text)
        except (TypeError, ValueError) as ex:
            self._status(u'The family data cannot be read: {}'.format(ex), 'danger')
            return
        errors, _ = validate_family_schema(schema) if isinstance(schema, dict) else (
            [u'$: must be a JSON object'], [])
        if errors:
            self.post({'type': 'chat', 'role': 'error',
                       'text': u'Nothing was created: the model has {} problem(s). See Checks, '
                               u'or ask Claude to fix them.'.format(len(errors))})
            self._status(u'{} problem(s) must be fixed before Create Family.'.format(len(errors)),
                         'danger')
            return
        into_family = bool(getattr(self.revit, 'into_family', False))
        folder = self.prefs.get('folder', u'')
        if not into_family and not folder:
            folder = self.revit.pick_folder(u'')
            if not folder:
                self._status(u'Choose the output folder, then Create Family.', 'warning')
                return
            self._remember('folder', folder)
            self.post({'type': 'folder', 'path': folder})
        load = bool(message.get('load', self.prefs.get('load', False)))
        self._set_busy(True, u'Revit is building "{}"…'.format(schema.get('family_name')))
        self._progress('create', 'active', u'Revit is building "{}"…'.format(schema.get('family_name')),
                       skip=() if self._steps.get('compare') == 'done' else ('compare',))

        def _done(report, error):
            self._set_busy(False)
            if error:
                self._progress('create', 'error', u'"{}" was not created: see the chat.'.format(
                    schema.get('family_name')))
                self.post({'type': 'chat', 'role': 'error', 'text': error})
                self._status(u'"{}" was not created.'.format(schema.get('family_name')), 'danger')
                return
            from FamilyGen.builder_report import report_lines
            lines = report_lines(report)
            self.post({'type': 'created', 'path': report.get('saved_path') or u'',
                       'lines': lines,
                       'problems': bool(report.get('skipped') or report.get('warnings'))})
            self._status(lines[0] + (u' ' + lines[1] if len(lines) > 1 else u''), 'success')
            self._progress('create', 'warn' if report.get('skipped') or report.get('warnings')
                           else 'done', lines[0])

        self.revit.create_family(schema, folder, load, _done)
