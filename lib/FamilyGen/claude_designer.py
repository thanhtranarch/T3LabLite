# -*- coding: utf-8 -*-
"""Direct Claude session for FamiGen Studio (pure Python, no Revit, no HTTP).

The session owns the conversation with Claude and nothing else: it builds the
Messages API request body, reads the response, validates the family Claude
submits, and asks Claude to repair it when the validator finds problems. The
caller sends the requests (``FamilyGen.net_http`` in Revit, a fake in tests),
so the same code runs from a DispatcherTimer, a worker thread or a unit test.

Flow of one user turn::

    body = session.begin_turn(text, current_json, attachments)
    while True:
        status, raw = <POST body to /v1/messages>
        outcome = session.handle_http(status, raw)
        if outcome.kind != 'continue':
            break                       # 'schema' | 'reply' | 'error'
        body = outcome.request          # a repair round or a retry

Contract with the API (checked 2026-11):

* ``submit_family`` is a client tool, ``tool_choice`` stays ``auto``: forced
  tool choice returns a 400 on Claude Opus 5.5, Sonnet 5.5 and Fable 5.1/5.2.
  The system prompt tells Claude to deliver through the tool; a reply without
  a tool call is shown to the user as a question or an answer.
* The transcript is append-only. Assistant content is stored exactly as
  returned (thinking blocks included), and every ``tool_use`` gets its
  ``tool_result`` at the start of the next user message - preserved thinking
  rejects edited history.
* Optional request features - ``output_config.effort``, server-side
  ``fallbacks`` (beta header) and top-level ``cache_control`` - are dropped
  for the rest of the session when the API answers 400 naming them, so an
  older model or account still works.
* Model ids are never hardcoded: the caller passes the id it resolved from
  the live ``/v1/models`` list (``ClaudeProvider``).
"""
import base64
import json
import os

from Intelligence.family_schema import (
    AUTO_CATEGORY, SUPPORTED_CATEGORIES, build_system_prompt, meaning_errors, schema_summary,
    validate_ai_schema, validate_family_schema,
)

API_URL = 'https://api.anthropic.com/v1/messages'
API_VERSION = '2023-06-01'
FALLBACK_BETA = 'server-side-fallback-2026-07-01'

DEFAULT_MAX_TOKENS = 32000
DEFAULT_EFFORT = 'high'
EFFORT_LEVELS = ('low', 'medium', 'high', 'xhigh', 'max')
MAX_REPAIRS = 2
MAX_POLISH = 2          # extra rounds when a valid family does not yet work in space

# Attachments Claude can read: images (vision) and PDFs (document blocks).
IMAGE_TYPES = {'.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
               '.gif': 'image/gif', '.webp': 'image/webp'}
PDF_TYPE = 'application/pdf'
TEXT_TYPES = {'.txt': 'text/plain', '.md': 'text/plain', '.csv': 'text/plain',
              '.json': 'text/plain', '.xml': 'text/plain'}
MAX_IMAGE_BYTES = 5 * 1024 * 1024
MAX_PDF_BYTES = 32 * 1024 * 1024
MAX_TEXT_BYTES = 400 * 1024
MAX_ATTACHMENTS = 8
MEDIA_KIND = dict([(m, 'image') for m in IMAGE_TYPES.values()] + [(PDF_TYPE, 'document')])

SUBMIT_TOOL = {
    'name': 'submit_family',
    'description': (
        'Deliver the complete FamiGen family JSON to the user. FamiGen validates it, '
        'shows it in a 3D review pane and lets the user create the Revit family. '
        'Call it once per answer with the WHOLE family (not a diff). If FamiGen finds '
        'problems it returns them as an error result: fix every listed path and call '
        'submit_family again with the complete corrected family.'),
    'input_schema': {
        'type': 'object',
        'properties': {
            'family': {
                'type': 'object',
                'description': ('The family JSON (schema_version 2): family_name, '
                                'family_category, materials, parameters, types and geometry, '
                                'exactly as the system prompt specifies. Millimeters, radians.'),
            },
            'summary': {
                'type': 'string',
                'description': 'One or two sentences for the user: what you built or changed.',
            },
            'assumptions': {
                'type': 'array',
                'items': {'type': 'string'},
                'description': ('Dimensions, materials or details you chose because the '
                                'request did not state them, one short line each.'),
            },
        },
        'required': ['family', 'summary'],
    },
}

GUIDE_TOOL = {
    'name': 'get_category_spec',
    'description': (
        'Auto-detect only: returns FamiGen\'s design spec for one family category (anatomy, '
        'typical sizes, origin and axes, forms, parameters, detail levels, pitfalls). Call it '
        'once with the category you chose BEFORE writing the family, then follow it.'),
    'input_schema': {
        'type': 'object',
        'properties': {'category': {'type': 'string', 'enum': list(SUPPORTED_CATEGORIES)}},
        'required': ['category'],
    },
}
MAX_GUIDE_CALLS = 3

ANALYSIS_TOOL = {
    'name': 'record_analysis',
    'description': (
        'With reference images or drawings attached: record what you see BEFORE modelling - '
        'the camera of the picture, the overall size, what the object is and how it is used, '
        'its levels, how the parts connect, which parts are hollow, and every component with '
        'its form, size, position and colour. FamiGen shows it to the user and renders your model '
        'from the same camera to compare. Call it once per reference, then submit_family.'),
    'input_schema': {
        'type': 'object',
        'properties': {
            'camera': {
                'type': 'object',
                'description': ('Viewpoint of the main picture in FamiGen axes (X right, Y back, Z up; '
                                'the front of the object faces -Y). azimuth_deg: direction from the '
                                'object to the camera, from +X counter-clockwise (front = -90, '
                                'front-left = -135, front-right = -45, right = 0). elevation_deg: '
                                'above the horizon (0 = eye level, 90 = plan).'),
                'properties': {'azimuth_deg': {'type': 'number'}, 'elevation_deg': {'type': 'number'}},
                'required': ['azimuth_deg', 'elevation_deg'],
            },
            'view_type': {'type': 'string',
                          'description': 'photo, render, sketch, elevation, plan, section, cut sheet'},
            'overall_size_mm': {'type': 'array', 'items': {'type': 'number'},
                                'description': '[width X, depth Y, height Z] and how you derived it in notes'},
            'scale_basis': {'type': 'string',
                            'description': 'What sets the scale: a dimension in the reference, a standard '
                                           'size (door 2100, seat 450, rail 1100 ...), or an assumption.'},
            'components': {
                'type': 'array',
                'description': 'Every visible part group, largest first.',
                'items': {'type': 'object', 'properties': {
                    'name': {'type': 'string'},
                    'form': {'type': 'string', 'description': 'Extrusion, Blend, Revolution, Sweep, '
                                                              'SweptBlend, Cylinder, Tube or Void'},
                    'size_mm': {'type': 'string'}, 'position': {'type': 'string'},
                    'color': {'type': 'string'}, 'count': {'type': 'integer'}},
                    'required': ['name', 'form']},
            },
            'concept': {'type': 'string',
                        'description': 'What the object is, who or what uses it and the route taken '
                                       'through it, where it stands or hangs.'},
            'levels': {'type': 'array', 'items': {'type': 'string'},
                       'description': 'Every level something stands on, with its height: '
                                      '"ground 0", "seat 450", "deck 2350".'},
            'connections': {'type': 'array', 'items': {'type': 'string'},
                            'description': 'How the parts connect, one per line: "stair: lower deck '
                                           'north edge -> porch front edge", "spout -> basin". These '
                                           'become joints in the family.'},
            'hollow': {'type': 'array', 'items': {'type': 'string'},
                       'description': 'Parts that are hollow in reality (pipes, tubes, basins, bowls, '
                                      'housings) and their wall thickness.'},
            'notes': {'type': 'string'},
        },
        'required': ['camera', 'overall_size_mm', 'components'],
    },
}
MAX_CHECKS = 2

_WORKFLOW = u'''
WORKING WITH THE USER (FamiGen Studio, direct Claude session):
- The user is a BIM modeller in Autodesk Revit. They review your family in a 3D
  preview with its materials, then create the .rfa themselves.
- Every family you produce goes through submit_family with the complete JSON.
  Do not paste JSON in your text and do not describe JSON you did not submit.
- Missing dimensions or finishes: choose sensible real-world values for the
  category, build the family, and list each choice in assumptions. Ask a short
  question instead of building only when the request cannot be modelled at all.
- Refinements: change only what the user asks; keep part ids, names, materials,
  parameters and types that they did not mention. When the user message
  includes "Current family JSON", that JSON is the source of truth - the user
  may have edited it by hand.
- Reference images, sketches and PDF spec sheets: take proportions and
  dimensions from them. Dimension text in the reference wins over proportions
  measured from the picture. If a reference gives no scale, say which overall
  size you assumed.
- Model the visible form with the fewest parts that read correctly at 1:20:
  merge parts that share a material and would never be scheduled separately.
- Keep summary to one or two sentences. Plain language, no JSON terms.

THINK IN THIS ORDER (every family, with or without a reference):
1. Concept: what is it, who or what uses it and by which route, where does it stand
   or hang. Write it in "concept".
2. Levels and footprints: the heights things stand on, and the plan area of each main
   element. Unrelated elements do not share a footprint.
3. Parts and roles: what each part does (support, surface, enclosure, container,
   conduit, access, guard, fixing, control, decor) - and from that, solid or hollow.
4. Joints: where each part meets its neighbour, as one shared coordinate. Write the
   important ones in "joints"; FamiGen checks them.
5. Only then the geometry, derived from those levels and joint points.
If submit_family comes back with Not connected / Joint / Hollow / Access / Crossing
findings, the geometry disagrees with your own concept: correct the geometry.

REFERENCE ANALYSIS (when images, drawings or cut sheets are attached):
- First call record_analysis: the camera of the picture, the scale basis, the overall
  size, the concept, the levels, how the parts connect, which parts are hollow, and
  EVERY component (form, size, position, colour, count). Count repeated
  members (pickets, slats, legs) instead of guessing. Then build exactly that list.
- Measure in proportions: pick the most reliable known size (a door, a seat height, a
  stated dimension), convert every other length by its ratio to it in the picture,
  and correct for perspective (parts farther from the camera look smaller).
- Positions follow the picture's layout: what is left / right / in front / behind in
  the picture keeps that place in plan once the camera azimuth is accounted for.
FIDELITY TO REFERENCES:
- With a photo, render, drawing or cut sheet attached, reproduce THAT product:
  same number and arrangement of parts, same proportions, same profile shapes
  and the same colours. First identify each view (front, side, plan, iso), then
  take dimensions from text in the reference, and the rest from ratios between
  visible edges against one known overall size.
- Pick the form that matches the shape: Revolution for turned/round parts,
  Sweep for bent tubes and handles, Blend for tapers, Extrusion with
  inner_loops or a void for openings, Cylinder for rods and legs.

COMPLEX FAMILIES:
- Use as many parts as the object needs; model repeated members once with
  repeat. Put small details at medium/fine detail levels.
- For box-like products give the overall Width, Depth and Height reference
  planes, labelled dimensions and face locks so the family resizes per type.
'''


class Outcome(object):
    """What one API response means for the UI.

    kind: 'schema'   - a family was submitted (``schema``; ``errors`` is empty
                       when it passed validation, else the last invalid draft)
          'reply'    - Claude answered in text only (a question or an answer)
          'continue' - send ``request`` next (repair round or feature retry)
          'error'    - the turn failed; ``message`` says why in plain words
    """

    def __init__(self, kind, **fields):
        self.kind = kind
        self.schema = fields.get('schema')
        self.summary = fields.get('summary') or u''
        self.assumptions = list(fields.get('assumptions') or [])
        self.text = fields.get('text') or u''
        self.errors = list(fields.get('errors') or [])
        self.warnings = list(fields.get('warnings') or [])
        self.request = fields.get('request')
        self.analysis = fields.get('analysis')
        self.message = fields.get('message') or u''
        self.repairs = fields.get('repairs', 0)

    def to_dict(self):
        return {'kind': self.kind, 'schema': self.schema, 'summary': self.summary,
                'assumptions': self.assumptions, 'text': self.text, 'errors': self.errors,
                'warnings': self.warnings, 'message': self.message, 'repairs': self.repairs}


def continuity_problems(family):
    """Spatial findings for a valid family: loose parts, missed joints, solid lumps
    that should be hollow, access parts without landings, blocked routes ([] when
    the family reads as one working object). See FamilyGen.spatial_check."""
    try:
        from FamilyGen.preview_mesh import build_preview
        model = build_preview(family)
    except Exception:
        return []
    from FamilyGen.spatial_check import PREFIXES
    return [w for w in model.warnings if w.startswith(PREFIXES)]


def load_attachment(path):
    """Read a reference file into an attachment record, or raise ValueError."""
    if not path or not os.path.isfile(path):
        raise ValueError(u'File not found: {}'.format(path))
    ext = os.path.splitext(path)[1].lower()
    size = os.path.getsize(path)
    name = os.path.basename(path)
    if ext in TEXT_TYPES:
        if size > MAX_TEXT_BYTES:
            raise ValueError(u'{} is {} KB; text references are limited to {} KB.'.format(
                name, size // 1024, MAX_TEXT_BYTES // 1024))
        with open(path, 'rb') as handle:
            text = handle.read().decode('utf-8-sig', 'replace')
        return {'name': name, 'kind': 'text', 'media_type': 'text/plain', 'size': size,
                'data': text}
    if ext in IMAGE_TYPES:
        media = IMAGE_TYPES[ext]
    elif ext == '.pdf':
        media = PDF_TYPE
    else:
        raise ValueError(u'{}: use an image (PNG, JPG, GIF, WEBP), a PDF or a text file '
                         u'(TXT, MD, CSV, JSON).'.format(name))
    with open(path, 'rb') as handle:
        raw = handle.read()
    return attachment_from_bytes(name, media, raw)


def _limit(kind):
    return MAX_IMAGE_BYTES if kind == 'image' else MAX_PDF_BYTES


def attachment_from_bytes(name, media_type, raw):
    """Attachment record from file bytes (picked, pasted or dropped)."""
    kind = MEDIA_KIND.get(media_type)
    if kind is None:
        raise ValueError(u'{}: {} cannot be read by Claude; use an image or a PDF.'.format(
            name, media_type))
    if len(raw) > _limit(kind):
        raise ValueError(u'{} is {:.1f} MB; the limit for this file type is {} MB.'.format(
            name, len(raw) / 1048576.0, _limit(kind) // 1048576))
    return {'name': name, 'kind': kind, 'media_type': media_type, 'size': len(raw),
            'data': base64.b64encode(raw).decode('ascii')}


def attachment_from_base64(name, media_type, data):
    """Attachment record from base64 sent by the page (clipboard paste, drag-drop)."""
    if media_type and media_type.startswith('text/'):
        try:
            text = base64.b64decode(data or '').decode('utf-8', 'replace')
        except Exception:
            raise ValueError(u'{} cannot be read.'.format(name))
        if len(text.encode('utf-8')) > MAX_TEXT_BYTES:
            raise ValueError(u'{} is larger than {} KB.'.format(name, MAX_TEXT_BYTES // 1024))
        return {'name': name, 'kind': 'text', 'media_type': 'text/plain',
                'size': len(text.encode('utf-8')), 'data': text}
    try:
        raw = base64.b64decode(data or '', validate=True)
    except Exception:
        raise ValueError(u'{} cannot be read.'.format(name))
    return attachment_from_bytes(name, media_type, raw)


def _attachment_block(att):
    if att['kind'] == 'text':
        return {'type': 'document', 'title': att.get('name') or 'Reference',
                'source': {'type': 'text', 'media_type': 'text/plain', 'data': att['data']}}
    source = {'type': 'base64', 'media_type': att['media_type'], 'data': att['data']}
    if att['kind'] == 'image':
        return {'type': 'image', 'source': source}
    return {'type': 'document', 'source': source, 'title': att.get('name') or 'Reference'}


def _parse_error(raw):
    """(error type, message) from an API error body, tolerant of anything."""
    try:
        body = json.loads(raw) if isinstance(raw, str) else (raw or {})
        err = body.get('error') or {}
        return err.get('type') or u'', err.get('message') or u'{}'.format(raw)[:400]
    except Exception:
        return u'', u'{}'.format(raw or u'')[:400]


class ClaudeFamilySession(object):
    """One FamiGen Studio conversation with Claude for one family category."""

    def __init__(self, category, model, overlay=u'', effort=DEFAULT_EFFORT,
                 max_tokens=DEFAULT_MAX_TOKENS, max_repairs=MAX_REPAIRS, guide_reader=None):
        if category not in SUPPORTED_CATEGORIES and category != AUTO_CATEGORY:
            raise ValueError(u'Choose a supported family category first.')
        if not model:
            raise ValueError(u'No Claude model is available. Check the Claude API key in '
                             u'Support > LLMs Setting.')
        self.category = category
        self.model = model
        self.effort = effort if effort in EFFORT_LEVELS else DEFAULT_EFFORT
        self.max_tokens = int(max_tokens)
        self.max_repairs = int(max_repairs)
        self.auto = category == AUTO_CATEGORY
        self.guide_reader = guide_reader if self.auto else None
        self._guide_calls = 0
        self.system_prompt = build_system_prompt(category, overlay, delivery='tool') + _WORKFLOW
        if self.guide_reader is not None:
            self.system_prompt += (u'\nAUTO-DETECT: decide the category first, call '
                                   u'get_category_spec with it, then build the family.\n')
        self.has_references = False    # an image / PDF / text file was attached
        self.analysis = None           # last record_analysis input
        self.detected_category = None  # Auto mode: the category Claude chose
        self.messages = []           # append-only transcript
        self._pending_results = []   # tool_result blocks owed to the next user message
        self._repairs = 0
        self._polish_left = MAX_POLISH
        self._dropped = set()        # optional features the API refused
        self._feature_retries = 0
        self.last_schema = None      # last family that passed validation
        self.turns = 0
        self.usage = {'input_tokens': 0, 'output_tokens': 0,
                      'cache_read_input_tokens': 0, 'cache_creation_input_tokens': 0,
                      'requests': 0}

    # ── request ──────────────────────────────────────────────────────────
    def headers(self, api_key):
        h = {'x-api-key': api_key, 'anthropic-version': API_VERSION,
             'content-type': 'application/json'}
        if 'fallbacks' not in self._dropped:
            h['anthropic-beta'] = FALLBACK_BETA
        return h

    def build_request(self):
        body = {
            'model': self.model,
            'max_tokens': self.max_tokens,
            'system': [{'type': 'text', 'text': self.system_prompt}],
            'tools': [SUBMIT_TOOL] + ([GUIDE_TOOL] if self.guide_reader is not None else [])
                     + ([ANALYSIS_TOOL] if self.has_references else []),
            'tool_choice': {'type': 'auto'},
            'messages': self.messages,
        }
        if 'output_config' not in self._dropped:
            body['output_config'] = {'effort': self.effort}
        if 'fallbacks' not in self._dropped:
            body['fallbacks'] = 'default'
        if 'cache_control' not in self._dropped:
            body['cache_control'] = {'type': 'ephemeral'}
        return body

    def begin_turn(self, text, current_json=None, attachments=()):
        """Append the user's message and return the request body to send."""
        text = (text or u'').strip()
        if not text:
            raise ValueError(u'Describe the family or the change you want first.')
        attachments = list(attachments or [])
        if len(attachments) > MAX_ATTACHMENTS:
            raise ValueError(u'Attach at most {} reference files.'.format(MAX_ATTACHMENTS))
        content = list(self._pending_results)
        self._pending_results = []
        content.extend(_attachment_block(a) for a in attachments)
        if attachments:
            self.has_references = True
        if attachments:
            names = u', '.join(a.get('name') or u'reference' for a in attachments)
            text = u'Reference files attached: {}.\n\n{}'.format(names, text)
        context = self._json_context(current_json)
        if context:
            text = context + u'\n\n' + text
        content.append({'type': 'text', 'text': text})
        self.messages.append({'role': 'user', 'content': content})
        self._repairs = 0
        self._polish_left = MAX_POLISH
        self._feature_retries = 0
        self.turns += 1
        return self.build_request()

    def begin_check(self, renders, views=u'the reference camera and isometric'):
        """A verification turn: renders of the last family vs. the references.

        ``renders`` are base64 PNG strings captured from the 3D view. Returns
        the request body. Claude either submits a corrected family or answers
        in text starting with "Matches".
        """
        renders = [r for r in renders or [] if r]
        if not renders:
            raise ValueError(u'No render to compare.')
        content = list(self._pending_results)
        self._pending_results = []
        for data in renders[:4]:
            content.append({'type': 'image', 'source': {'type': 'base64',
                                                        'media_type': 'image/png', 'data': data}})
        content.append({'type': 'text', 'text': (
            u'FamiGen check: these are renders ({}) of the family you just submitted, in '
            u'its material colours. Compare them with the reference files earlier in this '
            u'conversation: number and arrangement of parts, proportions, profile shapes, '
            u'colours. If anything differs noticeably, call submit_family with the complete '
            u'corrected family and say in summary what you fixed. If it already matches, '
            u'reply in one short sentence that starts with "Matches" and do not call the '
            u'tool.').format(views)})
        self.messages.append({'role': 'user', 'content': content})
        self._repairs = 0
        self._polish_left = MAX_POLISH
        self._feature_retries = 0
        return self.build_request()

    def _json_context(self, current_json):
        """The editor JSON when it differs from what Claude last submitted."""
        if not current_json:
            return u''
        try:
            current = json.loads(current_json) if isinstance(current_json, str) else current_json
        except ValueError:
            return (u'Current family JSON (cannot be parsed; the user is editing it):\n'
                    + u'{}'.format(current_json)[:20000])
        if not isinstance(current, dict) or current == self.last_schema:
            return u''
        return (u'Current family JSON (source of truth, may include the user\'s edits):\n'
                + json.dumps(current, ensure_ascii=False, separators=(',', ':')))

    # ── response ─────────────────────────────────────────────────────────
    def handle_http(self, status, raw):
        """Turn one HTTP result into an Outcome. ``status`` 0 = transport error."""
        if status == 200:
            try:
                data = json.loads(raw) if isinstance(raw, str) else raw
            except ValueError:
                return Outcome('error', message=u'Claude returned an unreadable response.')
            return self.handle_response(data)
        etype, message = _parse_error(raw)
        if status == 400 and self._feature_retries < 3:
            dropped = self._drop_refused_feature(message)
            if dropped:
                self._feature_retries += 1
                return Outcome('continue', request=self.build_request())
        return Outcome('error', message=self._friendly_error(status, etype, message))

    def _drop_refused_feature(self, message):
        text = (message or u'').lower()
        for feature, words in (('fallbacks', ('fallback', 'server-side-fallback', 'beta')),
                               ('output_config', ('output_config', 'effort')),
                               ('cache_control', ('cache_control',))):
            if feature in self._dropped:
                continue
            if any(w in text for w in words):
                self._dropped.add(feature)
                return feature
        return None

    @staticmethod
    def _friendly_error(status, etype, message):
        if status == 0:
            return u'Claude could not be reached: {}. Check the internet connection or proxy.'.format(message)
        if status == 401 or etype == 'authentication_error':
            return u'Claude rejected the API key. Update it in Support > LLMs Setting.'
        if status == 403 or etype == 'permission_error':
            return u'This Claude API key may not use that model: {}'.format(message)
        if status == 404 or etype == 'not_found_error':
            return u'Claude does not offer that model to this key: {}'.format(message)
        if status == 429 or etype == 'rate_limit_error':
            return u'Claude rate limit reached. Wait a minute and try again.'
        if status in (500, 502, 503, 504, 529) or etype in ('overloaded_error', 'api_error'):
            return u'Claude is busy right now ({}). Try again in a moment.'.format(status)
        return u'Claude returned an error ({}): {}'.format(status, message)

    def _add_usage(self, usage):
        self.usage['requests'] += 1
        for key in ('input_tokens', 'output_tokens', 'cache_read_input_tokens',
                    'cache_creation_input_tokens'):
            try:
                self.usage[key] += int((usage or {}).get(key) or 0)
            except (TypeError, ValueError):
                pass

    def handle_response(self, data):
        if not isinstance(data, dict):
            return Outcome('error', message=u'Claude returned an unreadable response.')
        self._add_usage(data.get('usage'))
        content = data.get('content') or []
        # Append-only: keep the assistant turn exactly as returned.
        self.messages.append({'role': 'assistant',
                              'content': content or [{'type': 'text', 'text': u''}]})
        stop = data.get('stop_reason')
        texts = [b.get('text') or u'' for b in content
                 if isinstance(b, dict) and b.get('type') == 'text']
        calls = [b for b in content if isinstance(b, dict) and b.get('type') == 'tool_use']
        text = u'\n'.join(t for t in texts if t).strip()

        if stop == 'refusal':
            for call in calls:
                self._owe(call, u'Not processed.', True)
            details = data.get('stop_details') or {}
            return Outcome('error', message=(
                u'Claude declined this request{}. Rephrase it as a building product '
                u'description and try again.'.format(
                    u' ({})'.format(details.get('category')) if details.get('category') else u'')))
        if stop == 'max_tokens':
            for call in calls:
                self._owe(call, u'The family was cut off at the output limit.', True)
            if self._repairs < self.max_repairs:
                self._repairs += 1
                self._append_user_text(
                    u'Your answer hit the output limit before the family was complete. Call '
                    u'submit_family again with a more compact family: fewer parts, shorter '
                    u'profiles, no repeated geometry.')
                return Outcome('continue', request=self.build_request(), repairs=self._repairs)
            return Outcome('error', message=(
                u'The family was too large for one answer. Ask for a simpler family '
                u'or split it into separate families.'))

        submit = None
        guided = False
        analysis = None
        for call in calls:
            if call.get('name') == 'submit_family' and submit is None:
                submit = call
            elif call.get('name') == 'record_analysis':
                analysis = call.get('input') or {}
                self.analysis = analysis
                guided = True
                self._owe(call, u'Recorded and shown to the user. Now build exactly these '
                                u'components with submit_family.', False)
            elif call.get('name') == 'get_category_spec' and self.guide_reader is not None:
                guided = True
                self._guide_calls += 1
                category = (call.get('input') or {}).get('category')
                if category not in SUPPORTED_CATEGORIES:
                    self._owe(call, u'Unknown category. Use one of: ' +
                              u', '.join(SUPPORTED_CATEGORIES), True)
                elif self._guide_calls > MAX_GUIDE_CALLS:
                    self._owe(call, u'Spec limit reached; build the family now.', True)
                else:
                    try:
                        text = self.guide_reader(category) or u''
                    except Exception:
                        text = u''
                    self._owe(call, text or u'No extra spec for {}; follow the contract.'.format(
                        category), False)
            else:
                self._owe(call, u'Unknown tool. Use submit_family only.', True)
        if submit is None and guided:
            self._flush_pending_as_user()
            return Outcome('continue', request=self.build_request(), analysis=analysis)
        if submit is None:
            if stop == 'pause_turn':
                self._append_user_text(u'Continue.')
                return Outcome('continue', request=self.build_request())
            return Outcome('reply', text=text or u'Claude did not return a family. '
                                                  u'Describe the object in more detail.')
        outcome = self._handle_submit(submit, text)
        if analysis is not None:
            outcome.analysis = analysis
        return outcome

    def _handle_submit(self, call, text):
        args = call.get('input') or {}
        family = args.get('family')
        if isinstance(family, str):          # tolerate a JSON string argument
            try:
                family = json.loads(family)
            except ValueError:
                family = None
        summary = (args.get('summary') or text or u'').strip()
        assumptions = [u'{}'.format(a) for a in (args.get('assumptions') or []) if a]
        if not isinstance(family, dict):
            errors = [u'$: family must be a JSON object']
        elif self.auto:
            errors, _ = validate_family_schema(family)
            errors = list(errors) + meaning_errors(family)
        else:
            errors = validate_ai_schema(family, self.category)
        if not errors:
            _, warnings = validate_family_schema(family, None if self.auto else self.category)
            loose = continuity_problems(family)
            if loose and self._polish_left > 0:
                # Valid, but parts float: one polish round before the user sees it.
                self._polish_left -= 1
                self._owe(call, u'Valid, but it does not yet work as one built object:\n- {}\n'
                                u'Re-read your concept (what it is, the route of use) and fix the '
                                u'geometry, not the claims: move or extend parts so real joints overlap '
                                u'1-5 mm, land each access part on both levels, keep routes and bores '
                                u'clear, make containers and conduits hollow. Then call submit_family '
                                u'again with the complete family.'.format(u'\n- '.join(loose)), False)
                self._flush_pending_as_user()
                return Outcome('continue', request=self.build_request(), errors=loose,
                               repairs=self._repairs)
            warnings = list(warnings) + loose
            if self.auto:
                self.detected_category = family.get('family_category')
            self.last_schema = family
            info = schema_summary(family)
            self._owe(call, json.dumps({
                'accepted': True,
                'shown_to_user': True,
                'parts': info.get('parts'), 'materials': len(info.get('materials') or []),
                'warnings': warnings[:20],
                'next': 'Wait for the user. They may ask for changes or create the family.',
            }), False)
            return Outcome('schema', schema=family, summary=summary, assumptions=assumptions,
                           warnings=warnings, repairs=self._repairs)
        shown = errors[:40]
        if self._repairs < self.max_repairs:
            self._repairs += 1
            self._owe(call, u'FamiGen found {} problem(s). Fix every one and call '
                            u'submit_family again with the complete family:\n- {}'.format(
                                len(errors), u'\n- '.join(shown)), True)
            self._flush_pending_as_user()
            return Outcome('continue', request=self.build_request(), errors=errors,
                           repairs=self._repairs)
        self._owe(call, u'Still invalid after {} repair round(s); shown to the user as a '
                        u'draft:\n- {}'.format(self._repairs, u'\n- '.join(shown)), True)
        return Outcome('schema', schema=family if isinstance(family, dict) else None,
                       summary=summary, assumptions=assumptions, errors=errors,
                       repairs=self._repairs)

    # ── transcript helpers ───────────────────────────────────────────────
    def _owe(self, call, text, is_error):
        block = {'type': 'tool_result', 'tool_use_id': call.get('id') or u'',
                 'content': text}
        if is_error:
            block['is_error'] = True
        self._pending_results.append(block)

    def _flush_pending_as_user(self):
        if self._pending_results:
            self.messages.append({'role': 'user', 'content': list(self._pending_results)})
            self._pending_results = []

    def _append_user_text(self, text):
        content = list(self._pending_results) + [{'type': 'text', 'text': text}]
        self._pending_results = []
        self.messages.append({'role': 'user', 'content': content})

    # ── reporting ────────────────────────────────────────────────────────
    def usage_summary(self):
        """Token counts plus a USD estimate (None when the model is unknown)."""
        out = dict(self.usage)
        try:
            from Intelligence.pricing import estimate_usd
            out['usd'] = estimate_usd(self.model, self.usage['input_tokens'],
                                      self.usage['output_tokens'],
                                      self.usage['cache_read_input_tokens'],
                                      self.usage['cache_creation_input_tokens'])
        except Exception:
            out['usd'] = None
        out['model'] = self.model
        return out


def run_turn(session, post, text, current_json=None, attachments=(), max_requests=8):
    """Blocking helper: run one turn with ``post(body) -> (status, raw)``.

    For tests and for callers that already run off the UI thread. Returns the
    final Outcome (never 'continue').
    """
    body = session.begin_turn(text, current_json, attachments)
    for _ in range(max_requests):
        status, raw = post(body)
        outcome = session.handle_http(status, raw)
        if outcome.kind != 'continue':
            return outcome
        body = outcome.request
    return Outcome('error', message=u'Claude needed too many rounds for this request.')
