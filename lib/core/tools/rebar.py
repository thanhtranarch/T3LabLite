# -*- coding: utf-8 -*-
"""Rebar tools as specs: rebar_wizard (Rebar Wizard window, headless).
Doc: docs/agents/headless-tools.md. Test: dev/test_headless_tools.py."""
from __future__ import unicode_literals

from core.tool_specs import ToolSpec
from core.tools._common import MAX_ROWS


def snap_diameters(values, fields, available):
    """Each *_dia to the nearest bar type in the model, as the window does.
    Returns notes for the ones that moved."""
    notes = []
    if not available:
        return notes
    for f in fields:
        if not f.endswith('_dia') or values.get(f) in (None, ''):
            continue
        try:
            asked = float(values[f])
        except (TypeError, ValueError):
            continue
        near = min(available, key=lambda d: abs(d - asked))
        if abs(near - asked) > 0.05:
            notes.append('{}: no bar type of {} mm, used {} mm'.format(f, asked, near))
        values[f] = near
    return notes


def choose_kind(kinds, asked=None):
    """Asked kind, else the most common one (ties: beam, column, footing)."""
    if asked:
        return asked
    order = ('beam', 'column', 'footing')
    counts = dict((k, kinds.count(k)) for k in order)
    return max(order, key=lambda k: (counts[k], -order.index(k))) if kinds else None


def rebar_wizard(doc, args):
    from Snippets import _rebar_wizard as wiz
    from Snippets._compat import disposing, eid_value, elem_name, make_eid, short_error

    kind_asked = args.get('kind')
    if kind_asked and kind_asked not in wiz.KINDS:
        return {'error': 'kind must be one of: {}.'.format(', '.join(wiz.KINDS))}
    elements, ignored = wiz.hosts_from_ids(doc, args.get('element_ids') or [])
    hosts, out_of_scope = [], []
    for el in elements:
        section = wiz.read_section(doc, el)
        label = '{} [{}]'.format(elem_name(el), eid_value(el.Id))
        if section.in_scope:
            hosts.append((el, section, label))
        else:
            out_of_scope.append({'host': label, 'reason': section.reason})
    if not hosts:
        return {'error': 'No rectangular straight beam, rectangular column or pad footing in '
                         'element_ids ({} other element(s) ignored).'.format(ignored),
                'out_of_scope': out_of_scope[:MAX_ROWS]}
    kind = choose_kind([s.kind for _e, s, _l in hosts], kind_asked)
    fields = wiz.WizardInput.field_names(kind)

    presets = wiz.load_presets(wiz.preset_path())
    preset = args.get('preset') or wiz.DEFAULT_PRESET
    if preset != wiz.DEFAULT_PRESET and preset not in presets:
        return {'error': 'No Rebar Wizard preset "{}".'.format(preset),
                'presets': [wiz.DEFAULT_PRESET] + sorted(presets)}
    values = dict((presets.get(preset) or wiz.default_preset())[kind])
    overrides = args.get('values') or {}
    unknown = [k for k in overrides if k not in fields]
    if unknown:
        return {'error': '{} has no field: {}.'.format(kind, ', '.join(unknown)),
                'fields': dict((f, values.get(f)) for f in fields)}
    values.update(overrides)

    bars, hooks = wiz.resolve_types(doc)
    if not bars:
        return {'error': 'The model has no Rebar Bar Type. Load rebar types first.'}
    notes = snap_diameters(values, fields, sorted(bars))
    if 'hook' in values and values['hook'] not in hooks and values['hook'] != wiz.NO_HOOK:
        values['hook'] = wiz.default_hook_name(hooks)
    inp = wiz.WizardInput.from_dict(kind, values)
    problems = inp.validate()
    if problems:
        return {'error': ' '.join(problems), 'fields': inp.to_dict()}

    rows, ready = [], []
    for el, section, label in hosts:
        if section.kind != kind:
            rows.append({'host': label, 'status': 'skipped',
                         'detail': 'a {}, run again with kind={}'.format(section.kind,
                                                                         section.kind)})
            continue
        try:
            plans = wiz.plan_for(section, inp)
        except wiz.LayoutError as exc:
            rows.append({'host': label, 'status': 'failed', 'detail': '{}'.format(exc)})
            continue
        rows.append({'host': label, 'section': section.label(), 'sets': len(plans),
                     'bars': wiz.total_bars(plans), 'summary': wiz.summarize(plans)})
        ready.append((el, section, label, plans, rows[-1]))

    out = {'kind': kind, 'preset': preset, 'fields': inp.to_dict()}
    if notes:
        out['adjusted'] = notes
    if out_of_scope:
        out['out_of_scope'] = out_of_scope[:MAX_ROWS]
    if args.get('dry_run') or not ready:
        out.update(dry_run=bool(args.get('dry_run')), applied=False, hosts=rows[:MAX_ROWS])
        if not ready and not args.get('dry_run'):
            out['error'] = 'No host fits this preset; nothing was created.'
        return out

    from Autodesk.Revit.DB import TransactionGroup
    created = []
    try:
        with disposing(TransactionGroup(doc, 'T3Lab: Rebar Wizard')) as group:
            group.Start()
            for el, section, label, plans, row in ready:
                host = doc.GetElement(make_eid(eid_value(el.Id)))
                if host is None:
                    row.update(status='failed', detail='host no longer exists')
                    continue
                results = wiz.create_plans(doc, host, section, plans, (bars, hooks),
                                           add_to_assembly=inp.add_assembly,
                                           created=created, host_label=label)
                bad = [r for r in results if r.status == wiz.STATUS_FAILED
                       and r.name != 'Assembly']
                if bad:
                    cause = [r for r in bad if r.detail != 'rolled back with the host'] or bad
                    row.update(status='failed', detail=cause[0].detail)
                else:
                    row['status'] = 'ok'
                    asm = [r for r in results if r.name == 'Assembly']
                    if asm:
                        row['assembly'] = asm[0].detail
            if created:
                group.Assimilate()
            else:
                group.RollBack()
    except Exception as exc:
        return {'error': 'Rebar Wizard stopped: {}. The whole run was undone.'.format(
            short_error(exc)), 'hosts': rows[:MAX_ROWS]}
    out.update(success=bool(created), applied=bool(created), created_count=len(created),
               created_rebar_ids=created[:MAX_ROWS], hosts=rows[:MAX_ROWS])
    if not created:
        out['error'] = 'No rebar was created; see hosts.'
    return out


SPECS = [
    ToolSpec(
        'rebar_wizard',
        ('Reinforce rectangular straight beams, rectangular columns or pad footings from a Rebar '
         'Wizard preset (the Rebar Wizard window, without the window). element_ids may be hosts, '
         'their rebar or their assemblies. dry_run returns the bars per host only; an unknown '
         'field in values returns the field list.'),
        rebar_wizard,
        properties={
            'element_ids': {'type': 'array', 'items': {'type': ['integer', 'string']},
                            'description': 'Hosts, their rebar or assemblies; ids or "@set".'},
            'kind': {'type': 'string', 'enum': ['beam', 'column', 'footing'],
                     'description': 'Default: the kind with the most hosts.'},
            'preset': {'type': 'string', 'description': 'Saved preset. Default "Default".'},
            'values': {'type': 'object', 'description': (
                'Field overrides in mm, e.g. beam {"bot_n": 4, "bot_dia": 20, "s_end": 100}. '
                "Diameters snap to the model's bar types.")},
            'dry_run': {'type': 'boolean', 'description': 'Return the bars, create nothing.'},
        },
        required=['element_ids'], effect='write', placement='modify',
        verify=('ids', 'must_change'),
        title_vi='Đặt thép tự động cho dầm, cột, móng',
        keywords=['cốt thép', 'rebar', 'thép dầm', 'thép cột', 'thép móng'],
        ribbon='RebarWizard.pushbutton'),
]
