# -*- coding: utf-8 -*-
"""
CPython 3 test: the Agent Cursors bar, driven for real on a fake WPF.

Run:  python3 lite_guard/test_agent_cursors_overlay.py      Exit 0 = all pass.

OverlayHost.call() swallows every exception, so a NameError in the bar would
only show as "the bar did not update" in Revit. Here core/agent_cursors_overlay
runs against a stand-in `System.Windows` that records the tree it builds:
bands (one colour block per agent, replay fill, sweep), feed, counts, legend,
the cursor tag and what moves on a redraw are read back from that tree. The
Settings window's logic (GUI/AgentCursorsDialog.py) runs on stand-in controls.
The tests run with "On model" on, except test_bar_only (Lite's default).
The stand-in accepts any member name, so real WPF names, layout and pixels
are NOT checked: NEEDS VERIFICATION in Revit.
"""
import ast
import os
import shutil
import sys
import tempfile
import types

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'lib')
sys.path.insert(0, LIB)

_TMP = tempfile.mkdtemp(prefix='t3overlay_')
os.environ['APPDATA'] = _TMP

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok   ' + name)
    else:
        print('  FAIL ' + name + (' -- ' + str(detail) if detail else ''))
        FAILURES.append(name)


# ─── fake WPF ────────────────────────────────────────────────────────────────

CALLS = []          # (static member, args): Grid.SetColumn(text, 1) leaves no other trace


class _Static(object):
    """A static member (enum value, dependency property, factory method)."""

    def __init__(self, name):
        self.name = name

    def __call__(self, *args, **kwargs):
        CALLS.append((self.name, args))
        made = Fake()
        made.__dict__['_made_by'] = (self.name, args)
        return made

    def __repr__(self):
        return self.name


class _Meta(type):
    def __getattr__(cls, name):
        if name.startswith('__'):
            raise AttributeError(name)
        value = _Static('{}.{}'.format(cls.__name__, name))
        setattr(cls, name, value)
        return value


class Fake(metaclass=_Meta):
    """Any .NET object: unknown members appear on first use, a member used as
    a collection is one, `+=` keeps event handlers."""

    def __init__(self, *args, **kwargs):
        self.__dict__['_args'] = args
        self.__dict__['_items'] = []
        self.__dict__['_handlers'] = []
        self.__dict__['_animations'] = {}

    def __getattr__(self, name):
        if name.startswith('__'):
            raise AttributeError(name)
        value = Fake()
        self.__dict__[name] = value
        return value

    def __call__(self, *args, **kwargs):
        return Fake()

    def __iadd__(self, handler):
        self._handlers.append(handler)
        return self

    def Add(self, item):
        self._items.append(item)

    def Insert(self, index, item):
        self._items.insert(index, item)

    def Remove(self, item):
        if item in self._items:
            self._items.remove(item)

    def Clear(self):
        del self._items[:]

    def BeginAnimation(self, prop, animation):
        self._animations[repr(prop)] = animation


class _Window(Fake):
    def __init__(self, *args, **kwargs):
        Fake.__init__(self, *args, **kwargs)
        self.IsVisible = False
        self.IsMouseOver = False
        self.ActualWidth = self.ActualHeight = 0.0

    def Show(self):
        self.IsVisible = True

    def Hide(self):
        self.IsVisible = False


class _Net(types.ModuleType):
    """A System.* namespace: every name is a class."""

    def __getattr__(self, name):
        if name.startswith('__'):
            raise AttributeError(name)
        cls = _Window if name == 'Window' else _Meta(str(name), (Fake,), {})
        setattr(self, name, cls)
        return cls


class _Finder(object):
    def find_module(self, name, path=None):
        return self if name == 'clr' or name.split('.')[0] == 'System' else None

    def find_spec(self, name, path=None, target=None):
        if name == 'clr' or name.split('.')[0] == 'System':
            from importlib.machinery import ModuleSpec
            return ModuleSpec(name, self, is_package=True)
        return None

    def create_module(self, spec):
        module = _Net(spec.name)
        module.__path__ = []
        return module

    def exec_module(self, module):
        pass


sys.meta_path.insert(0, _Finder())

import core.agent_cursors as AC                 # noqa: E402
import core.agent_cursors_overlay as OV         # noqa: E402

AC.set_option('model', True)    # Lite: cursor and boxes are behind "On model" (test_bar_only)


# ─── reading the tree back ───────────────────────────────────────────────────

def kind(node):
    return type(node).__name__


def walk(node):
    """`node` and everything under it (Child, Content, Children)."""
    seen, todo = [], [node]
    while todo:
        cur = todo.pop(0)
        if not isinstance(cur, Fake) or any(cur is s for s in seen):
            continue
        seen.append(cur)
        data = cur.__dict__
        todo.extend(data.get(k) for k in ('Child', 'Content') if k in data)
        if 'Children' in data:
            todo.extend(data['Children']._items)
    return seen


def texts(node):
    return [n.__dict__['Text'] for n in walk(node) if isinstance(n.__dict__.get('Text'), str)]


def brush(value):
    """('#RRGGBB', alpha) of a SolidColorBrush the overlay made."""
    name, args = value._args[0].__dict__['_made_by']
    assert name == 'Color.FromArgb', name
    return '#{:02X}{:02X}{:02X}'.format(*args[1:]), args[0]


def bands(ov):
    """The colour blocks on the open bar, top to bottom."""
    return [n for n in ov._full.Children._items
            if kind(n) == 'Border' and n.__dict__.get('ClipToBounds') is True]


def layers(band):
    return band.Child.Children._items


def stripe_of(band):
    return [n for n in layers(band) if n.__dict__.get('Width') == AC.BAND_STRIPE][0]


def sweeps_of(band):
    return [n for n in layers(band) if 'RenderTransform' in n.__dict__]


def moves(node, prop=None):
    """Animations started on `node` (its own and its RenderTransform's) and
    still running: {property: animation}."""
    found = dict((k, v) for k, v in node._animations.items() if v is not None)
    transform = node.__dict__.get('RenderTransform')
    if isinstance(transform, Fake):
        found.update((k, v) for k, v in transform._animations.items() if v is not None)
    return found if prop is None else found.get(prop)


def glyph_of(band, action):
    return [n for n in walk(band) if n.__dict__.get('Text') == AC.action_style(action)[1]][0]


def frames(anim):
    """[value] of a keyframe animation the overlay built."""
    return [f._args[0] for f in anim.KeyFrames._items]


OPACITY, WIDTH = 'UIElement.OpacityProperty', 'FrameworkElement.WidthProperty'
MOVE_X, MOVE_Y = 'TranslateTransform.XProperty', 'TranslateTransform.YProperty'
SCALE_X = 'ScaleTransform.ScaleXProperty'


def fills_of(band):
    return [n for n in layers(band) if kind(n) == 'Border' and 'Width' in n.__dict__
            and n.__dict__['Width'] != AC.BAND_STRIPE and 'RenderTransform' not in n.__dict__]


RECT = (100.0, 50.0, 1100.0, 850.0)


def snap(boxes=()):
    geo = {'id': 7, 'name': 'Level 1', 'rect': RECT, 'boxes': list(boxes), 'gone': [],
           'ghosts': []}
    return {'rect': RECT, 'views': [geo], 'active': geo, 'view': 'Level 1', 'offscreen': []}


def new_overlay():
    ov = OV._Overlay(1)
    ov._place(RECT)
    return ov


def finish(ov, agent, action, tool, failed=False, **kw):
    ov.show(agent, action, tool, failed, snap(), **kw)


# ─── tests ───────────────────────────────────────────────────────────────────

def test_running_band():
    print('[a running call: one block of its action colour]')
    AC.set_speed('calm')
    ov = new_overlay()
    ov.working('claude-desktop', 'create', 'Creating wall', 'create_wall', RECT, canvas=RECT)
    check('bar is on screen', ov.hud.IsVisible is True)
    found = bands(ov)
    check('one band', len(found) == 1, len(found))
    band = found[0]
    tone = AC.bar_tone(AC.action_style('create')[2], 'dark')
    check('band = the action colour as a wash',
          brush(band.Background) == (tone, AC.BAND_WASH['dark']), brush(band.Background))
    check('solid stripe on the left', brush(stripe_of(band).Background) == (tone, 255))
    words = texts(band)
    check('glyph, caption and the action name on it',
          AC.action_style('create')[1] in words and 'Creating wall' in words and 'Create' in words,
          words)
    check('caption keeps the bar text colour, not the action colour',
          [brush(n.Foreground)[0] for n in walk(band) if n.__dict__.get('Text') == 'Creating wall']
          == [ov._c('text')])
    check('running: one sweep crosses the band', len(sweeps_of(band)) == 1)
    sweep = sweeps_of(band)[0]
    anim = moves(sweep, MOVE_X)
    check('sweep repeats for ever', anim is not None
          and repr(anim.__dict__.get('RepeatBehavior')) == 'RepeatBehavior.Forever')
    check('running: the stripe breathes',
          repr(moves(stripe_of(band), OPACITY).__dict__.get('RepeatBehavior'))
          == 'RepeatBehavior.Forever')
    check('one agent: no agent mark on its band',
          not any(kind(n) == 'Path' for n in walk(band)))
    cur = ov._cursors['claude-desktop']
    check('cursor tag: the action icon only, in the tag\'s ink (no block, no colour)',
          cur['glyph'].Text == AC.action_style('create')[1]
          and brush(cur['glyph'].Foreground)[0] == '#FFFFFF'
          and repr(cur['glyph'].Visibility) == 'Visibility.Visible'
          and 'chip' not in cur and cur['glyph'] in cur['pill'].Child.Children._items)
    check('tag keeps no coloured edge', 'BorderBrush' not in cur['pill'].__dict__)
    check('a new cursor fades in and its tag grows out of the tip',
          frames(moves(cur['group'], OPACITY)) == [0.0, 1.0]
          and frames(moves(cur['pill'], SCALE_X)) == [0.6, 1.0] and 'new' not in cur)

    old_loops = list(ov._loops)
    ov._refresh_hud()
    check('rebuilding the bar stops every endless animation of the old one',
          old_loops and all(t._animations[repr(p)] is None for t, p in old_loops)
          and moves(sweep, MOVE_X) is None)
    ov._folded = True
    ov._refresh_hud()
    check('folded bar: no sweep, only its status dot pulses',
          len(ov._loops) == 1 and kind(ov._loops[0][0]) == 'Ellipse')
    ov._folded = False

    AC.set_speed('off')
    try:
        ov._drawn = {}
        ov._full._animations.clear()        # the stand-in never ends an animation
        ov._refresh_hud()
        check('Animation Off (hidden setting): nothing on the bar moves',
              len(bands(ov)) == 1 and sweeps_of(bands(ov)[0]) == [] and ov._loops == []
              and not any(moves(n) for n in walk(ov._full)))
    finally:
        AC.set_speed('calm')


def test_motion_on_change_only():
    print('[motion: a band moves when its step changes, never on a plain redraw]')
    AC.set_speed('calm')
    ov = new_overlay()
    ov.working('claude-desktop', 'read', 'Listing views', 'list_views', RECT, canvas=RECT)
    band = bands(ov)[0]
    check('first band slides in and fades in',
          frames(moves(band, MOVE_Y)) == [6.0, 0.0] and frames(moves(band, OPACITY)) == [0.0, 1.0])
    check('the bar itself fades in when it first shows',
          frames(moves(ov._full, OPACITY)) == [0.0, 1.0])
    ov._full._animations.clear()            # the stand-in never ends an animation
    ov._refresh_hud()
    band = bands(ov)[0]
    check('plain redraw (hover): the band does not move again',
          moves(band, MOVE_Y) is None and moves(band, OPACITY) is None
          and moves(glyph_of(band, 'read'), SCALE_X) is None and not moves(ov._full))

    finish(ov, 'claude-desktop', 'read', 'list_views')
    band = bands(ov)[0]
    check('step done: its icon pops, the band stays put',
          frames(moves(glyph_of(band, 'read'), SCALE_X)) == [1.45, 1.0]
          and moves(band, MOVE_Y) is None and moves(band, MOVE_X) is None)

    ov.working('claude-desktop', 'delete', 'Deleting element', 'delete_element', RECT, canvas=RECT)
    rows = [n for n in walk(ov._full) if kind(n) == 'Grid' and 'Listing views' in texts(n)
            and n not in walk(bands(ov)[0])]
    check('the finished step drops from the band into the feed',
          rows and frames(moves(rows[-1], MOVE_Y)) == [-6.0, 0.0], len(rows))
    cur = ov._cursors['claude-desktop']
    check('cursor tag: the icon pops when the action changes',
          cur['glyph'].Text == AC.action_style('delete')[1]
          and frames(moves(cur['glyph'], SCALE_X)) == [1.6, 1.0])
    finish(ov, 'claude-desktop', 'delete', 'delete_element', failed=True)
    band = bands(ov)[0]
    check('failed step: the band shakes', frames(moves(band, MOVE_X)) == [0.0, -6.0, 6.0, -4.0, 0.0])
    ov._refresh_hud()
    check('and not again on a redraw', moves(bands(ov)[0], MOVE_X) is None)

    ov.working('codex', 'create', 'Creating wall', 'create_wall', RECT, canvas=RECT)
    live = bands(ov)
    check('a second agent: only its band slides in',
          len(live) == 1 and frames(moves(live[0], MOVE_Y)) == [6.0, 0.0])
    finish(ov, 'codex', 'create', 'create_wall')

    def segments():
        return [n for n in walk(ov._full) if kind(n) == 'Border' and n.__dict__.get('Height') == 6]
    parts = segments()      # claude: read, delete (drawn before) · codex: create (new)
    check('counts bar: a new part grows from nothing, the others stay still',
          [bool(moves(n, WIDTH)) for n in parts] == [False, False, True]
          and frames(moves(parts[2], WIDTH)) == [0.0, AC.COUNT_BAR_W],
          [moves(n, WIDTH) for n in parts])
    half = parts[0].Width
    ov.working('claude-desktop', 'read', 'Listing views', 'list_views', RECT, canvas=RECT)
    finish(ov, 'claude-desktop', 'read', 'list_views')
    parts = segments()
    check('counts changed: each part grows or shrinks from the width it had',
          frames(moves(parts[0], WIDTH)) == [half, parts[0].Width] and parts[0].Width > half
          and frames(moves(parts[1], WIDTH)) == [half, parts[1].Width] and parts[1].Width < half)
    ov._refresh_hud()
    check('and stays still on a plain redraw', not any(moves(n, WIDTH) for n in segments()))

    ov.replay()
    n = len(ov._history)
    fill = fills_of(bands(ov)[0])[0]
    check('replay: the fill grows from 0 to step 1',
          frames(moves(fill, WIDTH)) == [0.0, AC.BAR_W / n])
    ov._on_replay_tick(None, None)
    fill = fills_of(bands(ov)[0])[0]
    check('then from step 1 to step 2',
          frames(moves(fill, WIDTH)) == [AC.BAR_W / n, 2 * AC.BAR_W / n])
    ov.hide()

    ov = new_overlay()
    ov.working('claude-desktop', 'read', 'Listing views', 'list_views', RECT, canvas=RECT)
    finish(ov, 'claude-desktop', 'read', 'list_views')
    callbacks = []
    ov.host = types.SimpleNamespace(callbacks={'idle': lambda: callbacks.append('idle')})
    ov._enter_summary()
    fade = moves(ov.canvas, OPACITY)
    check('agents idle: the cursors fade out, Glow goes off at once',
          frames(fade) == [1.0, 0.0] and callbacks == ['idle'] and ov._cursors
          and ov.win.IsVisible is True)
    ov.working('claude-desktop', 'read', 'Listing views', 'list_views', RECT, canvas=RECT)
    check('a call during the fade keeps its cursor',
          ov._leaving is None and moves(ov.canvas, OPACITY) is None and ov.canvas.Opacity == 1.0
          and 'claude-desktop' in ov._cursors)
    for handler in fade.Completed._handlers:
        handler(None, None)
    check('the stale fade does not clear the views afterwards',
          'claude-desktop' in ov._cursors and ov.win.IsVisible is True)
    finish(ov, 'claude-desktop', 'read', 'list_views')
    ov._enter_summary()
    for handler in moves(ov.canvas, OPACITY).Completed._handlers:
        handler(None, None)
    check('fade finished: cursors and boxes gone, canvas window hidden',
          ov._cursors == {} and ov.canvas.Children._items == [] and ov.win.IsVisible is False
          and ov._leaving is None and ov.canvas.Opacity == 1.0)


def test_several_agents():
    print('[several agents: a block each, in agent order]')
    ov = new_overlay()
    ov.working('codex', 'delete', 'Deleting element', 'delete_element', RECT, canvas=RECT)
    ov.working('claude-desktop', 'read', 'Listing views', 'list_views', RECT, canvas=RECT)
    found = bands(ov)
    check('two bands', len(found) == 2, len(found))
    check('agent order, whoever started first',
          ['Listing views' in texts(found[0]), 'Deleting element' in texts(found[1])] == [True, True],
          [texts(b) for b in found])
    check('each in its own action colour',
          [brush(b.Background)[0] for b in found]
          == [AC.bar_tone(AC.action_style(a)[2], 'dark') for a in ('read', 'delete')])
    check('2+ agents: each band carries its agent mark',
          all(any(kind(n) == 'Path' for n in walk(b)) for b in found))
    for k, name in enumerate(('a-bot', 'b-bot', 'c-bot')):
        ov.working(name, 'modify', 'Moving {}'.format(k), 'move_elements', RECT, canvas=RECT)
    check('at most BAND_MAX bands, the rest is counted',
          len(bands(ov)) == AC.BAND_MAX and '+2 more' in texts(ov._full),
          texts(ov._full))

    finish(ov, 'claude-desktop', 'read', 'list_views')
    check('a finished call leaves its band to the calls still running',
          'Listing views' not in [t for b in bands(ov) for t in texts(b)]
          and 'Listing views' in texts(ov._full))


def test_last_step_and_feed():
    print('[idle: the last step is the band, the feed lists the ones before]')
    ov = new_overlay()
    for tool in ('list_views', 'create_wall', 'set_parameter'):
        action = AC.classify_action(tool)
        ov.working('claude-desktop', action, AC.describe_call(tool), tool, RECT, canvas=RECT)
        finish(ov, 'claude-desktop', action, tool, changes=[(1, 'Mark: A → B')]
               if tool == 'set_parameter' else ())
    found = bands(ov)
    check('one band: the newest step', len(found) == 1 and 'Setting parameter' in texts(found[0]),
          [texts(b) for b in found])
    check('done: no sweep, no fill, nothing loops but the status dot',
          sweeps_of(found[0]) == [] and fills_of(found[0]) == []
          and [kind(t) for t, _p in ov._loops] == ['Ellipse'])
    check('before → after chips sit on the band', 'Mark: A → B' in texts(found[0]))
    feed = [t for t in texts(ov._full) if t in ('Listing views', 'Creating wall', 'Setting parameter')]
    check('the feed does not repeat the step on the band',
          feed.count('Setting parameter') == 1 and 'Creating wall' in feed and 'Listing views' in feed,
          feed)
    stripes = [n for n in walk(ov._full) if kind(n) == 'Border' and n.__dict__.get('Width') == 3
               and n.__dict__.get('Height') == 12]
    check('feed rows carry a stripe in their action colour',
          [brush(n.Background)[0] for n in stripes]
          == [AC.bar_tone(AC.action_style(a)[2], 'dark') for a in ('create', 'read')],
          [brush(n.Background)[0] for n in stripes])

    ov.working('claude-desktop', 'delete', 'Deleting element', 'delete_element', RECT, canvas=RECT)
    finish(ov, 'claude-desktop', 'delete', 'delete_element', failed=True)
    band = bands(ov)[0]
    bad = AC.bar_tone(ov._c('bad'), 'dark')
    check('failed step: red block with a cross, still named Delete',
          brush(band.Background)[0] == bad and '' in texts(band) and 'Delete' in texts(band)
          and any(t.startswith('Failed: ') for t in texts(band)), texts(band))

    ov.set_theme('light')
    band = bands(ov)[0]
    check('light theme: lighter wash, tone made for the light bar',
          brush(band.Background) == (AC.bar_tone(AC.HUD_PALETTES['light']['bad'], 'light'),
                                     AC.BAND_WASH['light']), brush(band.Background))


def test_waiting():
    print('[waiting for the user: amber block, unless a card already asks]')
    ov = new_overlay()
    ov.working('claude-desktop', 'delete', 'Deleting element', 'delete_element', RECT, canvas=RECT)
    ov.waiting('claude-desktop', 'delete_element', 'Deleting element', True)
    band = bands(ov)[0]
    check('amber, says Waiting',
          brush(band.Background)[0] == AC.bar_tone(ov._c('warn'), 'dark')
          and 'Waiting' in texts(band) and sweeps_of(band) == [], texts(band))
    ov.waiting('claude-desktop', 'delete_element', 'Deleting element', False)
    check('OK given: back to the running block',
          brush(bands(ov)[0].Background)[0] == AC.bar_tone(AC.action_style('delete')[2], 'dark'))
    finish(ov, 'claude-desktop', 'delete', 'delete_element')
    check('call over: nothing kept for that agent', ov._acting == {} and ov._working == {})

    ov.ask('r1', 'codex', 'delete_element', 'Deleting element', 'Delete 3 walls?', '', '', '', RECT)
    check('a call with an Allow / Deny card gets no second block',
          not any('Waiting' in texts(b) for b in bands(ov)), [texts(b) for b in bands(ov)])


def test_replay():
    print('[replay: the band fills to k / n, the feed is left alone]')
    ov = new_overlay()
    tools = ('list_views', 'create_wall', 'delete_element')
    for tool in tools:
        action = AC.classify_action(tool)
        ov.working('claude-desktop', action, AC.describe_call(tool), tool, RECT, canvas=RECT)
        finish(ov, 'claude-desktop', action, tool)
    feed_before = [(s['caption'], s['t']) for s in ov._feed]
    counts_before = dict(ov._counts['claude-desktop'])

    ov.working('codex', 'read', 'Listing views', 'list_views', RECT, canvas=RECT)
    ov.replay()
    check('no replay while a call runs, and the footer says why',
          ov._replay is None and 'Replay' in ov._note, ov._note)
    finish(ov, 'codex', 'read', 'list_views')
    feed_before = [(s['caption'], s['t']) for s in ov._feed]

    ov.replay()
    n = len(ov._history)
    band = bands(ov)[0]
    check('step 1: its own caption and colour, "Replay 1 / n" on the right',
          'Listing views' in texts(band) and 'Replay 1 / {}'.format(n) in texts(band)
          and brush(band.Background)[0] == AC.bar_tone(AC.action_style('read')[2], 'dark'),
          texts(band))
    fill = fills_of(band)
    check('the block is filled to 1 / n', len(fill) == 1
          and abs(fill[0].Width - AC.BAR_W / n) < 1e-6, [f.__dict__.get('Width') for f in fill])
    check('replay: no sweep', sweeps_of(band) == [])
    ov._on_replay_tick(None, None)
    band = bands(ov)[0]
    check('step 2 fills further and changes colour',
          abs(fills_of(band)[0].Width - 2 * AC.BAR_W / n) < 1e-6
          and brush(band.Background)[0] == AC.bar_tone(AC.action_style('create')[2], 'dark'))
    check('the feed keeps its real steps and times (2026-11-19: every row said "now")',
          [(s['caption'], s['t']) for s in ov._feed] == feed_before)
    check('a replay counts nothing twice', ov._counts['claude-desktop'] == counts_before)
    for _ in range(n):
        ov._on_replay_tick(None, None)
    check('done: footer note, band back on the last real step',
          ov._replay is None and ov._replay_step is None
          and ov._note == 'Replay done · {} steps'.format(n)
          and fills_of(bands(ov)[0]) == [], ov._note)

    ov.replay()
    ov.working('codex', 'modify', 'Moving elements', 'move_elements', RECT, canvas=RECT)
    check('a live call ends a replay', ov._replay is None and ov._replay_step is None
          and 'Moving elements' in texts(bands(ov)[0]) and fills_of(bands(ov)[0]) == [])


def test_counts_drawer_and_folded():
    print('[counts bar, colour legend, animation control, one-line bar]')
    ov = new_overlay()
    for agent, tool in (('claude-desktop', 'list_views'), ('claude-desktop', 'list_views'),
                        ('claude-desktop', 'create_wall'), ('codex', 'delete_element')):
        action = AC.classify_action(tool)
        ov.working(agent, action, AC.describe_call(tool), tool, RECT, canvas=RECT)
        finish(ov, agent, action, tool)
    segments = [n for n in walk(ov._full) if kind(n) == 'Border' and n.__dict__.get('Height') == 6]
    check('each agent: its counts as one stacked bar (2 actions + 1 action)',
          len(segments) == 3, len(segments))
    claude = segments[:2]
    check('segments in the action colours, read twice as wide as create',
          [brush(n.Background)[0] for n in claude]
          == [AC.bar_tone(AC.action_style(a)[2], 'dark') for a in ('read', 'create')]
          and abs(claude[0].Width - 2 * claude[1].Width) < 1e-6,
          [n.__dict__.get('Width') for n in claude])
    check('bar width = COUNT_BAR_W, gap included',
          abs(sum(n.Width for n in claude) + 1.0 - AC.COUNT_BAR_W) < 1e-6)

    ov._drawer = True
    ov._refresh_hud()
    words = texts(ov._full)
    check('drawer: no motion control at all (owner, 2026-11-19: motion is not an option)',
          not any(w in words for w in ('Animation', 'Speed', 'Less motion', 'Calm', 'Fast')), words)
    check('drawer: every option of OPTION_GROUPS as a switch row',
          all(AC.OPTION_LABELS[n][0] in words for _g, names in AC.OPTION_GROUPS for n in names))
    legend = [n for n in walk(ov._full) if kind(n) == 'WrapPanel'][-1]
    check('drawer: legend = icon + name per action, no colour swatch',
          'ACTIONS' in words and all(label in texts(legend) and glyph in texts(legend)
                                     for _k, label, glyph, _c in AC.ACTIONS)
          and not any(kind(n) == 'Border' for n in walk(legend)), texts(legend))
    check('legend icons carry the action colour as drawn on this bar',
          [brush(n.Foreground)[0] for n in walk(legend)
           if n.__dict__.get('Text') in [a[2] for a in AC.ACTIONS]]
          == [AC.bar_tone(a[3], 'dark') for a in AC.ACTIONS])
    ov._drawer = False

    ov._folded = True
    ov._refresh_hud()
    check('one-line bar shows the last step', 'Deleting element' in texts(ov._compact),
          texts(ov._compact))
    ov._folded = False
    ov.replay()
    ov._folded = True
    ov._refresh_hud()
    check('one-line bar during a replay: "Replay k / n · step"',
          any(t.startswith('Replay 1 / 4 ·') for t in texts(ov._compact)), texts(ov._compact))
    ov.hide()
    check('hide(): nothing left over', ov._loops == [] and ov._acting == {} and ov._drawn == {}
          and ov._replay is None and ov._replay_step is None and ov.hud.IsVisible is False)


def test_call_always_ends():
    print('[flow: a call that started always ends on the bar]')
    AC.set_speed('calm')
    ov = new_overlay()
    ov.working('claude-desktop', 'view', 'Closing document', 'close_document', RECT, canvas=RECT)
    check('running: Working, with its block', ov._state() == 'working' and len(bands(ov)) == 1
          and len(sweeps_of(bands(ov)[0])) == 1)
    ov.done('claude-desktop', 'view', 'close_document', False)
    band = bands(ov)[0]
    check('done() with no view left: the step is on the bar, nothing keeps running',
          ov._working == {} and ov._acting == {} and 'Closing document' in texts(band)
          and sweeps_of(band) == [] and ov._actions == 1
          and ov._counts == {'claude-desktop': {'view': 1}} and ov._history == [])
    ov._feed[0]['t'] -= 60
    check('so the bar goes idle and Replay is not refused', ov._state() == 'idle')
    cur = ov._cursors['claude-desktop']
    check('its cursor stops spinning', repr(cur['spin'].Visibility) == 'Visibility.Collapsed')

    ov.working('codex', 'read', 'Listing open documents', 'list_open_documents', None)
    check('working() with no view: the bar it already has says Working, no cursor is made',
          ov._state() == 'working' and 'codex' not in ov._cursors
          and 'Listing open documents' in texts(bands(ov)[0]))
    ov.done('codex', 'read', 'list_open_documents', True)
    check('and done() marks a failed call as failed',
          ov._working == {} and ov._failed == 1 and '\uE711' in texts(bands(ov)[0]))

    fresh = OV._Overlay(1)
    fresh.working('codex', 'read', 'Listing open documents', 'list_open_documents', None)
    fresh.done('codex', 'read', 'list_open_documents', False)
    check('a bar that never had a view to sit in stays away',
          fresh.hud.IsVisible is False and fresh._working == {} and fresh._actions == 1)

    ov = new_overlay()
    ov.working('claude-desktop', 'create', 'Creating wall', 'create_wall', RECT, canvas=RECT,
               targets=[(300.0, 200.0, 500.0, 300.0)])
    frames_before = len(ov._focus.get('claude-desktop', []))

    def broken(*args, **kwargs):
        raise RuntimeError('a WPF call failed')
    ov._draw_step = broken
    finish(ov, 'claude-desktop', 'create', 'create_wall')
    check('drawing on the model fails: the bar still gets the step and stops saying Working',
          frames_before == 1 and ov._working == {} and ov._focus == {} and ov._actions == 1
          and 'Creating wall' in texts(bands(ov)[0]) and sweeps_of(bands(ov)[0]) == [])
    ov._draw_working = broken
    ov.working('claude-desktop', 'read', 'Listing views', 'list_views', RECT, canvas=RECT)
    check('drawing the running call fails: the bar still says Working',
          'Listing views' in texts(bands(ov)[0]) and len(sweeps_of(bands(ov)[0])) == 1)

    AC.set_option('model', False)       # the default here: the bar only
    try:
        ov = new_overlay()
        ov.working('codex', 'view', 'Closing document', 'close_document', RECT, canvas=RECT)
        ov.done('codex', 'view', 'close_document', False)
        check('bar only: done() ends the call, no cursor is made',
              ov._working == {} and ov._cursors == {} and ov._actions == 1
              and sweeps_of(bands(ov)[0]) == [] and ov.win.IsVisible is False)
        ov.hide()
    finally:
        AC.set_option('model', True)


def test_bar_only():
    print('[bar only (On model off, the default): bands and motion on the bar, nothing on the views]')
    AC.set_option('model', False)
    try:
        ov = new_overlay()
        ov.working('claude-desktop', 'create', 'Creating wall', 'create_wall', RECT, canvas=RECT)
        band = bands(ov)[0]
        check('running: a block of the action colour, with its sweep',
              brush(band.Background)[0] == AC.bar_tone(AC.action_style('create')[2], 'dark')
              and 'Creating wall' in texts(band) and 'Create' in texts(band)
              and len(sweeps_of(band)) == 1)
        check('no cursor, canvas window hidden', ov._cursors == {} and ov.win.IsVisible is False)
        finish(ov, 'claude-desktop', 'create', 'create_wall', failed=True)
        band = bands(ov)[0]
        check('failed step: red block that shakes',
              brush(band.Background)[0] == AC.bar_tone(ov._c('bad'), 'dark')
              and frames(moves(band, MOVE_X)) == [0.0, -6.0, 6.0, -4.0, 0.0])
        check('still no cursor', ov._cursors == {} and ov._acting == {})
        ov.waiting('codex', 'delete_element', 'Deleting element', True)
        check('waiting: an amber block, no cursor made',
              'Waiting' in texts(bands(ov)[0]) and ov._cursors == {})
        ov.waiting('codex', 'delete_element', 'Deleting element', False)
        ov.replay()
        check('replay still fills its block', len(fills_of(bands(ov)[0])) == 1
              and ov._cursors == {})
        ov.hide()
    finally:
        AC.set_option('model', True)


def test_effect_switches():
    print('[Effects: Trail, Ripple and Labels each switch their own drawing off]')
    AC.set_speed('calm')
    wall = (300.0, 200.0, 500.0, 300.0)

    def draw(**switches):
        for name in ('trail', 'ripple', 'labels'):
            AC.set_option(name, switches.get(name, True))
        ov = new_overlay()
        ov.working('claude-desktop', 'create', 'Creating wall', 'create_wall', RECT, canvas=RECT,
                   targets=[wall])
        box = {'id': 1, 'box': wall, 'size': '1200 x 300', 'label': 'Walls \u00b7 1200 x 300'}
        ov.show('claude-desktop', 'create', 'create_wall', False, snap([box]),
                changes=[(1, 'Mark: A \u2192 B')], created='Walls \u00d71 created', fresh=[1])
        top = ov.canvas.Children._items
        cur = ov._cursors['claude-desktop']
        groups = [n for n in top if kind(n) == 'Canvas' and n is not cur['group']]
        return {'lines': [n for n in top if kind(n) == 'Line'],
                'rings': [n for n in top if kind(n) == 'Ellipse'],
                'chips': [n for n in top if kind(n) == 'Border'],
                'badges': [c for g in groups for c in g.Children._items if kind(c) == 'Border'],
                'frames': [c for g in groups for c in g.Children._items if kind(c) == 'Rectangle'],
                'click': bool(moves(cur['press'], SCALE_X)), 'texts': [t for n in top for t in texts(n)]}
    try:
        on = draw()
        check('all on: path to the target, click ring and squeeze, badge + created label + change',
              on['lines'] and on['rings'] and on['click'] and len(on['badges']) == 1
              and len(on['chips']) == 2 and 'Walls \u00d71 created' in on['texts']
              and 'Mark: A \u2192 B' in on['texts'] and 'Walls \u00b7 1200 x 300' in on['texts'],
              dict((k, len(v) if isinstance(v, list) else v) for k, v in on.items()))
        off = draw(trail=False)
        check('Trail off: no path, the rest stays',
              not off['lines'] and off['rings'] and off['click'] and off['badges'] and off['chips'])
        off = draw(ripple=False)
        check('Ripple off: no ring and no click of the pointer, the rest stays',
              not off['rings'] and not off['click'] and off['lines'] and off['badges'])
        off = draw(labels=False)
        check('Labels off: no badge, created label or changed value; the box itself stays',
              not off['badges'] and not off['chips'] and off['frames'] and off['lines']
              and off['rings'] and 'Walls \u00d71 created' not in off['texts'])
        check('the bar still names the step and its change with Labels off',
              ov_texts_have_change(), 'the change chip left the bar')
    finally:
        for name in ('trail', 'ripple', 'labels'):
            AC.set_option(name, True)


def ov_texts_have_change():
    """Labels switches the model's labels only: the band keeps its change chips."""
    AC.set_option('labels', False)
    ov = new_overlay()
    ov.working('claude-desktop', 'modify', 'Setting parameter', 'set_parameter', RECT, canvas=RECT)
    finish(ov, 'claude-desktop', 'modify', 'set_parameter', changes=[(1, 'Mark: A \u2192 B')])
    return 'Mark: A \u2192 B' in texts(bands(ov)[0])


def settings_window(options, enabled=False):
    """The window class on stand-in controls (the XAML itself needs WPF)."""
    import GUI.AgentCursorsDialog as SD
    win = SD.AgentCursorsSettingsWindow.__new__(SD.AgentCursorsSettingsWindow)
    for name in ('chk_enabled', 'btn_ok', 'row_model', 'card_options', 'pnl_benefits',
                 'pnl_drawbacks', 'details_body', 'details_arrow', 'opt_model', 'opt_follow',
                 'opt_zoom', 'opt_tile', 'opt_select', 'opt_glow', 'opt_trail', 'opt_ripple',
                 'opt_labels', 'opt_keep'):
        setattr(win, name, Fake())
    from System.Windows import WindowState
    win.Height, win.WindowState = SD.HEIGHT, WindowState.Normal
    win.Close = lambda: None
    win.FindResource = lambda key: key
    win.chk_enabled.IsChecked = True
    win._setup(enabled, options, AC.pros_cons_sections())
    return SD, win


def test_settings_window():
    print('[settings window: ticks in and out, what greys out, the folded card]')
    options = dict((n, d) for n, _k, d in AC.OPTIONS if n != 'model')
    options['glow'] = False
    SD, win = settings_window(options)
    check('opened to turn it on: ticked, button says Turn On',
          win.chk_enabled.IsChecked is True and win.btn_ok.Content == 'Turn On')
    check('each option ticked as saved',
          [(n, win._boxes[n].IsChecked) for n in sorted(options)]
          == [(n, options[n]) for n in sorted(options)])
    check('no "On model" option here: its row stays hidden, every box is live',
          'Visibility' not in win.row_model.__dict__ and 'model' not in win._boxes
          and all(b.IsEnabled is True for b in win._boxes.values()))
    check('the card of a ticked feature is inked, like a ticked format in BatchOut',
          win.card_options.BorderBrush == 'T3.Ink')
    shown = texts(win.pnl_benefits) + texts(win.pnl_drawbacks)
    check('benefits left, drawbacks right: one bullet row per line',
          [t for t in texts(win.pnl_benefits) if t in AC.PROS] == list(AC.PROS)
          and [t for t in texts(win.pnl_drawbacks) if t in AC.CONS] == list(AC.CONS)
          and shown.count('\uE73E') == len(AC.PROS) and shown.count('\uE7BA') == len(AC.CONS))

    win.opt_follow.IsChecked = False
    win.ok_clicked(None, None)
    want = dict(options, follow=False, enabled=True)
    check('Turn On returns every option and the switch', win.result == want, win.result)

    win.chk_enabled.IsChecked = False
    win.enabled_clicked(None, None)
    check('feature unticked: every option greys out and keeps its tick, the card loses its ink',
          all(b.IsEnabled is False for b in win._boxes.values())
          and win.opt_zoom.IsChecked is True and win.card_options.BorderBrush == 'T3.Border')
    check('saving it unticked turns the feature off, options as they were',
          win.values() == dict(want, enabled=False))
    win.cancel_clicked(None, None)
    check('Cancel returns nothing', win.result is None)

    check('the details card starts folded', 'Visibility' not in win.details_body.__dict__)
    win.details_header_clicked(None, None)
    check('header clicked: card open, arrow up, window grows',
          repr(win.details_body.Visibility) == 'Visibility.Visible'
          and win.details_arrow.Text == '\uE70E' and win.Height == SD.HEIGHT_OPEN)
    win.details_header_clicked(None, None)
    check('clicked again: folded, arrow down, window back',
          repr(win.details_body.Visibility) == 'Visibility.Collapsed'
          and win.details_arrow.Text == '\uE70D' and win.Height == SD.HEIGHT)

    from System.Windows import WindowState
    win.WindowState = WindowState.Maximized
    win.details_header_clicked(None, None)
    check('maximized: the card opens, the height it restores to is left alone',
          repr(win.details_body.Visibility) == 'Visibility.Visible' and win.Height == SD.HEIGHT)

    lite = dict(options, model=False, glow=True)        # T3LabLite: OPTIONS has 'model'
    SD, win = settings_window(lite, enabled=True)
    check('already on: the button says Save', win.btn_ok.Content == 'Save')
    check('"On model" row shown when the option exists, the window taller by that row',
          repr(win.row_model.Visibility) == 'Visibility.Visible'
          and win.Height == SD.HEIGHT + SD.MODEL_ROW)
    win.details_header_clicked(None, None)
    win.details_header_clicked(None, None)
    check('folding the details card again gives that height back',
          win.Height == SD.HEIGHT + SD.MODEL_ROW)
    check('"On model" off: what draws on the views greys out, Keep bar stays',
          [n for n in sorted(win._boxes) if win._boxes[n].IsEnabled]
          == ['keep', 'model'], dict((n, b.IsEnabled) for n, b in win._boxes.items()))
    win.opt_model.IsChecked = True
    win.model_clicked(None, None)
    check('"On model" ticked: every option is live',
          all(b.IsEnabled is True for b in win._boxes.values()))
    check('an option the XAML has no checkbox for is passed through untouched',
          settings_window(dict(options, future=True))[1].values()['future'] is True)
    try:
        SD.ask(False, options)
        built = True
    except Exception:
        built = False
    check('no WPF: ask() raises, so the button falls back to the plain dialog',
          built is False or SD._Window is not object)


def test_ironpython_syntax():
    print('[both modules stay IronPython 2.7 syntax]')
    for name in ('agent_cursors.py', 'agent_cursors_overlay.py'):
        with open(os.path.join(LIB, 'core', name), encoding='utf-8') as f:
            tree = ast.parse(f.read())
        bad = []
        for node in ast.walk(tree):
            label = None
            if isinstance(node, (ast.JoinedStr, ast.AnnAssign, ast.Nonlocal, ast.NamedExpr,
                                 ast.AsyncFunctionDef, ast.Await, ast.YieldFrom, ast.MatMult)):
                label = type(node).__name__
            elif isinstance(node, (ast.FunctionDef, ast.Lambda)):
                args = node.args
                if args.kwonlyargs or args.posonlyargs or getattr(node, 'returns', None) or any(
                        a.annotation for a in args.args):
                    label = 'py3-only signature'
            elif isinstance(node, ast.Raise) and node.cause is not None:
                label = 'raise ... from'
            elif isinstance(node, ast.Starred) and not isinstance(
                    getattr(node, 'ctx', None), ast.Load):
                label = 'starred assignment'
            elif isinstance(node, (ast.List, ast.Tuple, ast.Set)) and any(
                    isinstance(e, ast.Starred) for e in node.elts):
                label = 'unpacking in a literal'
            elif isinstance(node, ast.Dict) and None in node.keys:
                label = '** in a dict literal'
            elif isinstance(node, ast.ClassDef) and node.keywords:
                label = 'class keywords'
            elif (isinstance(node, ast.Call) and isinstance(node.func, ast.Name)
                  and node.func.id == 'super' and not node.args):
                label = 'super() without arguments'
            if label:
                bad.append('{} line {}'.format(label, getattr(node, 'lineno', '?')))
        check(name + ': no py3-only syntax', not bad, bad[:5])
        with open(os.path.join(LIB, 'core', name), encoding='utf-8') as f:
            head = f.read(2000)
        check(name + ': unicode_literals (glyph and "·" literals are text under 2.7)',
              'from __future__ import unicode_literals' in head)


if __name__ == '__main__':
    try:
        test_running_band()
        test_motion_on_change_only()
        test_several_agents()
        test_last_step_and_feed()
        test_waiting()
        test_replay()
        test_counts_drawer_and_folded()
        test_call_always_ends()
        test_bar_only()
        test_effect_switches()
        test_settings_window()
        test_ironpython_syntax()
    finally:
        shutil.rmtree(_TMP, ignore_errors=True)
    print()
    if FAILURES:
        print('FAILED: {}'.format(', '.join(FAILURES)))
        sys.exit(1)
    print('All agent cursor overlay tests passed.')
