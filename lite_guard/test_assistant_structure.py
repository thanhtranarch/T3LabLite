#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Safety net for the T3Lab Assistant window split.

    python3 lite_guard/test_assistant_structure.py

lib/GUI/T3LabAssistantDialog.py was one 11,930-line module: ~1,100 lines of
module-level helpers and one window class with ~260 methods. It is now the
window class plus lib/GUI/assistant/ (module-level helpers, and plain helper
classes holding groups of window methods: `self._sessions =
AssistantSessions(self)`, the moved code reaching the window as `self.win`).

Nobody can run Revit here, so these tests read the code instead:

  a. every event handler named in T3LabAssistant.xaml is a method of
     T3LabAssistantWindow (T3WPFWindow wires `Click="name"` with
     getattr(window, name)); every delegating stub reaches a real method
  b. every name the dialog module bound before the split is still importable
     from GUI.T3LabAssistantDialog, and every name other modules import from
     it exists
  c. every `self.<name>` the window and its helpers use resolves to a method
     or attribute (catches a method lost or mis-routed by the move)
  d. the window inherits ONLY T3WPFWindow (a CLR base plus a Python mixin
     crashes Revit, CLAUDE.md rule 6); helpers inherit only `object`
  e. every module compiles, and imports under plain CPython with the .NET /
     pyRevit modules stubbed; re-exported names are the same objects

Plain CPython 3, standard library only — no WPF, no clr, no Revit.
"""
import ast
import builtins
import importlib
import io
import os
import py_compile
import re
import symtable
import sys
import tempfile
import types
import unittest

sys.dont_write_bytecode = True

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(ROOT, 'lib')
GUI = os.path.join(LIB, 'GUI')
DIALOG = os.path.join(GUI, 'T3LabAssistantDialog.py')
PKG = os.path.join(GUI, 'assistant')
XAML = os.path.join(GUI, 'Tools', 'T3LabAssistant.xaml')
WPF_BASE = os.path.join(GUI, 'WPF_Base.py')

WINDOW = 'T3LabAssistantWindow'

# Module-level names GUI.T3LabAssistantDialog bound BEFORE the split (symtable
# of the 11,930-line file; `e` / `_theme_ex` were except-clause temporaries).
# Every one must stay importable from the dialog.
ORIGINAL_MODULE_NAMES = (
    'Action', 'AgentDispatcher', 'ApartmentState', 'BitmapImage',
    'ContextScout', 'DB', 'DROPPED_TOOLS', 'GraphOrchestrator', 'GridLength',
    'HAS_AGENTS', 'HAS_API_CONTEXT', 'HAS_DISCOVERY', 'HAS_EXECUTOR',
    'HAS_GRAPH', 'HAS_KNOWLEDGE', 'HAS_LANGUAGE', 'HAS_LINKS', 'HAS_NLP',
    'HAS_RAG', 'HAS_SCOUT', 'HAS_SPECIALISTS', 'HAS_TELEMETRY', 'HOST_APP',
    'LEGACY_DOC_KEY', 'MAX_ARCHIVED_SESSIONS', 'MAX_REPLAY_BUBBLES',
    'REVIT_VERSION', 'SUPPORTED_EXTS', 'StreamingJSONExtractor', 'System',
    'T3LabAssistantWindow', 'T3WPFWindow', 'TOOL_LAUNCHERS', 'TOOL_TITLES',
    'Thread', 'ThreadStart', 'Uri', 'UriKind', 'Visibility', 'WindowState',
    '_ANNOUNCE_FIXUP', '_CONFIRM_TIMEOUT', '_ICON_AMBER', '_ICON_ANALYZE',
    '_ICON_ATTACH', '_ICON_BLUE', '_ICON_EDIT', '_ICON_FAILED', '_ICON_GREEN',
    '_ICON_INFO', '_ICON_LIST', '_ICON_PENDING', '_ICON_RED', '_ICON_REFRESH',
    '_ICON_SEARCH', '_ICON_SLATE', '_ICON_STOP', '_ICON_SUCCESS', '_ICON_SYNC',
    '_ICON_WARNING', '_LAST_USER_VIET', '_LEGACY_CHAT_HISTORY_DIR',
    '_LIGHT_FALLBACK', '_MEM_SAVE_RX', '_PLAN_APPROVAL_TIMEOUT',
    '_PLAN_CANCEL', '_PLAN_EDIT', '_PLAN_RUN', '_PLAN_SUPERSEDED',
    '_RAG_SYSTEM_PREFIX', '_SHARED_PROVIDER_COLORS', '_SHARED_PROVIDER_GRAY',
    '_SPECIAL_LAUNCHERS', '_STEP_LOOK', '_THINK_BLOCK_RE', '_THINK_OPEN_RE',
    '_VIET_CHARS', '__author__', '__title__', '_active_budgets', '_active_doc',
    '_announces_work', '_bind_bg', '_bind_border', '_bind_fg', '_bind_stroke',
    '_bind_theme', '_budget_history', '_build_system_prompt', '_call_summary',
    '_chat_history_dir', '_cpython_bootstrap', '_cur', '_detect_viet',
    '_exc_text', '_fit_restored', '_get_doc_key', '_get_tool_script_dir',
    '_get_uiapp', '_greetings', '_gui_dir', '_hide_reasoning', '_history_file',
    '_history_path_for_read', '_is_stale_tool_call_blob', '_is_viet_text',
    '_json_safe', '_json_text', '_lang_analyzer', '_lang_detect', '_lib_dir',
    '_load_batchout_mod', '_load_script', '_note_user_language',
    '_open_file_shell', '_open_in_explorer', '_parent', '_project_scope_lines',
    '_provider_usable', '_prune_archived_sessions',
    '_register_discovered_launchers', '_relative_time', '_reply_language',
    '_tb', '_theme', '_theme_color', '_time_greeting', '_trgb', '_ui_viet',
    'build_specialist_prompt', 'build_text_context',
    'build_vision_content_blocks', 'clear_chat_history', 'clr',
    'configure_batchout_window', 'datetime', 'direct_export',
    'discover_new_tools', 'ensure_api_context', 'extension_dir',
    'find_learned_match', 'forms', 'get_active_provider_name',
    'get_active_store', 'get_global_store', 'get_local_model_name',
    'get_provider_display_label', 'get_registered_tools',
    'get_setup_guidance_message', 'get_spec', 'get_tool_title', 'has_api_key',
    'has_images', 'has_local_llm', 'inject_discovered_tools', 'io', 'is_image',
    'is_pdf', 'is_supported', 'json', 'keyword_parse', 'launch_batchout',
    'launch_batchout_configured', 'launch_loadfamily', 'learn_pattern',
    'lib_dir', 'link_reader', 'load_chat_history', 'load_chat_session',
    'load_chat_summary', 'logger', 'make_generic_launcher',
    'make_launcher_for', 'os', 'parse_command', 'parse_command_local',
    'parse_command_nlu', 're', 'revit', 'routing', 'run_in_api_context',
    'save_chat_history', 'script', 'show_assistant_dialog',
    'summarize_attachments', 'sys', 'telemetry',
)

# Rebound at runtime inside GUI.assistant.lang; the dialog serves it through a
# module __getattr__ (a `from ... import` copy would freeze).
LIVE_NAMES = ('_LAST_USER_VIET',)

# `self.<name>(` calls that did not resolve BEFORE the split either. Each is
# inside a try/except that logs and carries on; fix them separately, then
# delete the entry here.
KNOWN_UNRESOLVED = set()


# ── reading the code ──────────────────────────────────────────────────────────

def _read(path):
    with io.open(path, encoding='utf-8') as fh:
        return fh.read()


def _parse(path):
    return ast.parse(_read(path), filename=path)


def _class(tree, name):
    for node in tree.body:
        if isinstance(node, ast.ClassDef) and node.name == name:
            return node
    return None


def _methods(cls):
    return dict((n.name, n) for n in cls.body if isinstance(n, ast.FunctionDef))


def _class_attrs(cls):
    out = set()
    for node in cls.body:
        if isinstance(node, ast.Assign):
            for tg in node.targets:
                for n in ast.walk(tg):
                    if isinstance(n, ast.Name):
                        out.add(n.id)
    return out


def _is_static(fn):
    return any(isinstance(d, ast.Name) and d.id == 'staticmethod'
               for d in fn.decorator_list)


def _self_attrs(fn, root_name='self'):
    """[(node, attr_chain)] for `self.a`, `self.a.b` inside fn.

    attr_chain is ('a',) or ('a', 'b'). Skips nested functions / lambdas
    that take their own `self` parameter.
    """
    out = []

    def params(f):
        a = f.args
        return [x.arg for x in a.posonlyargs + a.args + a.kwonlyargs]

    def visit(node, shadowed):
        if isinstance(node, (ast.FunctionDef, ast.Lambda)) and node is not fn:
            inner = shadowed or (root_name in params(node))
            for child in (node.body if isinstance(node.body, list)
                          else [node.body]):
                visit(child, inner)
            return
        if (not shadowed and isinstance(node, ast.Attribute)
                and isinstance(node.value, ast.Name)
                and node.value.id == root_name):
            out.append((node, node.attr))
        for child in ast.iter_child_nodes(node):
            visit(child, shadowed)

    for stmt in fn.body:
        visit(stmt, False)
    return out


def _parents(tree):
    parents = {}
    for node in ast.walk(tree):
        for child in ast.iter_child_nodes(node):
            parents[child] = node
    return parents


def _event_names():
    """T3WPFWindow's XAML event attribute names (_EVENT_NAMES in WPF_Base)."""
    for node in _parse(WPF_BASE).body:
        if isinstance(node, ast.Assign) and any(
                isinstance(t, ast.Name) and t.id == '_EVENT_NAMES'
                for t in node.targets):
            return set(ast.literal_eval(node.value))
    raise AssertionError('WPF_Base._EVENT_NAMES not found')


def _xaml_handlers():
    alt = '|'.join(sorted(_event_names(), key=len, reverse=True))
    return set(re.findall(r'\b(?:' + alt + r')\s*=\s*"([^"]*)"', _read(XAML)))


def _xaml_element_names():
    return set(re.findall(r'\b(?:x:)?Name\s*=\s*"([^"]+)"', _read(XAML)))


def _base_members():
    """Methods + class attributes of every class in WPF_Base (T3WPFWindow)."""
    out = set()
    for node in _parse(WPF_BASE).body:
        if isinstance(node, ast.ClassDef):
            out |= set(_methods(node)) | _class_attrs(node)
            for fn in _methods(node).values():
                for n, attr in _self_attrs(fn):
                    out.add(attr)
    return out


def _helpers():
    """{attr: (module, class_name, ClassDef)} built in the window's __init__."""
    tree = _parse(DIALOG)
    window = _class(tree, WINDOW)
    init = _methods(window)['__init__']
    imports = {}
    for node in tree.body:
        if isinstance(node, ast.ImportFrom) and node.module and \
                node.module.startswith('GUI.assistant.'):
            for a in node.names:
                imports[a.asname or a.name] = node.module.split('.')[-1]
    out = {}
    for node in ast.walk(init):
        # self.<attr> = <Class imported from GUI.assistant.*>(self)
        if (isinstance(node, ast.Assign) and len(node.targets) == 1
                and isinstance(node.targets[0], ast.Attribute)
                and isinstance(node.targets[0].value, ast.Name)
                and node.targets[0].value.id == 'self'
                and isinstance(node.value, ast.Call)
                and isinstance(node.value.func, ast.Name)
                and node.value.func.id in imports
                and [ast.unparse(a) for a in node.value.args] == ['self']):
            cls_name = node.value.func.id
            module = imports[cls_name]
            cls = _class(_parse(os.path.join(PKG, module + '.py')), cls_name)
            out[node.targets[0].attr] = (module, cls_name, cls)
    return out


# ── a. XAML handlers ──────────────────────────────────────────────────────────

class XamlHandlersTest(unittest.TestCase):

    def test_every_xaml_handler_is_a_window_method(self):
        window = _class(_parse(DIALOG), WINDOW)
        have = set(_methods(window)) | _base_members()
        handlers = _xaml_handlers()
        self.assertTrue(len(handlers) >= 20, sorted(handlers))
        missing = sorted(h for h in handlers if h not in have)
        self.assertEqual(missing, [], 'XAML handlers with no method on the '
                                      'window: {}'.format(missing))

    def test_xaml_handlers_take_sender_and_args(self):
        """T3WPFWindow passes (sender, e); a stub must accept both."""
        window = _methods(_class(_parse(DIALOG), WINDOW))
        for name in _xaml_handlers():
            fn = window.get(name)
            if fn is None:
                continue        # inherited (reported by the test above)
            a = fn.args
            positional = len(a.posonlyargs + a.args)
            self.assertTrue(positional >= 3 or a.vararg is not None,
                            '{} cannot take (sender, e)'.format(name))

    def test_stubs_delegate_to_a_real_helper_method(self):
        """`return self.<helper>.<name>(...)` stubs point at existing methods."""
        helpers = _helpers()
        window = _methods(_class(_parse(DIALOG), WINDOW))
        stubs = 0
        for name, fn in window.items():
            body = [s for s in fn.body if not (
                isinstance(s, ast.Expr) and isinstance(s.value, ast.Constant))]
            if len(body) != 1 or not isinstance(body[0], ast.Return):
                continue
            call = body[0].value
            if not (isinstance(call, ast.Call)
                    and isinstance(call.func, ast.Attribute)
                    and isinstance(call.func.value, ast.Attribute)
                    and isinstance(call.func.value.value, ast.Name)
                    and call.func.value.value.id == 'self'
                    and call.func.value.attr in helpers):
                continue
            stubs += 1
            module, cls_name, cls = helpers[call.func.value.attr]
            self.assertEqual(call.func.attr, name,
                             'stub {} delegates to {}'.format(name, call.func.attr))
            self.assertIn(name, _methods(cls),
                          '{}.{} missing (stub on the window)'.format(cls_name, name))
        self.assertTrue(stubs > 0)


# ── b. names the dialog module exports ────────────────────────────────────────

def _module_bound_names(path):
    table = symtable.symtable(_read(path), path, 'exec')
    return set(s.get_name() for s in table.get_symbols()
               if s.is_assigned() or s.is_imported() or s.is_namespace())


def _getattr_serves(path):
    """Names a module-level __getattr__ answers (string compares in it)."""
    for node in _parse(path).body:
        if isinstance(node, ast.FunctionDef) and node.name == '__getattr__':
            return set(n.value for n in ast.walk(node)
                       if isinstance(n, ast.Constant) and isinstance(n.value, str))
    return set()


class DialogExportsTest(unittest.TestCase):

    def test_every_pre_split_name_is_still_bound(self):
        bound = _module_bound_names(DIALOG) | _getattr_serves(DIALOG)
        missing = sorted(n for n in ORIGINAL_MODULE_NAMES if n not in bound)
        self.assertEqual(missing, [], 'no longer importable from '
                                      'GUI.T3LabAssistantDialog: {}'.format(missing))

    def test_live_names_are_served_by_getattr(self):
        served = _getattr_serves(DIALOG)
        for name in LIVE_NAMES:
            self.assertIn(name, served)
            # and never frozen by a module-level binding in the dialog
            self.assertNotIn(name, _module_bound_names(DIALOG))

    def test_names_other_modules_import_exist(self):
        bound = _module_bound_names(DIALOG) | _getattr_serves(DIALOG)
        wanted = []
        for dirpath, dirnames, filenames in os.walk(ROOT):
            dirnames[:] = [d for d in dirnames if d not in ('.git', '__pycache__')]
            for fname in filenames:
                if not fname.endswith('.py'):
                    continue
                path = os.path.join(dirpath, fname)
                if os.path.abspath(path) == os.path.abspath(DIALOG):
                    continue
                try:
                    tree = _parse(path)
                except (SyntaxError, ValueError):
                    continue
                for node in ast.walk(tree):
                    if isinstance(node, ast.ImportFrom) and node.module in (
                            'GUI.T3LabAssistantDialog', 'T3LabAssistantDialog'):
                        for a in node.names:
                            wanted.append((path, a.name))
        self.assertTrue(wanted, 'expected at least the pushbutton import')
        missing = [(os.path.relpath(p, ROOT), n) for p, n in wanted
                   if n != '*' and n not in bound]
        self.assertEqual(missing, [])


# ── c. self.<name> resolves ───────────────────────────────────────────────────

class SelfReferencesTest(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        tree = _parse(DIALOG)
        cls.window = _class(tree, WINDOW)
        cls.window_methods = _methods(cls.window)
        cls.helpers = _helpers()
        stored = set()
        for fn in cls.window_methods.values():
            for node, attr in _self_attrs(fn):
                if isinstance(node.ctx, (ast.Store, ast.Del)):
                    stored.add(attr)
        cls.helper_methods = {}
        for attr, (module, cls_name, hcls) in cls.helpers.items():
            cls.helper_methods[attr] = _methods(hcls)
            for fn in _methods(hcls).values():
                parents = _parents(fn)
                for node, a in _self_attrs(fn):
                    p = parents.get(node)
                    if a == 'win' and isinstance(p, ast.Attribute) and \
                            isinstance(p.ctx, (ast.Store, ast.Del)):
                        stored.add(p.attr)
        cls.window_names = (set(cls.window_methods) | _class_attrs(cls.window)
                            | stored | _base_members() | _xaml_element_names()
                            | set(cls.helpers))

    def _window_has(self, name):
        # CLR members of System.Windows.Window (Close, Dispatcher, ...) are
        # PascalCase; nothing of ours is.
        return name in self.window_names or name[:1].isupper()

    def test_helpers_were_found(self):
        self.assertTrue(len(self.helpers) >= 5, sorted(self.helpers))

    def test_window_self_calls_resolve(self):
        bad = []
        for name, fn in self.window_methods.items():
            for node, attr in _self_attrs(fn):
                if not self._window_has(attr) and attr not in KNOWN_UNRESOLVED:
                    bad.append('{}: self.{}'.format(name, attr))
        self.assertEqual(sorted(set(bad)), [])

    def test_window_calls_into_helpers_resolve(self):
        bad = []
        for name, fn in self.window_methods.items():
            parents = _parents(fn)
            for node, attr in _self_attrs(fn):
                if attr not in self.helpers:
                    continue
                p = parents.get(node)
                if isinstance(p, ast.Attribute) and \
                        p.attr not in self.helper_methods[attr]:
                    bad.append('{}: self.{}.{}'.format(name, attr, p.attr))
        self.assertEqual(sorted(set(bad)), [])

    def test_helper_self_references_resolve(self):
        """Helpers keep NO state: `self.x` is a helper method or `self.win`;
        `self.win.x` is on the window; `self.win.<helper>.x` exists."""
        bad = []
        for attr, methods in self.helper_methods.items():
            for name, fn in methods.items():
                if name == '__init__':
                    continue
                parents = _parents(fn)
                for node, a in _self_attrs(fn):
                    where = '{}.{}'.format(attr, name)
                    if a != 'win':
                        if a not in methods:
                            bad.append('{}: self.{}'.format(where, a))
                        continue
                    p = parents.get(node)
                    if not isinstance(p, ast.Attribute):
                        continue            # bare self.win (passed along)
                    if not self._window_has(p.attr) and \
                            p.attr not in KNOWN_UNRESOLVED:
                        bad.append('{}: self.win.{}'.format(where, p.attr))
                        continue
                    if p.attr in self.helpers:
                        pp = parents.get(p)
                        if isinstance(pp, ast.Attribute) and \
                                pp.attr not in self.helper_methods[p.attr]:
                            bad.append('{}: self.win.{}.{}'.format(
                                where, p.attr, pp.attr))
        self.assertEqual(sorted(set(bad)), [])

    def test_no_method_lives_in_two_places(self):
        """A moved method is either gone from the window or a 1-line stub."""
        for attr, methods in self.helper_methods.items():
            for name in methods:
                if name == '__init__' or name not in self.window_methods:
                    continue
                fn = self.window_methods[name]
                body = [s for s in fn.body if not (
                    isinstance(s, ast.Expr) and isinstance(s.value, ast.Constant))]
                self.assertEqual(len(body), 1, '{} is both on the window and '
                                               'in {}'.format(name, attr))

    def test_helpers_never_name_the_window_class(self):
        for attr, (module, cls_name, hcls) in self.helpers.items():
            src = _read(os.path.join(PKG, module + '.py'))
            self.assertNotIn(WINDOW + '.', src, module)


# ── d. inheritance ────────────────────────────────────────────────────────────

class InheritanceTest(unittest.TestCase):

    def test_window_inherits_only_t3wpfwindow(self):
        window = _class(_parse(DIALOG), WINDOW)
        bases = [ast.unparse(b) for b in window.bases]
        self.assertEqual(bases, ['T3WPFWindow'])
        self.assertEqual(window.keywords, [])

    def test_no_class_in_the_dialog_mixes_a_clr_base(self):
        for node in ast.walk(_parse(DIALOG)):
            if isinstance(node, ast.ClassDef) and any(
                    ast.unparse(b) == 'T3WPFWindow' for b in node.bases):
                self.assertEqual(len(node.bases), 1, node.name)

    def test_helper_classes_inherit_only_object(self):
        for fname in sorted(os.listdir(PKG)):
            if not fname.endswith('.py'):
                continue
            for node in ast.walk(_parse(os.path.join(PKG, fname))):
                if isinstance(node, ast.ClassDef):
                    bases = [ast.unparse(b) for b in node.bases]
                    self.assertIn(bases, ([], ['object']),
                                  '{}: class {}({})'.format(fname, node.name,
                                                            ', '.join(bases)))

    def test_helpers_take_the_window_and_keep_it_as_win(self):
        for attr, (module, cls_name, hcls) in _helpers().items():
            init = _methods(hcls)['__init__']
            self.assertEqual([a.arg for a in init.args.args], ['self', 'window'])
            self.assertEqual(ast.unparse(init.body[0]), 'self.win = window',
                             cls_name)

    def test_helpers_are_built_before_the_xaml_loads(self):
        """XAML handler stubs reach the helpers, so they exist first."""
        init = _methods(_class(_parse(DIALOG), WINDOW))['__init__']
        helper_lines = [n.lineno for n in ast.walk(init)
                        if isinstance(n, ast.Assign)
                        and isinstance(n.value, ast.Call)
                        and isinstance(n.value.func, ast.Name)
                        and n.value.func.id.startswith('Assistant')]
        xaml_line = min(n.lineno for n in ast.walk(init)
                        if isinstance(n, ast.Call)
                        and ast.unparse(n.func) == 'T3WPFWindow.__init__')
        self.assertTrue(helper_lines)
        self.assertTrue(max(helper_lines) < xaml_line)


# ── e. compile + import under stubs ───────────────────────────────────────────

def _package_modules():
    return sorted(f[:-3] for f in os.listdir(PKG)
                  if f.endswith('.py') and f != '__init__.py')


class _Anything(object):
    """Stands in for any .NET / pyRevit object: every attribute, call,
    subclass and iteration works and yields another stand-in."""

    def __init__(self, *a, **kw):
        pass

    def __getattr__(self, name):
        if name.startswith('__') and name.endswith('__'):
            raise AttributeError(name)
        return _Anything()

    def __call__(self, *a, **kw):
        return _Anything()

    def __iter__(self):
        return iter(())

    def __bool__(self):
        return False

    def __int__(self):
        raise ValueError('stub')


class _StubModule(types.ModuleType):
    def __getattr__(self, name):
        if name.startswith('__') and name.endswith('__'):
            raise AttributeError(name)
        # PascalCase: a .NET type (may be subclassed); else a value. Cached,
        # so two `from System import Action` get the same object.
        value = type(str(name), (_Anything,), {}) if name[:1].isupper() \
            else _Anything()
        setattr(self, name, value)
        return value


def _stub_modules():
    import logging
    stubs = {}
    for name in ('clr', 'System', 'System.Windows', 'System.Windows.Media',
                 'System.Windows.Media.Imaging', 'System.Windows.Controls',
                 'System.Windows.Threading', 'System.Windows.Input',
                 'System.Threading', 'System.IO', 'Autodesk', 'Autodesk.Revit',
                 'Autodesk.Revit.DB', 'Autodesk.Revit.UI', 'Python',
                 'Python.Runtime'):
        stubs[name] = _StubModule(name)
    pyrevit = _StubModule('pyrevit')
    script = _StubModule('pyrevit.script')
    script.get_logger = lambda: logging.getLogger('T3LabAssistant.test')
    pyrevit.script = script
    pyrevit.revit = _StubModule('pyrevit.revit')
    pyrevit.forms = _StubModule('pyrevit.forms')
    stubs.update({'pyrevit': pyrevit, 'pyrevit.script': script,
                  'pyrevit.revit': pyrevit.revit, 'pyrevit.forms': pyrevit.forms})
    # The real bootstrap records a ribbon click; never run it from a test.
    boot = types.ModuleType('_cpython_bootstrap')
    boot.init_cpython_paths = lambda *a, **kw: None
    stubs['_cpython_bootstrap'] = boot
    # The real base needs WPF; the dialog only subclasses it.
    base = types.ModuleType('GUI.WPF_Base')

    class T3WPFWindow(object):
        def __init__(self, *a, **kw):
            pass
    base.T3WPFWindow = T3WPFWindow
    stubs['GUI.WPF_Base'] = base
    return stubs


class CompileAndImportTest(unittest.TestCase):

    def test_every_module_compiles(self):
        tmp = tempfile.mkdtemp(prefix='t3_assistant_compile_')
        paths = [DIALOG] + [os.path.join(PKG, f) for f in sorted(os.listdir(PKG))
                            if f.endswith('.py')]
        for path in paths:
            py_compile.compile(path, cfile=os.path.join(
                tmp, os.path.basename(path) + 'c'), doraise=True)

    def test_new_modules_keep_the_header_style(self):
        for name in ['__init__'] + _package_modules():
            src = _read(os.path.join(PKG, name + '.py'))
            self.assertTrue(src.startswith('# -*- coding: utf-8 -*-\n"""'), name)
            self.assertIn('__author__  = "Tran Tien Thanh"', src, name)
            fstrings = [n.lineno for n in ast.walk(ast.parse(src))
                        if isinstance(n, ast.JoinedStr)]
            self.assertEqual(fstrings, [], '{}: no f-strings (py2-style '
                                           'neighbours)'.format(name))

    def _import_all(self):
        saved_mods = dict(sys.modules)
        saved_path = list(sys.path)
        sys.modules.update(_stub_modules())
        if LIB not in sys.path:
            sys.path.insert(0, LIB)
        for name in list(sys.modules):
            if name == 'GUI.T3LabAssistantDialog' or name.startswith('GUI.assistant'):
                del sys.modules[name]

        def restore():
            sys.modules.clear()
            sys.modules.update(saved_mods)
            sys.path[:] = saved_path
        self.addCleanup(restore)
        mods = dict((m, importlib.import_module('GUI.assistant.' + m))
                    for m in _package_modules())
        dialog = importlib.import_module('GUI.T3LabAssistantDialog')
        return mods, dialog

    def test_modules_and_dialog_import_with_stubs(self):
        mods, dialog = self._import_all()
        self.assertTrue(hasattr(dialog, WINDOW))
        self.assertTrue(callable(dialog.show_assistant_dialog))
        for attr, (module, cls_name, hcls) in _helpers().items():
            helper_cls = getattr(mods[module], cls_name)
            self.assertIs(getattr(dialog, cls_name), helper_cls)

    def test_reexports_are_the_same_objects(self):
        """Every `from GUI.assistant.X import name` in the dialog binds the
        very object X holds (no copies that could drift)."""
        mods, dialog = self._import_all()
        checked = 0
        for node in _parse(DIALOG).body:
            stmts = node.body if isinstance(node, ast.Try) else [node]
            for st in stmts:
                if not (isinstance(st, ast.ImportFrom) and st.module
                        and st.module.startswith('GUI.assistant.')):
                    continue
                mod = mods[st.module.split('.')[-1]]
                for a in st.names:
                    bound = a.asname or a.name
                    if not hasattr(mod, a.name):
                        continue        # optional import failed, as before
                    self.assertIs(dialog.__dict__[bound], getattr(mod, a.name),
                                  '{}.{}'.format(st.module, a.name))
                    checked += 1
        self.assertTrue(checked > 140, checked)

    def test_without_the_nlp_layer_keyword_parse_still_exists(self):
        """A broken install (Intelligence.t3lab_assistant fails to import)
        must still answer the offline path with guidance, not a NameError."""
        broken = types.ModuleType('Intelligence.t3lab_assistant')
        saved = sys.modules.get('Intelligence.t3lab_assistant')
        sys.modules['Intelligence.t3lab_assistant'] = broken   # no names in it
        try:
            mods, dialog = self._import_all()
        finally:
            if saved is None:
                sys.modules.pop('Intelligence.t3lab_assistant', None)
            else:
                sys.modules['Intelligence.t3lab_assistant'] = saved
        deps = mods['deps']
        self.assertFalse(deps.HAS_NLP)
        for name in ('keyword_parse', 'parse_command', 'has_api_key'):
            self.assertIs(getattr(dialog, name), getattr(deps, name), name)
        self.assertIsNone(dialog.keyword_parse(u'mở batchout'))
        self.assertFalse(dialog.has_api_key())
        self.assertTrue(dialog.get_setup_guidance_message(False))

    def test_shared_registries_are_one_object(self):
        mods, dialog = self._import_all()
        launchers = mods['launchers']
        for name in ('TOOL_LAUNCHERS', 'TOOL_TITLES', 'DROPPED_TOOLS'):
            self.assertIs(getattr(dialog, name), getattr(launchers, name))
        launchers._register_discovered_launchers(
            [{'intent': 'open_t3_test', 'script_path': None}])
        self.assertIn(('open_t3_test', u'<no path>'), dialog.DROPPED_TOOLS)

    def test_last_user_language_is_live_through_the_dialog(self):
        mods, dialog = self._import_all()
        lang = mods['lang']
        lang._LAST_USER_VIET = None
        self.assertIsNone(dialog._LAST_USER_VIET)
        lang._LAST_USER_VIET = True
        self.assertIs(dialog._LAST_USER_VIET, True)
        from GUI.T3LabAssistantDialog import _LAST_USER_VIET
        self.assertIs(_LAST_USER_VIET, True)
        lang._LAST_USER_VIET = None
        with self.assertRaises(AttributeError):
            dialog.no_such_name_t3

    def test_a_helper_reads_and_writes_state_on_the_window(self):
        """`self.win` is the window: the helper keeps no state of its own."""
        mods, dialog = self._import_all()
        win = types.SimpleNamespace(_session_id='sid', _session_notes=['n'],
                                    _session_path=None)
        helper = mods['sessions'].AssistantSessions(win)
        self.assertEqual(helper._session_state(),
                         {'session_id': 'sid', 'session_notes': ['n'],
                          'session_path': None})
        self.assertNotIn('_session_id', vars(helper))
        plans = mods['plans'].AssistantPlans(win)
        self.assertTrue(plans._plan_needs_approval(object()) in (True, False))


if __name__ == '__main__':
    unittest.main(verbosity=2)
