#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Project memory shared by the T3Lab Assistant, Claude Code and MCP apps --
no Revit needed.

    python3 lite_guard/test_memory_hub.py

Covers lib/Intelligence/memory_hub.py and its two front doors:
  - Claude Code memory of the folders linked to the active project
    (CLAUDE.md, .claude/CLAUDE.md, ~/.claude/projects/<folder>/memory/*.md),
    its on/off setting, search and read across both sources
  - the Assistant prompt: every fact while small, an index + recall_memory
    once long or when Claude Code notes exist; unchanged without tools
  - recall_memory in the Assistant's tool list (never deferred, local)
  - t3lab_memory / t3lab_memory_save on the MCP server: answered without
    Revit, read-only mode, the documentation toolset, the memory resource

Plain CPython 3, standard library only.
"""
import io
import json
import os
import shutil
import sys
import tempfile
import unittest

sys.dont_write_bytecode = True

_TMP = tempfile.mkdtemp(prefix='t3lab_memhub_test_')
os.environ['APPDATA'] = _TMP
os.environ['HOME'] = _TMP
os.environ['USERPROFILE'] = _TMP
os.environ['CLAUDE_CONFIG_DIR'] = os.path.join(_TMP, 'claude-config')

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(ROOT, 'lib')
if LIB not in sys.path:
    sys.path.insert(0, LIB)

from core import paths                          # noqa: E402
from core import mcp_policy as mp               # noqa: E402
from core import server as srv_mod              # noqa: E402
from config.project_store import ProjectStore   # noqa: E402
from config.settings import get_settings        # noqa: E402
from Intelligence import assistant_memory as am  # noqa: E402
from Intelligence import memory_hub as mh       # noqa: E402
from Intelligence import tool_schema            # noqa: E402

# Tool calls below go through the usage tracking hook: never send rows.
paths.save_settings({'tracking_enabled': False})

_SERVER = None


def _server():
    global _SERVER
    if _SERVER is None:
        _SERVER = srv_mod.T3LabAIServer()
    return _SERVER


def _write(path, text):
    folder = os.path.dirname(path)
    if not os.path.isdir(folder):
        os.makedirs(folder)
    with io.open(path, 'w', encoding='utf-8') as f:
        f.write(text)


class _ProjectCase(unittest.TestCase):
    """A project with one linked folder that Claude Code also knows."""

    def setUp(self):
        self.ps = ProjectStore()
        meta = self.ps.create_project(u'Tower A')
        self.pid = meta['id'] if isinstance(meta, dict) else meta
        self.ps.set_active_project(self.pid)
        self.folder = tempfile.mkdtemp(prefix='TowerA ', dir=_TMP)
        self.ps.add_knowledge_dir(self.pid, self.folder)
        _write(os.path.join(self.folder, 'CLAUDE.md'),
               u'# Tower A\n\nSheet numbering: A-1xx plans, A-2xx elevations.\n')
        enc = mh.encode_project_dir(self.folder)
        self.auto = os.path.join(os.environ['CLAUDE_CONFIG_DIR'], 'projects',
                                 enc, 'memory')
        _write(os.path.join(self.auto, 'MEMORY.md'),
               u'# Levels\n\nLevel names are L01, L02 ... roof is RF.\n\n'
               u'# Client\n\nThe client wants mm on every dimension.\n')
        get_settings().set_agent_option(mh.SETTING_KEY, True)
        am.clear_facts(everything=True)
        mh.clear_cache()

    def tearDown(self):
        am.clear_facts(everything=True)
        self.ps.set_active_project(None)
        shutil.rmtree(self.folder, ignore_errors=True)
        shutil.rmtree(os.path.join(os.environ['CLAUDE_CONFIG_DIR'], 'projects'),
                      ignore_errors=True)


class ClaudeCodeSourceTests(_ProjectCase):

    def test_encoding_matches_claude_code(self):
        self.assertEqual(mh.encode_project_dir('/home/u/T3Lab Lite'),
                         '-home-u-T3Lab-Lite')
        self.assertNotRegex(mh.encode_project_dir(self.folder), r'[^A-Za-z0-9-]')

    def test_finds_claude_md_and_auto_memory(self):
        sources = mh.claude_sources(self.pid)
        kinds = sorted(s['kind'] for s in sources)
        self.assertEqual(kinds, ['CLAUDE.md', 'auto memory'])
        self.assertEqual(mh.stats(self.pid)['claude_files'], 2)

    def test_source_list_is_cached(self):
        self.assertEqual(len(mh.claude_sources(self.pid)), 2)
        _write(os.path.join(self.auto, 'extra.md'), u'# Extra\n\nmore notes\n')
        self.assertEqual(len(mh.claude_sources(self.pid)), 2)    # within TTL
        mh.clear_cache()
        self.assertEqual(len(mh.claude_sources(self.pid)), 3)

    def test_switch_off(self):
        get_settings().set_agent_option(mh.SETTING_KEY, False)
        self.assertEqual(mh.claude_sources(self.pid), [])
        get_settings().set_agent_option(mh.SETTING_KEY, True)

    def test_no_linked_folder_means_no_claude_files(self):
        self.assertEqual(mh.claude_sources(None), [])


class SearchReadTests(_ProjectCase):

    def test_search_both_sources(self):
        am.add_fact(u'Sheet prefix for this job is WH-', 'project', self.pid)
        out = mh.search(u'sheet numbering', self.pid)
        sources = [r['source'] for r in out['results']]
        self.assertIn('t3lab', sources)
        self.assertIn('claude', sources)
        out = mh.search(u'level names roof', self.pid)
        self.assertEqual(out['results'][0]['source'], 'claude')
        self.assertIn('RF', out['results'][0]['text'])
        self.assertEqual(mh.search(u'zzqx', self.pid)['count'], 0)

    def test_vietnamese_query(self):
        am.add_fact(u'Tên tầng dùng dạng Tầng 01, Tầng 02', 'project', self.pid)
        out = mh.search(u'ten tang', self.pid)
        self.assertEqual(out['results'][0]['id'], 'fact:1')

    def test_read(self):
        am.add_fact(u'Answer in Vietnamese', 'global', self.pid)
        self.assertEqual(mh.read('fact:1', self.pid)['scope'], 'all projects')
        whole = mh.read('claude:2', self.pid)
        self.assertIn('client wants mm', whole['text'])
        self.assertIn('error', mh.read('claude:9', self.pid))
        self.assertIn('error', mh.read('nonsense', self.pid))

    def test_call_read_defaults(self):
        self.assertIn('facts', mh.call_read({}, self.pid))                 # list
        self.assertIn('results', mh.call_read({'query': 'levels'}, self.pid))
        self.assertIn('error', mh.call_read({'action': 'read'}, self.pid))
        self.assertIn('error', mh.call_read({'action': 'drop'}, self.pid))

    def test_save_update_forget(self):
        ok = mh.call_save({'fact': u'Dimensions in mm'}, self.pid)
        self.assertTrue(ok['success'])
        self.assertEqual(am.list_facts(self.pid)[0][0], 'project')
        mh.call_save({'action': 'update', 'replaces': 'fact:1',
                      'fact': u'Dimensions in mm, no decimals'}, self.pid)
        self.assertIn('no decimals', am.list_facts(self.pid)[0][1]['text'])
        out = mh.call_save({'action': 'forget', 'replaces': 'fact:1',
                            'fact': 'x'}, self.pid)
        self.assertTrue(out['success'])
        self.assertEqual(am.list_facts(self.pid), [])
        self.assertIn('error', mh.call_save({'action': 'forget',
                                             'fact': u'nothing like it'}, self.pid))


class PromptBlockTests(_ProjectCase):

    def test_without_tools_unchanged(self):
        am.add_fact(u'Dimensions in mm', 'project', self.pid)
        self.assertEqual(mh.prompt_block(self.pid, recall_tool=False),
                         am.build_memory_block(self.pid, include_tool_hint=False))

    def test_small_memory_without_claude_files_is_inlined(self):
        get_settings().set_agent_option(mh.SETTING_KEY, False)
        am.add_fact(u'Dimensions in mm', 'project', self.pid)
        block = mh.prompt_block(self.pid, recall_tool=True)
        self.assertEqual(block, am.build_memory_block(self.pid, include_tool_hint=True))
        self.assertNotIn(mh.RECALL_TOOL_NAME, block)
        get_settings().set_agent_option(mh.SETTING_KEY, True)

    def test_claude_files_give_an_index(self):
        am.add_fact(u'Dimensions in mm', 'project', self.pid)
        block = mh.prompt_block(self.pid, recall_tool=True)
        self.assertIn('Dimensions in mm', block)
        self.assertIn('2 Claude Code memory file(s)', block)
        self.assertIn('`recall_memory`', block)
        self.assertIn('remember_fact', block)                 # tool hint kept
        self.assertNotIn('A-1xx', block)                      # content NOT inlined

    def test_long_memory_shows_newest_and_counts_the_rest(self):
        get_settings().set_agent_option(mh.SETTING_KEY, False)
        for i in range(30):
            am.add_fact(u'Convention number {} for this project is fixed'.format(i),
                        'project', self.pid)
        full = am.build_memory_block(self.pid, include_tool_hint=True)
        block = mh.prompt_block(self.pid, recall_tool=True)
        self.assertLess(len(block), len(full) * 0.7)
        self.assertIn('Convention number 29', block)          # newest kept
        self.assertNotIn('Convention number 0 ', block)
        self.assertIn('more saved fact(s)', block)
        get_settings().set_agent_option(mh.SETTING_KEY, True)

    def test_no_memory_no_block(self):
        get_settings().set_agent_option(mh.SETTING_KEY, False)
        self.assertEqual(mh.prompt_block(self.pid, recall_tool=True), u'')
        get_settings().set_agent_option(mh.SETTING_KEY, True)


class AssistantToolTests(unittest.TestCase):

    def test_recall_tool_is_local_and_never_deferred(self):
        tool = tool_schema.make_recall_tool()
        self.assertEqual(tool['name'], tool_schema.RECALL_TOOL_NAME)
        self.assertIn(tool_schema.RECALL_TOOL_NAME, tool_schema._EXTRA_TOOL_NAMES)
        from Intelligence import agent_loop
        self.assertIn(tool_schema.RECALL_TOOL_NAME, agent_loop._LOCAL_PSEUDO_TOOLS)
        deferred = tool_schema.with_deferred_loading(
            [dict(tool), {'name': 'list_levels', 'description': 'x',
                          'input_schema': {'type': 'object'}}], keep=frozenset())
        mine = [t for t in deferred if t.get('name') == tool['name']][0]
        self.assertFalse(mine.get('defer_loading'))

    def test_mcp_memory_tools_stay_out_of_the_assistant_catalog(self):
        self.assertIn('t3lab_memory', tool_schema._EXTERNAL_ONLY_TOOLS)
        self.assertIn('t3lab_memory_save', tool_schema._EXTERNAL_ONLY_TOOLS)
        self.assertIn('t3lab_memory', tool_schema.READ_ONLY_TOOL_NAMES)
        self.assertNotIn('t3lab_memory_save', tool_schema.READ_ONLY_TOOL_NAMES)

    def test_dialog_wires_the_tool(self):
        with io.open(os.path.join(LIB, 'GUI', 'T3LabAssistantDialog.py'),
                     encoding='utf-8') as f:
            src = f.read()
        self.assertIn('_extras.append(tool_schema.make_recall_tool())', src)
        self.assertIn('if name == tool_schema.RECALL_TOOL_NAME:', src)


class McpTests(_ProjectCase):

    def setUp(self):
        super(McpTests, self).setUp()
        self.s = _server()
        self.s._external_event = None          # no Revit: must still answer

    def tearDown(self):
        data = paths.load_settings() or {}
        for key in (mp.READ_ONLY_KEY, mp.TOOLSET_KEY):
            data.pop(key, None)
        paths.save_settings(data)
        super(McpTests, self).tearDown()

    def _call(self, name, arguments):
        res = self.s._handle_tool_call({'name': name, 'arguments': arguments})
        return res, json.loads(res['content'][0]['text'])

    def _listed(self):
        return {t['name'] for t in self.s._handle_tools_list()['tools']}

    def test_listed_and_answered_without_revit(self):
        self.assertTrue({'t3lab_memory', 't3lab_memory_save'} <= self._listed())
        res, out = self._call('t3lab_memory_save', {'fact': u'Roof level is RF'})
        self.assertFalse(res['isError'])
        res, out = self._call('t3lab_memory', {'query': 'roof level'})
        self.assertFalse(res['isError'])
        self.assertEqual(out['results'][0]['source'], 't3lab')
        # The Assistant sees what the AI app saved.
        self.assertIn('Roof level is RF', mh.prompt_block(self.pid, recall_tool=True))

    def test_read_only_mode_keeps_search_but_not_save(self):
        mp.save_policy(read_only=True)
        listed = self._listed()
        self.assertIn('t3lab_memory', listed)
        self.assertNotIn('t3lab_memory_save', listed)
        res, _out = self._call('t3lab_memory_save', {'fact': u'x y z w'})
        self.assertTrue(res['isError'])

    def test_documentation_toolset_keeps_both(self):
        mp.save_policy(toolset='documentation')
        self.assertTrue({'t3lab_memory', 't3lab_memory_save'} <= self._listed())

    def test_resource(self):
        am.add_fact(u'Dimensions in mm', 'project', self.pid)
        self.assertEqual(mh.resource_entry()['uri'], mh.RESOURCE_URI)
        out = srv_mod._memory_read_resource(mh.RESOURCE_URI)
        text = out['contents'][0]['text']
        self.assertIn('Dimensions in mm', text)
        self.assertIn('Claude Code memory files', text)
        self.assertIsNone(srv_mod._memory_read_resource('t3lab://memory/other'))

    def test_instructions_name_the_tools(self):
        from core import bridge
        self.assertIn('t3lab_memory', mp.SERVER_INSTRUCTIONS)
        self.assertIn('t3lab_memory_save', mp.SERVER_INSTRUCTIONS)
        self.assertEqual(bridge.SERVER_INSTRUCTIONS, mp.SERVER_INSTRUCTIONS)


def tearDownModule():
    shutil.rmtree(_TMP, ignore_errors=True)


if __name__ == '__main__':
    unittest.main(verbosity=1)
