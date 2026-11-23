# -*- coding: utf-8 -*-
"""Owner side of the BIM Harness in Revit (lib/Services/harness_settings.py):
board rows and the summary sentence of MCP Control > Harness, review as the
person at the machine through the real bridge CLI, Harness Studio launch, and
the MCP Control wiring. Linux, no Revit. docs/harness.md.
From t3lab_dev's dev/test_mcp_agent_settings.py test_harness."""
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'lib')
sys.path.insert(0, LIB)
TMP = tempfile.mkdtemp(prefix='t3lab_harness_settings_')
os.environ['APPDATA'] = TMP

from Services import harness_settings as M  # noqa: E402

FAILURES = []


def check(name, cond, detail=''):
    print('  {}  {}{}'.format('ok  ' if cond else 'FAIL', name, '' if cond else '  ' + str(detail)))
    if not cond:
        FAILURES.append(name)


def test_harness():
    print('[harness board]')
    folder = os.path.join(TMP, 'harness_dir')
    os.environ['T3LAB_HARNESS_DIR'] = folder
    check('no board: nothing listed', M.list_harness_sessions() == [] and M.harness_rows() == [])
    check('no board: sentence', M.harness_summary() == 'No harness session is open.')
    # A real board, written by the real store + bridge.
    bridge = os.path.join(LIB, 'core', 'bridge.py')
    env = dict(os.environ, PYTHONPATH='')

    def cli(*args):
        done = subprocess.run([sys.executable, bridge, 'harness'] + list(args),
                              capture_output=True, text=True, env=env)
        return json.loads(done.stdout)

    sid = cli('open', '--json', json.dumps({'goal': 'Audit doors',
                                             'revit_document': 'PRJ.rvt'}))['session']
    cli('plan', '--json', json.dumps({'tasks': [
        {'key': 'a', 'title': 'List doors', 'assignee': 'cli', 'scope': {'write': False}},
        {'key': 'b', 'title': 'Fix marks', 'scope': {'document': 'B.rvt'}}]}))
    cli('next')
    cli('report', '--json', json.dumps({'task': 'a', 'summary': '40 doors'}))
    rows = M.harness_rows()
    check('rows: review first', [(r['task'], r['status_text'], r['owner'], r['document'])
                                 for r in rows] ==
          [('a', 'Waiting for review', 'cli', 'PRJ.rvt'), ('b', 'To do', '', 'B.rvt')], rows)
    check('rows carry the session', rows[0]['session'] == sid)
    check('summary sentence', M.harness_summary() ==
          'Audit doors (1 agent): 1 waiting for review, 0 of 2 done', M.harness_summary())
    code, out = M.harness_review(sid, 'a', 'redo', note='also level 2',
                                 python=sys.executable, bridge=bridge)
    check('send back as the owner', code == 0 and out.get('status') == 'todo', (code, out))
    check('sentence: sent back', M.describe_harness_result('redo', code, out, rows[0]) ==
          'Sent "List doors" back. A worker takes it again.')
    cli('next')
    cli('report', '--json', json.dumps({'task': 'a', 'summary': '52 doors'}))
    code, out = M.harness_review(sid, 'a', 'accept', python=sys.executable, bridge=bridge)
    check('accept as the owner', code == 0 and out.get('status') == 'done', (code, out))
    check('sentence: accepted', 'Accepted "List doors"' in
          M.describe_harness_result('accept', code, out, rows[0]))
    code, out = M.harness_review(sid, 'a', 'accept', python=sys.executable, bridge=bridge)
    check('second accept refused', code == 1 and out.get('error_code') == 'not_in_review',
          (code, out))
    check('sentence: no longer in review', 'no longer waiting' in
          M.describe_harness_result('accept', code, out, rows[0]))
    check('sentence: bridge did not answer', 'did not answer' in
          M.describe_harness_result('accept', None, 'The bridge is still running.'))
    check('board shows the result', [r['status'] for r in M.harness_rows()] == ['todo', 'done'])
    check('no master set up: not named', M.harness_master() == '')
    done = subprocess.run([sys.executable, bridge, 'harness', 'setup', '--master=cli',
                           '--write-mode=serial'], capture_output=True, text=True, env=env)
    check('setup from the CLI', json.loads(done.stdout)['setup']['master'] == 'cli', done.stdout)
    check('master read from the board', M.harness_master() == 'cli')
    check('summary names the master', M.harness_summary().startswith('Master: cli. Audit doors'),
          M.harness_summary())
    cli('close')
    serial = cli('open', '--json', json.dumps({'goal': 'Second'}))
    check('new session takes the setup', serial.get('write_mode') == 'serial', serial)
    cli('plan', '--json', json.dumps({'tasks': [{'key': 'x', 'title': 'x'}]}))
    check('summary says one writer at a time',
          M.harness_summary() == 'Master: cli. Second (1 agent): 0 of 1 done, one writer at a time',
          M.harness_summary())
    check('no session: master still shown',
          M.harness_summary([], 'cli') == 'Master: cli. No harness session is open.')
    del os.environ['T3LAB_HARNESS_DIR']


def test_launchers():
    print('[launchers]')
    seen = []
    M.open_studio('py', 'bridge.py', popen=lambda argv, **kw: seen.append((argv, kw)))
    argv, kw = seen[0]
    check('studio started detached', argv[:4] == ['py', 'bridge.py', 'harness', 'studio'], argv)
    check('studio exits when idle', argv[4] == '--idle-exit={}'.format(M.STUDIO_IDLE_EXIT_MIN))
    os.environ['T3LAB_AGENT'] = 'claude-desktop'
    env = M._child_env()
    del os.environ['T3LAB_AGENT']
    check('the bridge child is the person, not an AI app', 'T3LAB_AGENT' not in env
          and env.get('PYTHONIOENCODING') == 'utf-8')
    args = M.review_args('s1', 'a', 'accept', 'ok')
    check('review = owner from the CLI', args[:3] == ['harness', 'review', '--session=s1']
          and json.loads(args[4]) == {'task': 'a', 'verdict': 'accept', 'owner': True,
                                      'note': 'ok'}, args)


def test_mcp_control_wiring():
    print('[MCP Control > Harness]')
    xaml_path = os.path.join(LIB, 'GUI', 'Tools', 'MCPControl.xaml')
    with open(xaml_path, encoding='utf-8') as f:
        xaml = f.read()
    ET.fromstring(xaml.encode('utf-8'))
    names = set(re.findall(r'x:Name="([^"]+)"', xaml))
    with open(os.path.join(LIB, 'GUI', 'MCPControlDialog.py'), encoding='utf-8') as f:
        dialog = f.read()
    used = set(re.findall(r"FindName\('(harness_[a-z_]+)'\)", dialog))
    check('every Harness control the dialog finds is in the XAML', used and used <= names,
          sorted(used - names))
    check('the tab chip selects the third tab', 'x:Name="chip_tab_harness"' in xaml
          and 'Tag="2"' in xaml.split('x:Name="chip_tab_harness"')[1][:300])
    check('tab 2 refreshes the board', 'SelectedIndex == 2' in dialog
          and 'self._refresh_harness()' in dialog)
    check('one empty state on the grid', 'x:Name="harness_empty" Style="{StaticResource T3.Empty}"'
          in xaml)
    check('no hex colours in the tab', not re.search(
        r'#[0-9A-Fa-f]{6}', xaml.split('TAB 2: HARNESS')[1].split('</TabItem>')[0]))
    for key in set(re.findall(r'StaticResource (T3\.[A-Za-z.]+)\}',
                              xaml.split('TAB 2: HARNESS')[1].split('</TabItem>')[0])):
        check('style key exists: ' + key, 'x:Key="{}"'.format(key) in xaml)


def main():
    try:
        test_harness()
        test_launchers()
        test_mcp_control_wiring()
    finally:
        shutil.rmtree(TMP, ignore_errors=True)
    print('')
    if FAILURES:
        print('FAILED: {} check(s): {}'.format(len(FAILURES), ', '.join(FAILURES)))
        return 1
    print('ALL PASS')
    return 0


if __name__ == '__main__':
    sys.exit(main())
