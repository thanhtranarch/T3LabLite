"""BatchOut PDF engine "PDF24 printer" (Services/batchout_pdf24.py + its wiring)."""
import os
import re
import sys
import time
import unittest
from tempfile import TemporaryDirectory

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(ROOT, 'lib')
sys.path.insert(0, LIB)

from Services import batchout_pdf24 as pdf24  # noqa: E402


class FakeRegistry(object):
    def __init__(self, hkcu=None, hklm=None, devices=None):
        self.data = {
            (pdf24.HKCU, pdf24.SERVICE_KEY): dict(hkcu or {}),
            (pdf24.HKLM, pdf24.SERVICE_KEY): dict(hklm or {}),
            (pdf24.HKCU, pdf24.DEVICES_KEY): dict((n, ('winspool,Ne01:', 1)) for n in devices or []),
        }

    def read(self, hive, key, name):
        return self.data.get((hive, key), {}).get(name)

    def names(self, hive, key):
        return list(self.data.get((hive, key), {}))

    def write(self, hive, key, name, value, kind):
        self.data.setdefault((hive, key), {})[name] = (value, kind)

    def delete(self, hive, key, name):
        self.data.get((hive, key), {}).pop(name, None)

    def hkcu(self):
        return self.data[(pdf24.HKCU, pdf24.SERVICE_KEY)]


class PrinterTests(unittest.TestCase):
    def test_pdf24_itself_wins(self):
        self.assertEqual(pdf24.pick_printer(['Microsoft Print to PDF', 'PDF24 Fax', 'PDF24']), 'PDF24')

    def test_renamed_printer_but_never_fax(self):
        self.assertEqual(pdf24.pick_printer(['PDF24 Fax', 'PDF24 PDF']), 'PDF24 PDF')
        self.assertIsNone(pdf24.pick_printer(['PDF24 Fax', 'Adobe PDF']))
        self.assertIsNone(pdf24.pick_printer([]))

    def test_found_from_the_user_devices_key(self):
        reg = FakeRegistry(devices=['OneNote', 'PDF24'])
        self.assertEqual(pdf24.find_printer(reg), 'PDF24')


class NameTests(unittest.TestCase):
    def test_percent_is_doubled_and_extension_dropped(self):
        self.assertEqual(pdf24.auto_save_filename('A101 - 100% Plan.pdf'), 'A101 - 100%% Plan')
        self.assertEqual(pdf24.auto_save_filename('  A102 '), 'A102')

    def test_paper_matching(self):
        papers = ['A0', 'A1', 'A10', 'A3', 'A4 (210 x 297 mm)', 'ANSI D', 'Letter']
        self.assertEqual(pdf24.match_paper('A1', papers), 'A1')
        self.assertEqual(pdf24.match_paper('ISO_A3', papers), 'A3')
        self.assertEqual(pdf24.match_paper('A4', papers), 'A4 (210 x 297 mm)')
        self.assertEqual(pdf24.match_paper('ANSI_D', papers), 'ANSI D')
        self.assertIsNone(pdf24.match_paper('A2', papers))
        self.assertIsNone(pdf24.match_paper('Use Sheet Size', papers))
        self.assertIsNone(pdf24.match_paper('A1', ['A10']))

    def test_join_command(self):
        cmd = pdf24.join_command('C:/PDF24/pdf24-DocTool.exe', ['a.pdf', 'b.pdf'], 'C:/out', 'A101-A102_Combined')
        self.assertEqual(cmd[:2], ['C:/PDF24/pdf24-DocTool.exe', '-join'])
        self.assertIn('special/justAssemble', cmd)
        self.assertEqual(cmd[cmd.index('-outputFile') + 1], 'A101-A102_Combined.pdf')
        self.assertEqual(cmd[-2:], ['a.pdf', 'b.pdf'])

    def test_doctool_lookup(self):
        env = {'ProgramFiles': 'C:/PF', 'ProgramFiles(x86)': 'C:/PF86'}
        want = os.path.join('C:/PF86', 'PDF24', pdf24.DOCTOOL)
        self.assertEqual(pdf24.find_doctool(env, exists=lambda p: p == want), want)
        self.assertIsNone(pdf24.find_doctool(env, exists=lambda p: False))


class AutoSaveTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.state = os.path.join(self.temp.name, 'state')

    def test_start_sets_auto_save_and_restore_puts_user_values_back(self):
        user = {'Handler': ('assistant', pdf24.REG_SZ), 'AutoSaveOpenDir': (1, pdf24.REG_DWORD)}
        reg = FakeRegistry(hkcu=user)
        session = pdf24.AutoSave(reg, self.state)
        session.start('D:/Exports')
        session.set_filename('A101 - 50%')
        cur = reg.hkcu()
        self.assertEqual(cur['Handler'][0], 'autoSave')
        self.assertEqual(cur['AutoSaveDir'][0], 'D:/Exports')
        self.assertEqual(cur['AutoSaveFilename'][0], 'A101 - 50%%')
        self.assertEqual(cur['AutoSaveOverwriteFile'][0], 1)
        self.assertEqual(cur['AutoSaveOpenDir'][0], 0)
        self.assertEqual(cur['AutoSaveUseFileChooser'][0], 0)
        self.assertTrue(os.path.exists(session.state_path))
        session.restore()
        self.assertEqual(reg.hkcu(), user)
        self.assertFalse(os.path.exists(session.state_path))

    def test_crashed_run_is_undone_before_the_next_snapshot(self):
        user = {'Handler': ('creator', pdf24.REG_SZ)}
        reg = FakeRegistry(hkcu=user)
        pdf24.AutoSave(reg, self.state).start('D:/A')      # Revit dies: no restore()
        self.assertEqual(reg.hkcu()['Handler'][0], 'autoSave')
        session = pdf24.AutoSave(reg, self.state)
        session.start('D:/B')
        session.restore()
        self.assertEqual(reg.hkcu(), user)

    def test_recover_alone(self):
        reg = FakeRegistry()
        pdf24.AutoSave(reg, self.state).start('D:/A')
        self.assertTrue(pdf24.AutoSave(reg, self.state).recover())
        self.assertEqual(reg.hkcu(), {})
        self.assertFalse(pdf24.AutoSave(reg, self.state).recover())

    def test_admin_lock_blocks(self):
        reg = FakeRegistry(hklm={'NoUserOverride': (1, pdf24.REG_DWORD)})
        self.assertIn('NoUserOverride', pdf24.AutoSave(reg, self.state).blocked_reason())
        self.assertEqual(pdf24.AutoSave(FakeRegistry(), self.state).blocked_reason(), '')


class WaitTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = self.temp.name
        self.now = [0.0]

    def clock(self):
        return self.now[0]

    def sleep(self, s):
        self.now[0] += s

    def write(self, name, data=b'%PDF-1.7'):
        path = os.path.join(self.folder, name)
        with open(path, 'wb') as f:
            f.write(data)
        return path

    def wait(self, expected, **kw):
        kw.setdefault('timeout', 5)
        return pdf24.wait_for_pdf(expected, kw.pop('prev', None), kw.pop('started', time.time()),
                                  clock=self.clock, sleep=self.sleep, **kw)

    def test_expected_file_once_stable(self):
        path = self.write('A101.pdf')
        self.assertEqual(self.wait(path), path)

    def test_missing_file_times_out(self):
        self.assertIsNone(self.wait(os.path.join(self.folder, 'A101.pdf'), timeout=1))

    def test_old_file_is_not_taken_for_new(self):
        path = self.write('A101.pdf')
        os.utime(path, (1000, 1000))
        self.assertIsNone(self.wait(path, prev=1000, started=2000, timeout=1))

    def test_pdf24_own_name_is_accepted_unless_taken(self):
        other = self.write('2026-11-17 10-00-00 A101.pdf')
        expected = os.path.join(self.folder, 'A101.pdf')
        self.assertEqual(self.wait(expected), other)
        self.assertIsNone(self.wait(expected, taken=[other], timeout=1))


class VirtualPrinterTests(unittest.TestCase):
    """Revit can class PDF24 as a virtual printer and refuse PrintToFile = False
    (owner, 2026-11-17): it prints to a file, PostScript goes through Ghostscript."""

    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)

    def write(self, name, data):
        path = os.path.join(self.temp.name, name)
        with open(path, 'wb') as f:
            f.write(data)
        return path

    def test_sniff(self):
        self.assertEqual(pdf24.sniff(self.write('a.pdf', b'%PDF-1.7\n...')), 'pdf')
        self.assertEqual(pdf24.sniff(self.write('b.prn', b'%!PS-Adobe-3.0\n')), 'ps')
        pjl = b'\x1b%-12345X@PJL JOB\r\n@PJL ENTER LANGUAGE=POSTSCRIPT\r\n%!PS-Adobe-3.0\n'
        self.assertEqual(pdf24.sniff(self.write('c.prn', pjl)), 'ps')
        self.assertEqual(pdf24.sniff(self.write('d.prn', b'\x00\x01binary')), 'other')
        self.assertIsNone(pdf24.sniff(os.path.join(self.temp.name, 'missing')))

    def test_ghostscript(self):
        env = {'ProgramFiles': 'C:/PF'}
        want = os.path.join('C:/PF', 'PDF24', 'gs', 'bin', pdf24.GS_EXE)
        self.assertEqual(pdf24.find_ghostscript(env, exists=lambda p: p == want), want)
        cmd = pdf24.ps_to_pdf_command('gs.exe', 'in.prn', 'out.pdf')
        self.assertIn('-sDEVICE=pdfwrite', cmd)
        self.assertIn('-sOutputFile=out.pdf', cmd)
        self.assertEqual(cmd[-1], 'in.prn')
        self.assertTrue(pdf24.prn_path('C:/out', 'A101').endswith('A101' + pdf24.PRN_SUFFIX))

    def test_print_to_file_output_is_found(self):
        prn = self.write('A101' + pdf24.PRN_SUFFIX, b'%!PS-Adobe-3.0\n')
        expected = os.path.join(self.temp.name, 'A101.pdf')
        now = [0.0]
        got = pdf24.wait_for_pdf(expected, None, time.time(), 2, extra=[prn],
                                 clock=lambda: now[0], sleep=lambda s: now.__setitem__(0, now[0] + s))
        self.assertEqual(got, prn)

    def test_dialog_prints_to_file_on_a_virtual_printer(self):
        with open(os.path.join(LIB, 'GUI', 'BatchOutDialog.py'), encoding='utf-8') as f:
            src = f.read()
        body = src.split('def _export_pdf_with_pdf24', 1)[1].split('\n    def _pdf24_paper_for', 1)[0]
        self.assertNotIn('pm.PrintToFile = False', body)
        self.assertIn('pm.PrintToFile = virtual', body)
        self.assertIn("'{}'.format(pm.IsVirtual) != 'None'", body)
        self.assertIn('self._pdf24_as_pdf(found, expected_file, gs)', body)
        self.assertIn('pm.PrintToFileName = print_file', src)


class WiringTests(unittest.TestCase):
    def read(self, *parts):
        with open(os.path.join(LIB, *parts), encoding='utf-8') as f:
            return f.read()

    def test_dialog_routes_pdf_to_pdf24_and_saves_the_choice(self):
        src = self.read('GUI', 'BatchOutDialog.py')
        self.assertIn('self.PDFEngine = "Revit"', src)
        self.assertIn("'PDFEngine': self.PDFEngine", src)
        self.assertIn('profile.PDFEngine = self._pdf_engine()', src)
        self.assertRegex(src, r'if self\._pdf_engine\(\) == "PDF24":\s+return self\._export_pdf_with_pdf24')
        body = src.split('def _export_pdf_with_pdf24', 1)[1].split('\n    def ', 1)[0]
        self.assertIn('session.restore()', body.split('finally:', 1)[1])
        self.assertIn('group.RollBack()', body.split('finally:', 1)[1])
        self.assertIn('SelectNewPrintDriver(old_printer)', body)

    def test_xaml_has_the_engine_picker(self):
        xaml = self.read('GUI', 'Tools', 'ExportManager.xaml')
        self.assertRegex(xaml, r'x:Name="pdf_engine"[^>]*SelectionChanged="pdf_engine_changed"')
        self.assertIn('x:Name="pdf24_status"', xaml)
        items = re.findall(r'<ComboBoxItem Content="([^"]+)"/>',
                           xaml.split('x:Name="pdf_engine"', 1)[1].split('</ComboBox>', 1)[0])
        self.assertEqual(items, ['Revit (built-in)', 'PDF24 printer'])


if __name__ == '__main__':
    unittest.main()
