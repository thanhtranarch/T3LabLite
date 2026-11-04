"""BatchOut export report + profile files (Services/batchout_report.py, batchout_profiles.py)."""
import csv
import json
import os
import sys
import unittest
import zipfile
from datetime import datetime
from tempfile import TemporaryDirectory
from types import SimpleNamespace

LIB = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'lib')
sys.path.insert(0, LIB)

from Services import batchout_profiles as bp  # noqa: E402
from Services import batchout_report as br  # noqa: E402

WHEN = datetime(2026, 11, 4, 9, 30, 15)


def row(number, fmt, status, path='', name='Plan', revision='C'):
    return SimpleNamespace(SheetNumber=number, SheetName=name, Revision=revision, Format=fmt,
                           Size='A1', Orientation='Landscape', Status=status, OutputFile=path)


class ReportRowsTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.pdf = os.path.join(self.temp.name, 'A101.pdf')
        with open(self.pdf, 'wb') as f:
            f.write(b'x' * 4096)

    def test_exported_row_carries_file_and_size(self):
        rows = br.build_rows([row('A101', 'PDF', 'Exported', self.pdf)], {}, 'Tower.rvt', WHEN)
        self.assertEqual(rows[0]['file'], self.pdf)
        self.assertEqual(rows[0]['file_kb'], 4)
        self.assertEqual(rows[0]['reason'], '')
        self.assertEqual(rows[0]['exported_at'], '2026-11-04 09:30')
        self.assertEqual(rows[0]['model'], 'Tower.rvt')

    def test_failed_row_gets_reason_and_no_file(self):
        reasons = {('A102', 'DWG'): 'Revit wrote no DWG file'}
        rows = br.build_rows([row('A102', 'DWG', 'Failed', self.pdf)], reasons, '', WHEN)
        self.assertEqual(rows[0]['reason'], 'Revit wrote no DWG file')
        self.assertEqual(rows[0]['file'], '')
        self.assertEqual(rows[0]['file_kb'], '')

    def test_reason_is_per_format(self):
        reasons = {('A102', 'DWG'): 'boom'}
        rows = br.build_rows([row('A102', 'PDF', 'Failed')], reasons, '', WHEN)
        self.assertEqual(rows[0]['reason'], '')

    def test_unreached_row_reads_not_exported(self):
        rows = br.build_rows([row('A103', 'PDF', '')], {}, '', WHEN)
        self.assertEqual(rows[0]['status'], br.NOT_EXPORTED)
        self.assertEqual(br.summarize(rows), {br.NOT_EXPORTED: 1})

    def test_columns_normalized(self):
        self.assertEqual(br.normalize_columns(['Status', 'number', 'bogus', 'status']),
                         ['status', 'number'])
        self.assertEqual(br.normalize_columns([]), br.DEFAULT_COLUMNS)
        self.assertEqual(br.normalize_columns(None), br.DEFAULT_COLUMNS)

    def test_format_normalized(self):
        self.assertEqual(br.normalize_format('csv'), 'CSV')
        self.assertEqual(br.normalize_format('pdf'), 'XLSX')
        self.assertEqual(br.normalize_format(None), 'XLSX')

    def test_report_path_is_timestamped(self):
        self.assertEqual(os.path.basename(br.report_path('x', 'CSV', WHEN)),
                         'BatchOut_Report_20261104_093015.csv')
        self.assertTrue(br.report_path('x', 'XLSX', WHEN).endswith('.xlsx'))


class ReportWriteTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.rows = br.build_rows([row('A101', 'PDF', 'Exported', name='Mặt bằng tầng 1'),
                                   row('A102', 'DWG', 'Failed')],
                                  {('A102', 'DWG'): 'Revit wrote no DWG file'}, 'Tower', WHEN)

    def test_csv_has_bom_headers_and_chosen_columns(self):
        path = br.write_report(os.path.join(self.temp.name, 'r.csv'), self.rows,
                               ['number', 'name', 'status', 'reason'], 'CSV')
        with open(path, 'rb') as f:
            self.assertTrue(f.read().startswith(b'\xef\xbb\xbf'))
        with open(path, encoding='utf-8-sig', newline='') as f:
            table = list(csv.reader(f))
        self.assertEqual(table[0], ['Number', 'Name', 'Status', 'Reason'])
        self.assertEqual(table[1], ['A101', 'Mặt bằng tầng 1', 'Exported', ''])
        self.assertEqual(table[2], ['A102', 'Plan', 'Failed', 'Revit wrote no DWG file'])

    def test_xlsx_is_a_workbook_with_every_row(self):
        path = br.write_report(os.path.join(self.temp.name, 'r.xlsx'), self.rows, None, 'XLSX')
        with zipfile.ZipFile(path) as zf:
            names = zf.namelist()
            shared = zf.read('xl/sharedStrings.xml').decode('utf-8')
            sheet = zf.read('xl/worksheets/sheet1.xml').decode('utf-8')
        self.assertIn('xl/workbook.xml', names)
        self.assertIn('Mặt bằng tầng 1', shared)
        self.assertIn('Revit wrote no DWG file', shared)
        self.assertEqual(sheet.count('<row '), 3)  # header + 2


class ProfileFileTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.dir = self.temp.name

    def write(self, name, data):
        path = os.path.join(self.dir, name)
        with open(path, 'w', encoding='utf-8') as f:
            f.write(data if isinstance(data, str) else json.dumps(data))
        return path

    def test_read_rejects_non_profiles(self):
        for content in ('[1, 2]', '{"Name": ""}', '{"ExportPDF": true}', 'not json'):
            with self.assertRaises(ValueError):
                bp.read_profile_file(self.write('bad.json', content))

    def test_read_accepts_bom(self):
        path = os.path.join(self.dir, 'bom.json')
        with open(path, 'w', encoding='utf-8-sig') as f:
            json.dump({'Name': 'Team A1'}, f)
        self.assertEqual(bp.read_profile_file(path)['Name'], 'Team A1')

    def test_list_skips_bookkeeping_and_bad_files(self):
        self.write('_latest_setup.json', {'Name': 'Latest Setup'})
        self.write('_export_crash_history.json', '[]')
        self.write('broken.json', '{')
        self.write('b.json', {'Name': 'beta'})
        self.write('a.json', {'Name': 'Alpha'})
        self.write('notes.txt', 'x')
        self.assertEqual([n for n, _p, _d in bp.list_profiles(self.dir)], ['Alpha', 'beta'])
        self.assertEqual(bp.list_profiles(os.path.join(self.dir, 'missing')), [])

    def test_safe_filename(self):
        self.assertEqual(bp.safe_filename('A1: Issue/For "Tender"?'), 'A1_ Issue_For _Tender__')
        self.assertEqual(bp.safe_filename('_hidden'), 'hidden')
        self.assertEqual(bp.safe_filename('   '), 'Profile')

    def test_unique_path_and_name(self):
        self.write('Team.json', {'Name': 'Team'})
        self.assertTrue(bp.unique_profile_path(self.dir, 'Team').endswith('Team (2).json'))
        self.assertEqual(bp.unique_profile_name(['Team', 'team (2)'], 'Team'), 'Team (3)')
        self.assertEqual(bp.unique_profile_name([], 'New'), 'New')

    def test_write_round_trip_keeps_unicode(self):
        path = bp.write_profile_file(os.path.join(self.dir, 'v.json'), {'Name': 'Hồ sơ'})
        self.assertEqual(bp.read_profile_file(path)['Name'], 'Hồ sơ')


class DialogWiringTests(unittest.TestCase):
    """The dialog keeps its side of the contract (no Revit needed)."""

    def setUp(self):
        src = os.path.join(LIB, 'GUI', 'BatchOutDialog.py')
        with open(src, encoding='utf-8-sig') as f:
            self.source = f.read()

    def test_start_export_writes_report_and_shows_it(self):
        self.assertIn('report_note = self._write_export_report(output_folder)', self.source)
        self.assertIn('skip_note, report_note)', self.source)

    def test_profile_carries_report_settings(self):
        for key in ("'WriteReport'", "'ReportFormat'", "'ReportColumns'"):
            self.assertIn(key, self.source)

    def _shipped_write_report(self):
        import ast
        from unittest.mock import Mock
        tree = ast.parse(self.source)
        cls = next(n for n in tree.body if isinstance(n, ast.ClassDef) and n.name == 'ExportManagerWindow')
        fn = next(n for n in cls.body if isinstance(n, ast.FunctionDef) and n.name == '_write_export_report')
        scope = dict(batchout_report=br, datetime=datetime, logger=Mock())
        exec(compile(ast.Module(body=[fn], type_ignores=[]), 'BatchOutDialog.py', 'exec'), scope)
        return scope['_write_export_report']

    def _window(self, checked=True, fmt='CSV'):
        return SimpleNamespace(
            write_report=SimpleNamespace(IsChecked=checked),
            export_items=[row('A101', 'PDF', 'Exported'), row('A102', 'DWG', 'Failed')],
            _selected_report_format=lambda: fmt, _report_columns=list(br.DEFAULT_COLUMNS),
            _failure_reasons={('A102', 'DWG'): 'no file'}, _export_started_at=WHEN.timestamp(),
            doc=SimpleNamespace(Title='Tower'))

    def test_shipped_method_writes_report_into_output_folder(self):
        write = self._shipped_write_report()
        with TemporaryDirectory() as folder:
            note = write(self._window(), folder)
            files = os.listdir(folder)
        self.assertEqual(files, ['BatchOut_Report_20261104_093015.csv'])
        self.assertIn('Report (2 rows)', note)

    def test_shipped_method_off_writes_nothing(self):
        write = self._shipped_write_report()
        with TemporaryDirectory() as folder:
            self.assertEqual(write(self._window(checked=False), folder), '')
            self.assertEqual(os.listdir(folder), [])

    def test_shipped_method_never_raises(self):
        write = self._shipped_write_report()
        missing = os.path.join(LIB, 'no', 'such', 'folder')
        self.assertIn('was not written', write(self._window(), missing))
        broken = self._window()
        del broken.write_report  # control missing: still no exception
        self.assertIn('was not written', write(broken, missing))

    def test_confirmed_exports_record_their_file(self):
        self.assertIn("item.OutputFile = getattr(self, '_last_output_file', '')", self.source)
        self.assertGreaterEqual(self.source.count('self._last_output_file = '), 5)


if __name__ == '__main__':
    unittest.main(verbosity=1)
