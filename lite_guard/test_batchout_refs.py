"""BatchOut DWG image-reference cleanup (Services/batchout_refs.py)."""
import os
import sys
import unittest
from tempfile import TemporaryDirectory

LIB = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'lib')
sys.path.insert(0, LIB)

from Services import batchout_refs as refs  # noqa: E402


class ReferenceFileTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = self.temp.name

    def write(self, name, mtime=None):
        path = os.path.join(self.folder, name)
        with open(path, 'wb') as f:
            f.write(b'x')
        if mtime is not None:
            os.utime(path, (mtime, mtime))
        return path

    def names(self, paths):
        return sorted(os.path.basename(p) for p in paths)

    def test_new_images_and_pdfs_are_found(self):
        before = refs.snapshot(self.folder)
        self.write('Model-Aedas_Logo-1464950.pdf')
        self.write('Model-Arup-146.PNG')
        self.write('Model-Beca.jpg')
        self.write('A101.dwg')
        found = refs.written_reference_files(self.folder, before)
        self.assertEqual(self.names(found),
                         ['model-aedas_logo-1464950.pdf', 'model-arup-146.png', 'model-beca.jpg'])

    def test_untouched_existing_files_are_left(self):
        self.write('A101.pdf', mtime=1000)
        self.write('photo.png', mtime=1000)
        before = refs.snapshot(self.folder)
        self.write('A101.dwg')
        self.assertEqual(refs.written_reference_files(self.folder, before), [])

    def test_rewritten_reference_from_an_earlier_run_is_found(self):
        self.write('Model-Logo.png', mtime=1000)
        before = refs.snapshot(self.folder)
        self.write('Model-Logo.png', mtime=2000)
        found = refs.written_reference_files(self.folder, before)
        self.assertEqual(self.names(found), ['model-logo.png'])

    def test_keep_paths_and_other_types_are_never_returned(self):
        before = refs.snapshot(self.folder)
        kept = self.write('A101.pdf')
        self.write('A101.dwg')
        self.write('A101.pcp')
        self.write('notes.txt')
        self.assertEqual(refs.written_reference_files(self.folder, before, keep=[kept]), [])

    def test_missing_folder_gives_nothing(self):
        gone = os.path.join(self.folder, 'nope')
        self.assertEqual(refs.snapshot(gone), {})
        self.assertEqual(refs.written_reference_files(gone, {}), [])

    def test_delete_reports_deleted_and_failed(self):
        path = self.write('Model-Logo.png')
        missing = os.path.join(self.folder, 'gone.png')
        deleted, failed = refs.delete_files([path, missing])
        self.assertEqual(deleted, [path])
        self.assertEqual(failed, [missing])
        self.assertFalse(os.path.exists(path))


if __name__ == '__main__':
    unittest.main()
