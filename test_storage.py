import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

from storage import Library, BusyError, ConflictError

PID = '123456-654321-FSA'
OTHER = '123456-654322-FSA-HGR'


class StorageTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.lib = Library(self.root / 'shared')
        self.source = self.root / 'input.docx'
        self.source.write_bytes(b'word version one')
        self.pdf = self.root / 'input.pdf'
        self.pdf.write_bytes(b'%PDF-test')

    def tearDown(self):
        self.tmp.cleanup()

    def save(self, expected=None):
        return self.lib.commit_record(PID, 'record1', '正式-4样本', self.source,
                                      self.pdf, 4, '2026-11-08 12:00', expected)

    def test_create_idempotent_and_validate(self):
        self.lib.create(PID)
        self.lib.create(PID)
        self.assertEqual(len(self.lib.list_projects()), 1)
        with self.assertRaises(ValueError):
            self.lib.create('../bad')

    def test_relative_paths_survive_library_move(self):
        record = self.save()
        self.assertFalse(Path(record['source']).is_absolute())
        moved = self.root / 'moved'
        self.lib.root.rename(moved)
        self.assertEqual(Library(moved).resolve(record['source']).read_bytes(), b'word version one')

    def test_update_retains_history_and_rejects_stale_revision(self):
        first = self.save()
        self.source.write_bytes(b'version two')
        second = self.save(first['version'])
        self.assertEqual(len(second['history']), 1)
        self.assertEqual(self.lib.resolve(first['source']).read_bytes(), b'word version one')
        with self.assertRaises(ConflictError):
            self.save(first['version'])
        self.assertEqual(self.lib.load(PID)['records'][0]['version'], second['version'])

    def test_expiry_conflict_and_independent_project(self):
        self.lib.create(PID)
        self.lib.create(OTHER)
        old = self.lib.load(PID)
        self.lib.set_expiry(PID, '2026-12-10', old['revision'])
        with self.assertRaises(ConflictError):
            self.lib.set_expiry(PID, None, old['revision'])
        self.assertIsNone(self.lib.load(OTHER)['expiry'])

    def test_failed_copy_does_not_publish(self):
        first = self.save()
        with patch('storage.shutil.copyfile', side_effect=OSError('network interrupted')):
            with self.assertRaises(OSError):
                self.save(first['version'])
        self.assertEqual(self.lib.load(PID)['records'][0]['version'], first['version'])

    def test_failed_atomic_replace_does_not_publish(self):
        first = self.save()
        with patch('storage.os.replace', side_effect=OSError('network interrupted')):
            with self.assertRaises(OSError):
                self.save(first['version'])
        self.assertEqual(self.lib.load(PID)['records'][0]['version'], first['version'])
        self.assertFalse(list(self.lib.folder(PID).glob('.pending-*')))

    def test_no_pdf_update_does_not_reuse_old_pdf(self):
        first = self.save()
        second = self.lib.commit_record(PID, 'record1', '正式-4样本', self.source,
                                        None, None, '2026-11-09', first['version'])
        self.assertEqual(second['pdf'], '')
        self.assertTrue(second['history'][0]['pdf'])

    def test_cross_process_lock_release_after_crash(self):
        code = '''import sys, time
from storage import Library
lib = Library(sys.argv[1])
with lib.lock(sys.argv[2]):
 print('locked', flush=True)
 time.sleep(30)
'''
        proc = subprocess.Popen([sys.executable, '-u', '-c', code, str(self.lib.root), PID],
                                cwd=Path(__file__).parent, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                text=True)
        try:
            self.assertEqual(proc.stdout.readline().strip(), 'locked')
            with self.assertRaises(BusyError):
                self.lib.create(PID)
            self.lib.create(OTHER)
        finally:
            proc.kill()
            proc.communicate(timeout=10)
        self.lib.create(PID)
        self.assertEqual(len(self.lib.list_projects()), 2)

    def test_path_escape_rejected(self):
        with self.assertRaises(ValueError):
            self.lib.resolve('../outside.docx')

    def test_corruption_is_reported(self):
        self.lib.create(PID)
        (self.lib.folder(PID) / 'project.json').write_text('{broken', encoding='utf-8')
        with self.assertRaises(ValueError):
            self.lib.list_projects()


if __name__ == '__main__':
    unittest.main()
