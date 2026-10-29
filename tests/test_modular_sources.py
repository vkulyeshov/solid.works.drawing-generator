"""Verify source manifest guards and module changes reaching the native binary."""
from pathlib import Path
import hashlib
import json
import sys
import tempfile
import unittest

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from build_swp import Compound, decompress
from project_paths import work_path
from rebuild_native import rebuild
from vba_sources import read_source, source_digest, source_files


class ModularSourceTests(unittest.TestCase):
    def setUp(self):
        root = work_path('module_tests')
        root.mkdir(parents=True, exist_ok=True)
        self.directory = tempfile.TemporaryDirectory(dir=root)
        self.addCleanup(self.directory.cleanup)
        self.project = Path(self.directory.name) / 'modules'
        self.project.mkdir()
        self.names = ['state.bas', 'helpers.bas', 'main.bas']
        self.write_module('state', 'Private message As String\n')
        self.write_module('helpers', 'Private Sub SetMessage()\n    message = "Ready"\nEnd Sub\n')
        self.write_module('main', 'Public Sub main()\n    SetMessage\nEnd Sub\n')
        self.write_manifest()

    def write_module(self, name, body):
        path = self.project / (name + '.bas')
        path.write_text(f'Attribute VB_Name = "{name}"\nOption Explicit\n{body}', encoding='utf-8')

    def write_manifest(self):
        (self.project / 'project.json').write_text(
            json.dumps({'version':1, 'modules':self.names}), encoding='ascii')

    def test_composition_keeps_one_header_and_private_scope(self):
        text = read_source(self.project)
        self.assertEqual(text.count('Attribute VB_Name'), 1)
        self.assertEqual(text.count('Option Explicit'), 1)
        self.assertIn('Private message As String', text)
        self.assertIn('Private Sub SetMessage()', text)
        self.assertLess(text.index('Private message'), text.index('Private Sub'))
        self.assertEqual([path.name for path in source_files(self.project)],
                         ['project.json', *self.names])

    def test_module_edit_reaches_binary_and_build_provenance(self):
        output = Path(self.directory.name) / 'generator.swp'
        rebuild(self.project, output)
        previous = output.read_bytes()
        before_digest = source_digest(self.project)
        self.write_module('helpers', 'Private Sub SetMessage()\n    message = "Updated"\nEnd Sub\n')
        rebuild(self.project, output)
        native = Compound(output)
        module = next(index for index in native.streams if native.name(index) == 'Macro11')
        self.assertIn(b'"Updated"', decompress(native.streams[module]))
        self.assertNotEqual(output.read_bytes(), previous)
        self.assertNotEqual(source_digest(self.project), before_digest)
        report = json.loads(output.with_suffix('.build.json').read_text('ascii'))
        self.assertEqual(report['source_sha256'], source_digest(self.project))
        self.assertEqual(report['source_files'], [
            {'path':str(path.resolve()), 'sha256':hashlib.sha256(path.read_bytes()).hexdigest()}
            for path in source_files(self.project)])

    def test_missing_module_is_rejected(self):
        (self.project / 'helpers.bas').unlink()
        with self.assertRaisesRegex(ValueError, 'Missing or external'):
            read_source(self.project)

    def test_unlisted_module_is_rejected(self):
        self.write_module('forgotten', "' A source omitted from the manifest\n")
        with self.assertRaisesRegex(ValueError, 'Unlisted'):
            read_source(self.project)

    def test_duplicate_module_is_rejected(self):
        self.names.append('helpers.bas')
        self.write_manifest()
        with self.assertRaisesRegex(ValueError, 'Duplicate VBA source module'):
            read_source(self.project)

    def test_external_module_paths_are_rejected(self):
        for name in ['../helpers.bas', 'C:/helpers.bas', 'child/helpers.bas']:
            with self.subTest(name=name):
                self.names = [name]
                self.write_manifest()
                with self.assertRaisesRegex(ValueError, 'Invalid VBA module filename'):
                    read_source(self.project)

    def test_mismatched_module_header_is_rejected(self):
        path = self.project / 'helpers.bas'
        path.write_text(path.read_text('utf-8').replace('"helpers"', '"wrong"'), encoding='utf-8')
        with self.assertRaisesRegex(ValueError, 'matching module attribute'):
            read_source(self.project)

    def test_duplicate_procedures_are_rejected_case_insensitively(self):
        self.write_module('main', 'Public Sub main()\nEnd Sub\nPrivate Sub SETMESSAGE()\nEnd Sub\n')
        with self.assertRaisesRegex(ValueError, 'Duplicate VBA procedure'):
            read_source(self.project)

    def test_missing_entry_point_is_rejected(self):
        self.write_module('main', 'Public Sub AnotherEntry()\nEnd Sub\n')
        with self.assertRaisesRegex(ValueError, 'exactly one Public Sub main'):
            read_source(self.project)

    def test_late_declarations_are_rejected(self):
        self.names = ['helpers.bas', 'state.bas', 'main.bas']
        self.write_manifest()
        with self.assertRaisesRegex(ValueError, 'declarations must precede'):
            read_source(self.project)

    def test_missing_final_newline_does_not_join_statements(self):
        path = self.project / 'helpers.bas'
        path.write_text(path.read_text('utf-8').rstrip(), encoding='utf-8')
        self.assertIn('End Sub\nPublic Sub main()', read_source(self.project))


if __name__ == '__main__':
    unittest.main(verbosity=2)
