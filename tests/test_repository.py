"""Check source consistency, ignored build artifacts, and external test outputs."""
from pathlib import Path
import ast
import os
import re
import sys
import subprocess
import unittest
from unittest.mock import patch

sys.dont_write_bytecode = True
REPOSITORY = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPOSITORY/'tools'))
from build_swp import Compound, decompress
from project_paths import BUILD_DIRECTORY, build_path, external_path, work_path
from rebuild_native import rebuild, module_source
from vba_project import empty_project
from vba_sources import DEFAULT_SOURCE, read_source, source_files


def source_paths():
    """Prune Git metadata and the ignored release tree before walking sources."""
    for parent, folders, files in os.walk(REPOSITORY):
        folders[:] = [name for name in folders if name != '.git'
                      and not (Path(parent) == REPOSITORY and name == 'build')]
        for name in folders + files:
            yield Path(parent) / name


class RepositoryTests(unittest.TestCase):
    def test_native_source_matches_modules(self):
        path = rebuild(DEFAULT_SOURCE, work_path('repository_tests/source.swp'))
        container = Compound(path)
        module = next(i for i in container.streams if container.name(i) == 'Macro11')
        expected = module_source(DEFAULT_SOURCE)
        self.assertEqual(decompress(container.streams[module]), expected)

    def test_canonical_sources_live_in_modules(self):
        self.assertTrue((DEFAULT_SOURCE/'project.json').is_file())
        self.assertFalse((REPOSITORY/'generator.bas').exists())
        self.assertFalse((REPOSITORY/'generator.utf8.txt').exists())
        self.assertFalse((REPOSITORY/'generator.swp').exists())
        source = read_source()
        self.assertEqual(source.count('Public Sub main()'), 1)

    def test_empty_project_has_no_generator_code(self):
        project = empty_project()
        module = next(i for i in project.streams if project.name(i) == 'Macro11')
        self.assertEqual(project.streams[module], b'')

    def test_external_output_guards(self):
        for path in [REPOSITORY, REPOSITORY/'tmp', REPOSITORY/'tools/../output', BUILD_DIRECTORY]:
            with self.assertRaises(ValueError):
                external_path(path)
        with patch.dict(os.environ, {'SW_DRAWING_WORK_DIR':str(REPOSITORY/'temporary')}):
            with self.assertRaises(ValueError):
                work_path('build')
        self.assertNotIn(REPOSITORY, work_path('unit').parents)

    def test_build_output_guards(self):
        self.assertEqual(build_path(), BUILD_DIRECTORY)
        self.assertEqual(build_path(BUILD_DIRECTORY/'bin/generator.swp'), BUILD_DIRECTORY/'bin/generator.swp')
        self.assertEqual(build_path(work_path('package')), work_path('package'))
        for path in [REPOSITORY, REPOSITORY/'templates', REPOSITORY/'README.md',
                     BUILD_DIRECTORY/'../tools']:
            with self.assertRaises(ValueError):
                build_path(path)

    def test_rebuild_is_external_and_preserves_source(self):
        source = DEFAULT_SOURCE
        expected = {path:path.read_bytes() for path in source_files(source)}
        output = rebuild(source, work_path('repository_tests/rebuilt.swp'))
        container = Compound(output)
        module = next(i for i in container.streams if container.name(i) == 'Macro11')
        self.assertEqual(decompress(container.streams[module]), module_source(source))
        self.assertEqual({path:path.read_bytes() for path in source_files(source)}, expected)

    def test_source_edits_reach_native_output(self):
        source = work_path('repository_tests/modified.bas')
        source.parent.mkdir(parents=True, exist_ok=True)
        original = read_source()
        source.write_text(original+"\n' Source freshness regression marker\n",encoding='utf8')
        output = rebuild(source, work_path('repository_tests/modified.swp'))
        native = Compound(output)
        module = next(i for i in native.streams if native.name(i) == 'Macro11')
        self.assertIn(b'Source freshness regression marker',decompress(native.streams[module]))

    def test_sources_parse_and_keep_english_prose(self):
        text_suffixes = {'.py','.ps1','.cs','.vbs','.bas','.txt','.md'}
        for path in source_paths():
            if not path.is_file():
                continue
            if path.suffix == '.py':
                ast.parse(path.read_text('utf8'), filename=str(path))
            if path.suffix in text_suffixes:
                text = path.read_text('utf8')
                if path.suffix == '.bas':
                    text = re.sub(r'"(?:[^"\n]|"")*"', '""', text)
                self.assertTrue(text.isascii(), str(path))

    def test_cyrillic_survives_native_source_encoding(self):
        source = DEFAULT_SOURCE
        original = read_source(source)
        literals = re.findall(r'"(?:[^"\n]|"")*"',original)
        localized = [literal for literal in literals if not literal.isascii()]
        self.assertGreater(len(localized), 0)
        path = rebuild(source, work_path('repository_tests/cyrillic.swp'))
        native = Compound(path)
        module = next(i for i in native.streams if native.name(i) == 'Macro11')
        decoded = decompress(native.streams[module]).decode('cp1251')
        for literal in localized:
            self.assertIn(literal, decoded)

    def test_generated_data_is_ignored_or_external(self):
        forbidden_dirs = {'tmp','temp','output','artifacts','build','dist','Ready','bin','Backups','__pycache__'}
        forbidden_suffixes = {'.log','.tmp','.zip','.pdf','.png','.pyc','.swp','.slddrw','.sldprt','.sldasm'}
        for path in source_paths():
            self.assertFalse(path.is_dir() and path.name in forbidden_dirs, str(path))
            self.assertFalse(path.is_file() and path.suffix.lower() in forbidden_suffixes, str(path))
        ignored = subprocess.run(['git', 'check-ignore', '--quiet', '--', 'build/bin/README.md'],
                                 cwd=REPOSITORY)
        self.assertEqual(ignored.returncode, 0, 'The complete build/ tree must be ignored')
        tracked = subprocess.run(['git', 'ls-files', '--', 'build'], cwd=REPOSITORY,
                                 capture_output=True, check=True)
        self.assertEqual(tracked.stdout, b'', 'Build artifacts must not be tracked')

    def test_four_runtime_templates(self):
        self.assertEqual({p.name for p in (REPOSITORY/'templates').iterdir()}, {
            'A4.landscape.drwdot','A4.portrait.drwdot','A3.landscape.drwdot','A3.portrait.drwdot'})


if __name__ == '__main__':
    unittest.main(verbosity=2)
