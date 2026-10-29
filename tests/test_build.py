"""Verify the standalone VBS builder and runtime package without starting CAD."""
from pathlib import Path
import sys
sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from project_paths import REPOSITORY, BUILD_DIRECTORY, work_path
import hashlib
import json
import os
import re
import shutil
import subprocess
import tempfile
import unittest
import zipfile
from vba_sources import source_digest, source_files


class BuildTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.host = Path(os.environ['WINDIR']) / 'System32/cscript.exe'
        cls.root = work_path('package_build')
        cls.root.mkdir(parents=True, exist_ok=True)
        cls.environment = dict(os.environ, SW_DRAWING_PYTHON=sys.executable)

    def invoke(self, script, arguments=(), code=0, cwd=REPOSITORY):
        result = subprocess.run([str(self.host), '//nologo', str(script), *map(str, arguments)],
                                cwd=cwd, env=self.environment,
                                capture_output=True, timeout=45)
        output = result.stdout + result.stderr
        self.assertEqual(result.returncode, code, output)
        return output

    def test_main_is_the_only_module_level_execution(self):
        for name in ['build.vbs', 'launcher.vbs']:
            source = (REPOSITORY / name).read_text('utf8')
            self.assertTrue(source.rstrip().endswith('\nMain'), name)
            self.assertEqual(len(re.findall(r'(?mi)^Sub Main\(\)', source)), 1, name)
            module = re.sub(r'(?ms)^(?:Sub|Function) .*?^End (?:Sub|Function)\s*', '', source)
            statements = [line.strip() for line in module.splitlines()
                          if line.strip() and not line.lstrip().startswith("'")]
            for line in statements[:-1]:
                self.assertTrue(line == 'Option Explicit' or line.startswith(('Const ', 'Dim ')),
                                (name, line))
            self.assertEqual(statements[-1], 'Main', name)

    def test_help_does_not_build(self):
        for name in ['build.vbs', 'launcher.vbs']:
            output = self.invoke(REPOSITORY / name, ['/help'])
            self.assertNotIn(b'[build] Source:', output)

    def test_invalid_build_arguments(self):
        for arguments in [['/package:directory'], ['/input:directory'], ['/unknown:value'], ['one', 'two'], ['']]:
            output = self.invoke(REPOSITORY / 'build.vbs', arguments, code=1)
            self.assertNotIn(b'[build] Source:', output)

    def test_failed_builder_exit_is_preserved(self):
        with tempfile.TemporaryDirectory(prefix='failed_build_', dir=self.root) as folder:
            checkout = Path(folder)
            shutil.copy2(REPOSITORY / 'build.vbs', checkout / 'build.vbs')
            (checkout / 'modules').mkdir()
            (checkout / 'modules/project.json').write_text('{}', encoding='ascii')
            (checkout / 'tools').mkdir()
            (checkout / 'tools/package_launcher.py').write_text(
                "import sys\nprint('Injected build failure', flush=True)\nsys.exit(9)\n",
                encoding='ascii')
            output = self.invoke(checkout / 'build.vbs', [checkout / 'package'], code=9)
            self.assertIn(b'Injected build failure', output)
            self.assertNotIn(b'Package ready:', output)

    def test_real_build_produces_self_contained_current_package(self):
        package = self.root / 'package \u0442\u0435\u0441\u0442 with spaces'
        result = self.invoke(REPOSITORY / 'build.vbs', [package])
        self.assertIn(b'Package ready:', result)
        bin_dir = package / 'bin'
        expected = {'launcher.vbs', 'generator.swp', 'generator.ini',
                    'README.md', *('templates/' + p.name for p in (REPOSITORY / 'templates').iterdir())}
        report = json.loads((package / 'generator.build.json').read_text('ascii'))
        digest = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
        self.assertEqual(report['source_sha256'], source_digest())
        self.assertEqual(report['source_files'], [
            {'path':str(path.resolve()), 'sha256':digest(path)} for path in source_files()])
        self.assertEqual(report['macro_sha256'], digest(bin_dir / 'generator.swp'))
        self.assertEqual((bin_dir / 'launcher.vbs').read_bytes(), (REPOSITORY / 'launcher.vbs').read_bytes())
        self.assertEqual((bin_dir / 'README.md').read_bytes(), (REPOSITORY / 'GENERATOR.md').read_bytes())
        self.assertFalse((bin_dir / 'generator.build.json').exists())
        self.assertFalse((package / 'Ready').exists())
        self.assertFalse((package / 'SolidWorks_DrawingMacro.zip').exists())
        with zipfile.ZipFile(package / 'bin.zip') as archive:
            self.assertEqual(set(archive.namelist()), expected)
            for name in expected:
                self.assertEqual(archive.read(name), (bin_dir / name).read_bytes())

    def test_build_migrates_legacy_directory_and_archive(self):
        with tempfile.TemporaryDirectory(prefix='legacy_build_', dir=self.root) as folder:
            package = Path(folder)
            old = package / 'Ready'
            old.mkdir()
            (old / 'generator.bas').write_text('Obsolete source copy', encoding='ascii')
            (old / 'generator.build.json').write_text('{}', encoding='ascii')
            (package / 'SolidWorks_DrawingMacro.zip').write_bytes(b'Obsolete archive')
            self.invoke(REPOSITORY / 'build.vbs', [package])
            self.assertFalse(old.exists())
            self.assertFalse((package / 'SolidWorks_DrawingMacro.zip').exists())
            self.assertFalse((package / 'bin/generator.bas').exists())
            self.assertFalse((package / 'bin/generator.build.json').exists())
            self.assertTrue((package / 'generator.build.json').is_file())
            self.assertTrue((package / 'bin.zip').is_file())

    def test_default_build_uses_ignored_repository_directory(self):
        self.invoke(REPOSITORY / 'build.vbs', cwd=self.root)
        self.assertTrue((BUILD_DIRECTORY / 'bin/launcher.vbs').is_file())
        self.assertTrue((BUILD_DIRECTORY / 'bin.zip').is_file())
        self.assertTrue((BUILD_DIRECTORY / 'generator.build.json').is_file())
        ignored = subprocess.run(['git', 'check-ignore', '--quiet', '--', 'build/bin/README.md'],
                                 cwd=REPOSITORY)
        self.assertEqual(ignored.returncode, 0)

    def test_relative_output_argument_uses_current_directory(self):
        self.invoke(REPOSITORY / 'build.vbs', ['relative output'], cwd=self.root)
        self.assertTrue((self.root / 'relative output/bin/launcher.vbs').is_file())
        self.assertTrue((self.root / 'relative output/bin.zip').is_file())


if __name__ == '__main__':
    unittest.main(verbosity=2)
