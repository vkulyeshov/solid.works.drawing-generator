"""Validate benchmark clock accounting and disposable source instrumentation."""
from contextlib import redirect_stdout
import io
import json
from pathlib import Path
import sys
import tempfile
import unittest

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from project_paths import work_path
from vba_sources import read_source
from benchmark_generation import instrument, report


class BenchmarkTests(unittest.TestCase):
    def setUp(self):
        root = work_path('benchmark/unit')
        root.mkdir(parents=True, exist_ok=True)
        self.temporary = tempfile.TemporaryDirectory(dir=root)
        self.addCleanup(self.temporary.cleanup)
        self.directory = Path(self.temporary.name)

    def test_midnight_job_and_overlapping_logging_cost(self):
        output = self.directory / 'output'
        output.mkdir()
        (self.directory / 'execution-timings.tsv').write_text(
            'stage\tseconds\napplication-start\t8\ninput-open\t1\nmacro\t34\n'
            'test-inspection\t90\ntest-shutdown\t5\ntest-total\t138\n', 'utf-16')
        (output / 'run.log').write_text(
            '23:59:55 JOB example\n23:59:57 STAGE example | Creating drawing\n'
            '00:00:10 STAGE example | Planning dimensions\n'
            '00:00:20 STAGE example | Exporting PDF\n00:00:22 OK example\n'
            '00:00:23 BATCH END drawings=1 failed=0\n', 'utf-16')
        (output / 'benchmark-layout.tsv').write_text(
            '1\tcandidate generation\t1,25\n1\tcandidate generation\t2,25\n'
            '1\tnative verification\t1\n', 'utf-16')
        (output / 'benchmark-macro.tsv').write_text(
            'macro-body\t33,75\nlogging-overlap\t1,25\n', 'utf-16')
        with redirect_stdout(io.StringIO()):
            report(self.directory)
        result = json.loads((self.directory / 'benchmark-report.json').read_text())
        self.assertEqual(result['job_seconds'], 27)
        self.assertEqual(result['layout_phases_seconds']['candidate generation'], 3.5)
        self.assertEqual(result['execution_seconds']['macro'], 34)
        self.assertEqual(result['cold_generation_seconds'], 43)
        self.assertEqual(result['macro_probe_seconds']['logging-overlap'], 1.25)

    def test_instrumented_copy_passes_source_validation(self):
        canonical = read_source()
        source = self.directory / 'benchmark.bas'
        source.write_text(instrument(canonical), 'utf8')
        result = read_source(source)
        self.assertIn('BenchmarkMark "candidate generation"', result)
        self.assertIn('BenchmarkMark "native verification"', result)
        self.assertNotIn('BenchmarkMark', canonical)
        self.assertEqual(read_source(), canonical)


if __name__ == '__main__':
    unittest.main(verbosity=2)
