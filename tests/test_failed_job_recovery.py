"""Verify native queue continuation, bounded retries, and version 2 migration.

Run test_paper_formats.py --case A4.landscape first in the same external work
directory. This test extends that disposable queue; never use production output.
"""
from pathlib import Path
import sys
sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from project_paths import work_path
import hashlib
import json
import os
import subprocess

base = Path(__file__).resolve().parents[1]
fixture = work_path('paper_formats/A4.landscape')
pack = fixture / 'out'
state = pack / 'generation-state.tsv'
lines = [line.split('\t') for line in state.read_text('utf-16').splitlines()]
assert len(lines) == 2 and lines[1][10] == 'DONE', 'Use a fresh single-job test fixture'
sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
protected = {p: sha(p) for folder in ['Drawings', 'PDF'] for p in (pack / folder).iterdir() if p.is_file()}
# A completed legacy job, a bad body, then a valid pending body. Batch size one
# forces the controller to continue even when a batch produces no drawings.
header, done = lines
header[1] = '2'
done = done[:15]
bad, good = done.copy(), done.copy()
for row, number, stem in [(bad, '2', '002_missing_body'), (good, '3', '003_valid_after_failure')]:
    row[1], row[10], row[11], row[12] = number, 'PENDING', stem, 'Queued by recovery test'
    row[13], row[14] = '0', '0'
bad[4] = '__missing_body_for_recovery_test__'
state.write_bytes(('\r\n'.join('\t'.join(row) for row in [header, done, bad, good]) + '\r\n').encode('utf-16'))
ini = fixture / 'retry-test.ini'
settings = (fixture / 'test.ini').read_text('utf-16')
settings += '\r\nMaxJobAttempts=2\r\nJobsPerBatch=1\r\n'
ini.write_bytes(settings.encode('utf-16'))
environment = dict(os.environ, SW_DRAWING_PYTHON=sys.executable)
command = [str(Path(os.environ['WINDIR']) / 'System32/cscript.exe'), '//nologo', str(base / 'launcher.vbs'),
           '/input:' + str(fixture / 'input/CarFenderGusset.SLDPRT'), '/ini:' + str(ini),
           '/out:' + str(pack), '/format:A4', '/orientation:landscape', '/background']
log = work_path('failed-job-recovery.log')
with log.open('wb') as stream:
    run = subprocess.run(command, stdout=stream, stderr=subprocess.STDOUT, env=environment, timeout=600)
text = log.read_bytes().decode('cp866', errors='replace')
print(text[-7000:])
assert run.returncode == 2, (run.returncode, text[-2000:])
rows = [line.split('\t') for line in state.read_text('utf-16').splitlines()]
assert rows[0][1] == '3', 'Legacy checkpoint was not migrated'
jobs = rows[1:]
assert [row[10] for row in jobs] == ['DONE', 'FAILED', 'DONE'], jobs
assert [int(row[15]) for row in jobs] == [1, 2, 1], jobs
assert text.count('; checkpoint completed=') == 3, 'Expected failure, next job, then deferred retry'
assert all(sha(path) == digest for path, digest in protected.items()), 'Completed artifacts changed'
report = (pack / 'failed-jobs.tsv').read_text('utf-16').splitlines()
assert len(report) == 2 and 'Body not found' in report[1], report
assert report[1].split('\t')[4:7] == ['2', '2', 'False'], report
result = {'passed': True, 'exit_code': run.returncode, 'statuses': [row[10] for row in jobs],
          'attempts': [int(row[15]) for row in jobs], 'log': str(log)}
work_path('failed-job-recovery.json').write_text(json.dumps(result, indent=2), encoding='utf8')
print(json.dumps(result, indent=2))
