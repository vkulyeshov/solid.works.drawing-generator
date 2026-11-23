"""Interrupt an intentionally blocked job in an owned background CAD process.

Use only an external fixture and copied input; the production modules and
completed drawing outputs are never changed by this fault-injection test.
"""
from pathlib import Path
import argparse
import os
import subprocess
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from project_paths import REPOSITORY, external_path
from rebuild_native import rebuild
from vba_sources import read_source
from prepare_model_test import prepare


def check(model, directory):
    directory = external_path(directory)
    prepare(model, directory, overrides=['JobTimeoutSeconds=1', 'StartupTimeoutSeconds=60',
                                        'LayoutTimeoutSeconds=0', 'MaxJobAttempts=3'])
    source = read_source()
    declaration = '''
#If VBA7 Then
Private Declare PtrSafe Sub TimeoutProbeSleep Lib "kernel32" Alias "Sleep" (ByVal milliseconds As Long)
#Else
Private Declare Sub TimeoutProbeSleep Lib "kernel32" Alias "Sleep" (ByVal milliseconds As Long)
#End If
Private Const TIMEOUT_PROBE_BLOCK_MILLISECONDS As Long = 60000
'''
    source = source.replace('Option Explicit\n', 'Option Explicit\n' + declaration, 1)
    marker = '    Log "JOB " & outName\n'
    assert source.count(marker) == 1
    source = source.replace(marker, marker + '    TimeoutProbeSleep TIMEOUT_PROBE_BLOCK_MILLISECONDS\n', 1)
    probe = directory / 'timeout-probe.bas'
    probe.write_text(source, 'utf8')
    rebuild(probe, directory / 'generator.swp')
    command = [str(Path(os.environ['WINDIR']) / 'System32/cscript.exe'), '//nologo',
               str(REPOSITORY / 'launcher.vbs'), '/input:' + str(directory / 'input' / model.name),
               '/ini:' + str(directory / 'generator.ini'), '/macro:' + str(directory / 'generator.swp'),
               '/out:' + str(directory / 'output')]
    console = directory / 'controller-console.txt'
    with console.open('wb') as stream:
        result = subprocess.run(command, stdout=stream, stderr=subprocess.STDOUT, timeout=120)
    assert result.returncode == 2, (result.returncode, console.read_bytes())
    output = directory / 'output'
    log = (output / 'launcher.log').read_text('utf16')
    assert 'Terminated timed-out dedicated SolidWorks PID=' in log, log[-2000:]
    assert 'Batch 2;' in log, 'The controller did not resume the checkpoint'
    rows = [line.split('\t') for line in (output / 'generation-state.tsv').read_text('utf16').splitlines()
            if line.startswith('JOB\t')]
    assert len(rows) == 1
    row = rows[0]
    assert row[10] == 'FAILED' and row[12].startswith('TIMEOUT '), row
    assert row[15] == '1', 'The timed-out job was repeated automatically'
    assert (output / 'timed-out-jobs.tsv').is_file()
    print('PASS: a real blocked COM job was interrupted in its owned process; '
          'checkpoint resumed without repeating the timeout.')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--model', type=Path, required=True)
    parser.add_argument('--work-directory', type=Path, required=True)
    arguments = parser.parse_args()
    check(arguments.model.resolve(strict=True), arguments.work_directory)
