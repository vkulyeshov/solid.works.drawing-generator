"""Regress native dirty flags before and after source validation on isolated copies."""
from pathlib import Path
import hashlib
import os
import subprocess
import sys
import tempfile

sys.dont_write_bytecode = True
base = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(base / 'tools'))
from project_paths import work_path
from rebuild_native import rebuild
from vba_sources import read_source
from prepare_model_test import prepare

model = base.parent / 'model/RightStandFrontCockpit.SLDPRT'
original_hash = hashlib.sha256(model.read_bytes()).hexdigest()
root = work_path('source_validation_native')
root.mkdir(parents=True, exist_ok=True)
holder = Path(tempfile.mkdtemp(prefix='run_', dir=root))
host = Path(os.environ['WINDIR']) / 'System32/WindowsPowerShell/v1.0/powershell.exe'
source = read_source()
dirty_flag = '''    original.SetReadOnlyState False
    original.Visible = True
    original.SetSaveFlag
    If original.GetSaveFlag = False Then
        Err.Raise vbObjectError + ERR_GENERATION_DATA, , "Cannot set a visible dirty flag in the native test copy."
    End If
'''

for case in ['dirty_after_validation', 'dirty_before_validation']:
    directory = holder / case
    prepare(model, directory)
    settings_file = directory / 'generator.ini'
    settings = settings_file.read_text('utf-16').replace('AutoDimensions=1', 'AutoDimensions=0')
    settings = settings.replace('NativeDimensions=1', 'NativeDimensions=0')
    settings_file.write_text(settings, encoding='utf-16')
    if case == 'dirty_after_validation':
        marker = '    ValidateDrawingTemplate\n'
        replacement = marker + dirty_flag
    else:
        marker = '    sourceWasClean = SnapshotSourceApproved'
        replacement = dirty_flag + marker
    assert source.count(marker) == 1
    injected = directory / 'dirty-flag-test.bas'
    injected.write_text(source.replace(marker, replacement), encoding='utf8')
    rebuild(injected, directory / 'generator.swp')
    input_file = directory / 'input' / model.name
    input_hash = hashlib.sha256(input_file.read_bytes()).hexdigest()
    result = subprocess.run([str(host), '-NoProfile', '-ExecutionPolicy', 'Bypass',
                             '-File', str(base / 'tests/RunSixViewGeneration.ps1'),
                             '-WorkDirectory', str(directory), '-ModelFileName', model.name],
                            cwd=base, capture_output=True, timeout=600)
    (directory / 'host-output.txt').write_bytes(result.stdout + result.stderr)
    assert hashlib.sha256(input_file.read_bytes()).hexdigest() == input_hash, 'The dirty source copy was saved'
    assert hashlib.sha256(model.read_bytes()).hexdigest() == original_hash, 'The original model changed'
    events = (directory / 'progress.tsv').read_text('utf-16')
    if case == 'dirty_after_validation':
        assert result.returncode == 0 and '\tCOMPLETE\t' in events, (result.stdout, result.stderr)
        log = (directory / 'output/run.log').read_text('utf-16')
        assert 'SOURCE validated before activation' in log
    else:
        assert result.returncode != 0 and 'Source already had unsaved changes before generation' in events
        assert not (directory / 'output/generation-state.tsv').exists()
    print('PASS:', case, '; source bytes unchanged; fixture:', directory, flush=True)
