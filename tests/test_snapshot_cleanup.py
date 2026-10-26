"""Exercise the generator's backup cleanup with real Windows file operations.

Run the filesystem-only VBA routine under WSH in disposable external folders;
no SolidWorks process or production artifact is touched.
"""
from pathlib import Path
import os
import re
import subprocess
import sys
import tempfile

sys.dont_write_bytecode = True
base = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(base / 'tools'))
from project_paths import work_path

source = (base / 'generator.bas').read_text('utf8')
routine = re.search(r'(?ms)^Private Sub CleanupSnapshotBackup\(\).*?^End Sub', source)[0]
routine = routine.replace('Private Sub ', 'Sub ').replace(' As String', '')
# WSH has no labelled VBA exception handlers. Keep the filesystem operations
# unchanged, map the explicit path rejection to warning/return, and let any
# unexpected filesystem error fail the harness instead of being swallowed.
routine = routine.split('CleanupFailed:')[0] + 'End Sub'
routine = routine.replace('    On Error GoTo CleanupFailed\n', '')
routine = re.sub(r'(?m)^(\s*)Err\.Raise[^\n]+$',
                 r'\1Warn "Invalid snapshot backup cleanup path."\n\1Exit Sub', routine)
constant_names = ['SOURCE_BACKUP_RELATIVE_DIRECTORY', 'SOURCE_BACKUP_INDEX_RELATIVE_PATH', 'ERR_GENERATION_DATA']
constants = []
for name in constant_names:
    declaration = re.search(r'^Private Const ' + name + r' As \w+ = (.+)$', source, re.M)[1]
    constants.append('Const ' + name + ' = ' + declaration)
fixture = work_path('snapshot_cleanup')
fixture.mkdir(parents=True, exist_ok=True)
quote = lambda value: '"' + str(value).replace('"', '""') + '"'
host = Path(os.environ['WINDIR']) / 'System32/cscript.exe'
for case in ['successful_snapshot', 'older_failed_backup', 'invalid_cleanup_path']:
    with tempfile.TemporaryDirectory(prefix=case + '_', dir=fixture) as folder:
        output = Path(folder) / 'out'
        current = output / 'SourceBackup/AD_current'
        current.mkdir(parents=True)
        (current / 'source.sldasm').write_bytes(b'Backup fixture')
        index = output / 'source-backup.tsv'
        index.write_bytes(b'Index fixture')
        prefix = 'AD_current'
        if case == 'older_failed_backup':
            older = output / 'SourceBackup/AD_failed'
            older.mkdir()
            (older / 'keep.sldprt').write_bytes(b'Preserve older backup')
        elif case == 'invalid_cleanup_path':
            outside = output / 'outside'
            outside.mkdir()
            (outside / 'keep.txt').write_bytes(b'Preserve unrelated files')
            prefix = '..\\outside'
        script = Path(folder) / 'cleanup.vbs'
        harness = '\n'.join(constants) + '\n' + '''Dim fs, runDir, prefix, warnings
Set fs = CreateObject("Scripting.FileSystemObject")
warnings = 0
''' + 'runDir = ' + quote(output) + '\nprefix = ' + quote(prefix) + '''
CleanupSnapshotBackup
WScript.Echo "Warnings=" & warnings
Sub Log(ByVal message)
    WScript.Echo message
End Sub
Sub Warn(ByVal message)
    warnings = warnings + 1
    WScript.Echo message
End Sub
''' + routine
        script.write_bytes(harness.encode('utf-16'))
        run = subprocess.run([str(host), '//nologo', str(script)], capture_output=True, timeout=15)
        assert run.returncode == 0, (case, run.stdout, run.stderr)
        if case == 'invalid_cleanup_path':
            assert current.exists() and index.exists() and (outside / 'keep.txt').exists()
            assert b'Warnings=1' in run.stdout, run.stdout
        else:
            assert not current.exists() and not index.exists(), run.stdout
            assert b'Warnings=0' in run.stdout, run.stdout
            if case == 'older_failed_backup':
                assert (older / 'keep.sldprt').read_bytes() == b'Preserve older backup'
            else:
                assert not (output / 'SourceBackup').exists()
        print(case + ': passed')
