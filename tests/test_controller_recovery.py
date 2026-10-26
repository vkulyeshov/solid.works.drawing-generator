"""Exercise the real WSH controller with deterministic worker fault injection.

Only WorkerMain is replaced in an external launcher copy. No CAD process starts.
"""
from pathlib import Path
import sys
sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from project_paths import work_path
import json
import os
import re
import subprocess

base = Path(__file__).resolve().parents[1]
fixture = work_path('controller_recovery')
fixture.mkdir(parents=True, exist_ok=True)
source = (base / 'launcher.vbs').read_text('utf8')
worker = '''Sub WorkerMain()
    Dim context, marker, count, mode, cells, header
    Set context = ReadKeyValues(fso.BuildPath(sessionDir, CONTEXT_FILENAME))
    outputDir = context("Output")
    progressFile = fso.BuildPath(sessionDir, PROGRESS_FILENAME)
    marker = fso.BuildPath(outputDir, "invocations.txt")
    count = 0
    If fso.FileExists(marker) Then
        count = CLng(ReadUnicode(marker))
    End If
    count = count + 1
    WriteText marker, CStr(count)
    mode = fso.GetFileName(outputDir)
    If mode = "validation" Then
        WriteWorkerResult EXIT_GENERATION_FAILURE, "Injected validation failure"
        exitCode = EXIT_GENERATION_FAILURE
    ElseIf mode = "startup_crash" Then
        exitCode = EXIT_LAUNCHER_FAILURE
    ElseIf count <= 3 Then
        cells = Split(Split(ReadUnicode(fso.BuildPath(outputDir, CHECKPOINT_FILENAME)), vbCrLf)(1), vbTab)
        cells(CHECKPOINT_ATTEMPTS_FIELD) = CStr(count)
        cells(CHECKPOINT_STATUS_FIELD) = STATE_RUNNING
        header = "STATE" & vbTab & "3"
        WriteText fso.BuildPath(outputDir, CHECKPOINT_FILENAME), header & vbCrLf & Join(cells, vbTab) & vbCrLf
        exitCode = EXIT_LAUNCHER_FAILURE
    Else
        WriteText progressFile, Time & vbTab & STATE_COMPLETE & vbTab & "1" & vbTab & "1" & vbTab & "1" & vbTab & "0" & vbTab & "0" & vbTab & outputDir & vbTab & "Injected completion" & vbCrLf
        WriteWorkerResult EXIT_SUCCESS, "Injected completion"
        exitCode = EXIT_SUCCESS
    End If
End Sub'''
modified, replacements = re.subn(r'(?ms)^Sub WorkerMain\(\).*?^End Sub', lambda _: worker, source)
assert replacements == 1
launcher = fixture / 'launcher.vbs'
launcher.write_text(modified, encoding='utf8')
(fixture / 'generator.swp').write_bytes(b'Unused worker fixture')
(fixture / 'input.sldprt').write_bytes(b'Unused input fixture')
ini = fixture / 'generator.ini'
ini.write_text('PaperFormat=A4\r\nDrawingTemplate=' + str(base / 'templates/A4.landscape.drwdot') + '\r\n', encoding='utf-16')
results = []
for name, expected_code, expected_invocations in [('validation', 2, 1), ('startup_crash', 1, 3), ('attempt_progress', 0, 4)]:
    out = fixture / name
    out.mkdir(exist_ok=True)
    marker = out / 'invocations.txt'
    if marker.exists():
        marker.unlink()
    row = ['JOB', '1', 'unused', 'Default', 'body', '', '1', 'False', 'False', '1', 'PENDING', '001_fixture', '', '0', '0', '0']
    (out / 'generation-state.tsv').write_bytes(('STATE\t3\r\n' + '\t'.join(row) + '\r\n').encode('utf-16'))
    command = [str(Path(os.environ['WINDIR']) / 'System32/cscript.exe'), '//nologo', str(launcher),
               '/input:' + str(fixture / 'input.sldprt'), '/out:' + str(out)]
    run = subprocess.run(command, capture_output=True, timeout=30)
    assert run.returncode == expected_code, (name, run.returncode, run.stdout, run.stderr)
    assert int(marker.read_text('utf-16')) == expected_invocations, (name, run.stdout)
    results.append({'test': name, 'passed': True, 'invocations': expected_invocations})
work_path('controller-recovery.json').write_text(json.dumps(results, indent=2), encoding='utf8')
print('Controller recovery checks passed:', len(results))
