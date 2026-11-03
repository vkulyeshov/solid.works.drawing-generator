"""Check pre-activation source validation without saving or discarding a CAD file."""
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
from vba_sources import read_source

source = read_source()
launcher = (base / 'launcher.vbs').read_text('utf8')
routines = []
for name in ['Setting', 'ReadSettings', 'SnapshotSourceApproved']:
    routine = re.search(r'(?ms)^Private (?:Function|Sub) ' + name + r'\(.*?^End (?:Function|Sub)', source)[0]
    routine = routine.replace('Private ', '', 1)
    routine = re.sub(r' As (?:Object|String|Long|Boolean)\b', '', routine)
    routine = re.sub(r'\b(\w+)\$\(', r'\1(', routine)
    routines.append(routine)
for name in ['RecordSourceValidation', 'ReadUnicode', 'WriteText']:
    routines.append(re.search(r'(?ms)^(?:Function|Sub) ' + name + r'\(.*?^End (?:Function|Sub)', launcher)[0])
constants = []
for name in ['SOURCE_VALIDATION_DIRTY', 'SOURCE_VALIDATION_CLEAN', 'FILE_FOR_READING', 'FILE_UNICODE']:
    value = re.search(r'^Private Const ' + name + r' As \w+ = (.+)$', source, re.M)[1]
    constants.append('Const ' + name + ' = ' + value)
value = re.search(r'^Const SETTINGS_FILENAME = (.+)$', launcher, re.M)[1]
constants.append('Const SETTINGS_FILENAME = ' + value)

# The policy must be evaluated before template validation can rebuild the source.
main = (base / 'modules/main.bas').read_text('utf8')
assert main.index('sourceWasClean = SnapshotSourceApproved') < main.index('    ValidateDrawingTemplate')
worker = re.search(r'(?ms)^Sub WorkerMain\(\).*?^End Sub', launcher)[0]
assert worker.index('inputWasClean = (doc.GetSaveFlag = False)') < worker.index('RecordSourceValidation inputFile')
assert worker.index('RecordSourceValidation inputFile') < worker.index('    doc.Visible = True')

harness = r'''
Dim fs, fso, settings, sessionDir
Sub Check(ByVal condition, ByVal message)
    If Not condition Then
        WScript.Echo message
        WScript.Quit 1
    End If
End Sub

Sub Main()
    Dim path, pid, runtimeIni, approvedBeforeTemplate
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fs = fso
    Set settings = CreateObject("Scripting.Dictionary")
    settings.CompareMode = 1
    sessionDir = fs.GetParentFolderName(WScript.ScriptFullName)
    runtimeIni = fs.BuildPath(sessionDir, SETTINGS_FILENAME)
    path = "C:\Models\Saved.SLDASM"
    pid = "12345"
    Check SnapshotSourceApproved(False,False,path,pid), "A clean direct-macro input was rejected."
    Check Not SnapshotSourceApproved(True,False,path,pid), "An already dirty direct-macro input was accepted."
    Check Not SnapshotSourceApproved(1,0,path,pid), "A COM true value of 1 bypassed the dirty-source guard."
    Check Not SnapshotSourceApproved(-1,0,path,pid), "A VBA true value bypassed the dirty-source guard."
    Check SnapshotSourceApproved(True,True,path,pid), "A separately owned read-only disk input was rejected."

    WriteText runtimeIni, "Font=GOST Type BU" & vbCrLf & "LauncherSourceWasClean=0" & vbCrLf
    RecordSourceValidation path, pid, True
    ReadSettings runtimeIni
    Check Setting("Font", "") = "GOST Type BU", "Recording validation destroyed existing settings."
    Check SnapshotSourceApproved(True,False,path,pid), "Rebuild changes invalidated the clean launcher check."
    Check SnapshotSourceApproved(True,False,LCase(path),pid), "Windows path case invalidated the launcher check."
    Check Not SnapshotSourceApproved(True,False,path,"54321"), "A validation from another process was trusted."
    Check Not SnapshotSourceApproved(True,False,"C:\Models\Other.SLDASM",pid), "A validation from another model was trusted."

    approvedBeforeTemplate = SnapshotSourceApproved(False,False,path,pid)
    Check approvedBeforeTemplate, "Template validation changed the captured pre-operation approval."
    RecordSourceValidation path, pid, False
    ReadSettings runtimeIni
    Check Setting("LauncherSourceWasClean", "") = "0", "A later dirty preflight retained a stale clean flag."
    Check Not SnapshotSourceApproved(True,False,path,pid), "An existing unsaved document passed the guard."
    WScript.Echo "PASS: clean preflight survives rebuilds; dirty inputs blocked; path/PID validation; saved-disk background input"
End Sub
Main
'''

root = work_path('source_validation')
root.mkdir(parents=True, exist_ok=True)
with tempfile.TemporaryDirectory(dir=root) as folder:
    script = Path(folder) / 'source_validation.vbs'
    script.write_text('Option Explicit\n' + '\n'.join(constants + routines) + '\n' + harness, encoding='ascii')
    result = subprocess.run([str(Path(os.environ['WINDIR']) / 'System32/cscript.exe'), '//nologo', str(script)],
                            capture_output=True, timeout=20)
    assert result.returncode == 0 and b'PASS:' in result.stdout, (result.stdout, result.stderr)
    print(result.stdout.decode('ascii').strip())
