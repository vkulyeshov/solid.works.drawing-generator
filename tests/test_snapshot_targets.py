"""Exercise the VBA snapshot target policy under WSH without starting CAD."""
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
routine = re.search(r'(?ms)^Private Function ConfigureSnapshotTargets\(.*?^End Function', source)[0]
routine = routine.replace('Private Function ', 'Function ')
routine = re.sub(r' As (?:Object|Variant|Long|String)\b', '', routine)
# VBScript omits the VBA loop variable after Next; keep the loop body unchanged.
routine = re.sub(r'(?m)^(\s*)Next \w+\s*$', r'\1Next', routine)
constants = []
for name in ['MODELS_RELATIVE_PATH', 'ERR_PACK_TARGETS']:
    value = re.search(r'^Private Const ' + name + r' As \w+ = (.+)$', source, re.M)[1]
    constants.append('Const ' + name + ' = ' + value)

harness = r'''
Dim fs, runDir, prefix
Class FakePackAndGo
    Public AddPrefix, AddSuffix, Mode
    Private folder
    Public Function SetSaveToName(ByVal override, ByVal path)
        folder = path
        SetSaveToName = (override And Mode <> "folder_rejected")
    End Function
    Public Function GetDocumentSaveToNames(ByRef targets, ByRef states)
        targets = Array(folder & "\" & AddPrefix & "assembly.sldasm", _
                        folder & "\" & AddPrefix & "part.sldprt", _
                        folder & "\" & AddPrefix & "fan.step")
        states = Array(0, 0, 0)
        Select Case Mode
            Case "native_only", "count_mismatch"
                targets = Array(folder & "\" & AddPrefix & "assembly.sldasm", folder & "\" & AddPrefix & "part.sldprt")
            Case "no_array"
                targets = Empty
            Case "excluded_step"
                targets(2) = ""
            Case "outside_models"
                targets(2) = fs.BuildPath(runDir, "fan.step")
            Case "duplicate"
                targets(1) = targets(0)
            Case "changed_type"
                targets(2) = folder & "\fan.sldprt"
        End Select
        GetDocumentSaveToNames = (Mode <> "names_rejected")
    End Function
    Public Function SetDocumentSaveToNames(ByVal names)
        Err.Raise vbObjectError + 999, , "Linked CAD files reject manual target arrays."
    End Function
End Class

Sub Log(ByVal message)
    WScript.Echo message
End Sub

Sub Main()
    Dim pg, names, result, number, description, mode
    Set fs = CreateObject("Scripting.FileSystemObject")
    runDir = fs.GetParentFolderName(WScript.ScriptFullName)
    prefix = "AD_TEST_"
    mode = WScript.Arguments(0)
    Set pg = New FakePackAndGo
    pg.Mode = mode
    names = Array("C:\source\assembly.sldasm", "C:\source\part.sldprt", "C:\source\fan.step")
    If mode = "native_only" Then
        names = Array("C:\source\assembly.sldasm", "C:\source\part.sldprt")
    ElseIf mode = "duplicate" Then
        names = Array("C:\source\assembly.sldasm", "C:\source\other.sldasm", "C:\source\fan.step")
    End If
    On Error Resume Next
    result = ConfigureSnapshotTargets(pg, names)
    number = Err.Number
    description = Err.Description
    On Error GoTo 0
    If mode = "linked_step" Or mode = "native_only" Then
        If number <> 0 Then
            WScript.Echo description
            WScript.Quit 1
        End If
        If pg.AddPrefix <> prefix Or pg.AddSuffix <> "" Then
            WScript.Quit 1
        End If
        If mode = "linked_step" Then
            If result(2) <> fs.BuildPath(runDir & MODELS_RELATIVE_PATH, prefix & "fan.step") Then
                WScript.Quit 1
            End If
        End If
        WScript.Echo "PASS " & mode
    Else
        If number <> vbObjectError + ERR_PACK_TARGETS Then
            WScript.Echo "Unexpected error " & number & ": " & description
            WScript.Quit 1
        End If
        WScript.Echo "REJECT " & mode & ": " & description
    End If
End Sub

Main
'''

root = work_path('snapshot_targets')
root.mkdir(parents=True, exist_ok=True)
cases = ['linked_step', 'native_only', 'folder_rejected', 'names_rejected', 'no_array',
         'count_mismatch', 'excluded_step', 'outside_models', 'duplicate', 'changed_type']
with tempfile.TemporaryDirectory(dir=root) as folder:
    script = Path(folder) / 'targets.vbs'
    script.write_text('Option Explicit\n' + '\n'.join(constants) + '\n' + routine + '\n' + harness,
                      encoding='ascii')
    for case in cases:
        run = subprocess.run([str(Path(os.environ['WINDIR']) / 'System32/cscript.exe'),
                              '//nologo', str(script), case], capture_output=True, timeout=15)
        assert run.returncode == 0, (case, run.stdout, run.stderr)
        assert ('PASS ' if case in ['linked_step', 'native_only'] else 'REJECT ').encode() in run.stdout
print('Snapshot target checks passed:', len(cases))
