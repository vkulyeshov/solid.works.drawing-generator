"""Build an external harness that runs the canonical native leader postprocessor."""
from pathlib import Path
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from project_paths import REPOSITORY, work_path
from vba_sources import read_source
from rebuild_native import rebuild

directory = work_path('hub_leaders')
directory.mkdir(parents=True, exist_ok=True)
source = read_source()
source += '''
Public Sub NativeLayoutTest()
    Dim d As SldWorks.ModelDoc2, result As Object, control As Object
    Dim expectedPid As Long, expectedPath As String
    On Error GoTo Failed
    Set sw = Application.SldWorks
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set settings = CreateObject("Scripting.Dictionary")
    runDir = sw.GetCurrentMacroPathFolder & "\\"
    Set d = sw.ActiveDoc
    Set control = fs.OpenTextFile(runDir & "postprocess-control.txt", FILE_FOR_READING, False, FILE_UNICODE)
    expectedPid = CLng(control.ReadLine)
    expectedPath = control.ReadLine
    control.Close
    If sw.GetProcessID <> expectedPid Or StrComp(d.GetPathName, expectedPath, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "The test does not own the active drawing."
    End If
    NormalizeDimensions d
    FinishNativeDimensionLayout d
    Set result = fs.CreateTextFile(runDir & "postprocess-result.txt", True)
    result.WriteLine "OK"
    result.Close
    Exit Sub
Failed:
    Set result = fs.CreateTextFile(runDir & "postprocess-result.txt", True)
    result.WriteLine "ERROR " & Err.Number & ": " & Err.Description
    result.Close
End Sub
'''
fixture = directory / 'postprocess.bas'
fixture.write_text(source, encoding='utf-8')
rebuild(fixture, directory / 'postprocess.swp')
