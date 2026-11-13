"""Rebuild only the native layout pass against a previously copied drawing fixture."""
from pathlib import Path
import argparse
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from project_paths import external_path
from vba_sources import read_source
from rebuild_native import rebuild

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--fixture', type=Path, required=True)
args = parser.parse_args()
fixture = external_path(args.fixture)
assert list((fixture / 'output/Drawings').glob('*.SLDDRW')), 'An isolated generated drawing is required'
driver = '''
Public Sub main()
    Dim d As SldWorks.ModelDoc2, file As Object, er As Long, wr As Long, sheet As Object, props As Variant
    Dim number As Long, message As String, pdf As Object
    On Error GoTo Failed
    Set sw = Application.SldWorks
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set settings = CreateObject("Scripting.Dictionary")
    settings.CompareMode = vbTextCompare
    ReadSettings sw.GetCurrentMacroPathFolder & SETTINGS_RELATIVE_PATH
    runDir = Setting("LauncherOutput", "")
    progressPath = Setting("LauncherProgress", "")
    settings("OrthographicViews") = "3"
    jobNumber = 1: plannedJobs = 1
    For Each file In fs.GetFolder(runDir & DRAWINGS_RELATIVE_PATH).Files
        If LCase$(fs.GetExtensionName(file.Path)) = "slddrw" Then
            Set d = sw.OpenDoc6(file.Path, SW_DOC_DRAWING, SW_OPEN_SILENT, "", er, wr)
            Exit For
        End If
    Next file
    If d Is Nothing Then
        Err.Raise vbObjectError + ERR_REOPEN_DRAWING, , "The isolated replay drawing could not be opened."
    End If
    PrepareDocumentWindow d
    Set sheet = AsDrawing(d).GetCurrentSheet
    props = sheet.GetProperties
    sheetWidth = props(SHEET_WIDTH_PROPERTY): sheetHeight = props(SHEET_HEIGHT_PROPERTY)
    FinishMemoryDimensionLayout d
    If d.Save3(SW_SAVE_SILENT, er, wr) = False Then
        Err.Raise vbObjectError + ERR_SAVE_DRAWING, , "Cannot save the isolated replay drawing."
    End If
    Set pdf = sw.GetExportFileData(SW_EXPORT_PDF)
    pdf.ViewPdfAfterSaving = False
    pdf.SetSheets SW_EXPORT_ALL_SHEETS, Empty
    If d.Extension.SaveAs(fs.BuildPath(runDir, "memory-replay.pdf"), SW_SAVE_CURRENT_VERSION, SW_SAVE_SILENT, pdf, er, wr) = False Then
        Err.Raise vbObjectError + ERR_SAVE_PDF, , "Cannot export the isolated replay PDF."
    End If
    ReportProgress STATE_COMPLETE, "Isolated memory layout completed"
    Exit Sub
Failed:
    number = Err.Number: message = Err.Description
    On Error Resume Next
    If Not d Is Nothing Then
        d.Save3 SW_SAVE_SILENT, er, wr
    End If
    ReportProgress STATE_FAILED, "Memory replay error " & number & ": " & message
End Sub
'''
source = read_source().replace('Public Sub main()', 'Private Sub GeneratorMain()', 1) + '\n' + driver
path = fixture / 'memory-replay.bas'
path.write_text(source, encoding='utf8')
rebuild(path, fixture / 'generator.swp')
