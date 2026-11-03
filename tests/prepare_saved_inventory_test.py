"""Clone a saved Models snapshot and build the real inventory for native checks."""
from pathlib import Path
import argparse
import shutil
import sys

sys.dont_write_bytecode = True
base = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(base / 'tools'))
from project_paths import external_path
from rebuild_native import rebuild
from vba_sources import read_source

inspection = r'''
Public Sub InspectInventory()
    Dim key As Variant, job As Object, doc As Object, stream As Object
    Dim code As Long, message As String
    On Error GoTo Failed
    Set sw = Application.SldWorks
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set settings = CreateObject("Scripting.Dictionary")
    settings.CompareMode = vbTextCompare
    ReadSettings sw.GetCurrentMacroPathFolder & SETTINGS_RELATIVE_PATH
    runDir = Setting("LauncherOutput", "")
    progressPath = Setting("LauncherProgress", "")
    prefix = Setting("SnapshotPrefix", "")
    Set quantities = CreateObject("Scripting.Dictionary")
    quantities.CompareMode = vbTextCompare
    Set rootModel = sw.ActiveDoc
    If rootModel Is Nothing Then
        Err.Raise vbObjectError + ERR_SOURCE_NOT_OPEN, , "The copied test assembly is not open."
    End If
    If rootModel.GetType <> SW_DOC_ASSEMBLY Then
        Err.Raise vbObjectError + ERR_SOURCE_TYPE, , "The copied inventory input must be an assembly."
    End If
    InventoryAllSavedComponents rootModel, rootModel.ConfigurationManager.ActiveConfiguration.Name
    Set stream = fs.CreateTextFile(runDir & "\inventory.tsv", True, True)
    stream.WriteLine "Path" & vbTab & "Configuration" & vbTab & "Quantity"
    For Each key In quantities.Keys
        Set job = quantities(key)
        Set doc = job("doc")
        stream.WriteLine doc.GetPathName & vbTab & job("cfg") & vbTab & job("qty")
    Next key
    stream.Close
    ReportProgress STATE_COMPLETE, "Inventory configurations=" & quantities.Count
    Exit Sub
Failed:
    code = Err.Number
    message = Err.Description
    ReportProgress STATE_FAILED, "Inventory error " & code & ": " & message
End Sub
'''

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--snapshot-model', type=Path, required=True)
parser.add_argument('--prefix', required=True)
parser.add_argument('--work-directory', type=Path, required=True)
parser.add_argument('--reuse', action='store_true')
args = parser.parse_args()
model = args.snapshot_model.resolve(strict=True)
directory = external_path(args.work_directory)
if directory.exists() and not args.reuse:
    raise ValueError('Use a fresh external inventory-test directory')
if not args.prefix or not model.name.casefold().startswith(args.prefix.casefold()):
    raise ValueError('The snapshot model must have the supplied prefix')
models = directory / 'Models'
models.mkdir(parents=True, exist_ok=args.reuse)
paths = [path for path in model.parent.iterdir() if path.is_file() and path.name.casefold().startswith(args.prefix.casefold())]
if args.reuse:
    if not (models / model.name).is_file() or not (directory / 'generator.ini').is_file():
        raise ValueError('Reuse requires an existing inventory-test snapshot')
else:
    for path in paths:
        shutil.copy2(path, models / path.name)
(directory / 'generator.ini').write_text('LauncherOutput=' + str(directory) + '\nLauncherProgress=' +
    str(directory / 'progress.tsv') + '\nSnapshotPrefix=' + args.prefix + '\n', encoding='utf-16')
test_source = directory / 'inventory-test.bas'
test_source.write_text(read_source() + '\n' + inspection, encoding='utf8')
rebuild(test_source, directory / 'generator.swp')
print('Copied snapshot files:', len(paths), '; fixture:', directory)
