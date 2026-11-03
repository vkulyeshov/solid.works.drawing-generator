"""Exercise saved-reference traversal, including unloaded suppressed subassemblies."""
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
functions = []
names = ['CopiedInventoryPath', 'RememberLoadedInventoryModel', 'OpenInventoryModel', 'RememberInventoryModel',
         'SavedInventoryReferences', 'InventorySavedAssembly', 'InventoryAllSavedComponents']
for name in names:
    body = re.search(r'(?ms)^Private (?:Function|Sub) ' + name + r'\(.*?^End (?:Function|Sub)', source)[0]
    body = body.replace('Private ', '', 1)
    body = body.replace(' As New Collection', '')
    body = re.sub(r' As (?:Object|Collection|String|Long|Variant)\b', '', body)
    body = body.replace('LCase$(', 'LCase(')
    body = re.sub(r'(?m)^(\s*)Next \w+\s*$', r'\1Next', body)
    if name == 'SavedInventoryReferences':
        body = body.split('InventoryFailed:')[0] + 'End Function'
        body = body.replace('    On Error GoTo InventoryFailed\n', '')
        body = body.replace('    key = ', '    Set references = CreateObject("Scripting.Dictionary")\n    key = ', 1)
        body = body.replace('references.Add reference', 'references.Add references.Count + 1, reference')
    if name == 'OpenInventoryModel':
        body = body.replace('    If models.Exists(path) Then', '    Set doc = Nothing\n    If models.Exists(path) Then', 1)
    body = body.replace('For Each reference In references', 'For Each reference In references.Items')
    functions.append(body)
constants = []
for name in ['MODELS_RELATIVE_DIRECTORY', 'ASSEMBLY_EXTENSION', 'PART_EXTENSION', 'SW_DOC_ASSEMBLY',
             'SW_DOC_PART', 'SW_OPEN_SILENT', 'ERR_INVENTORY_REFERENCE', 'ERR_CONFIGURATION']:
    value = re.search(r'^Private Const ' + name + r' As \w+ = (.+)$', source, re.M)[1]
    constants.append('Const ' + name + ' = ' + value)

harness = r'''
Dim fs, sw, runDir, prefix, quantities
Class FakeConfiguration
    Public Name
End Class
Class FakeConfigurationManager
    Public ActiveConfiguration
End Class
Class FakeComponent
    Public Path, ReferencedConfiguration, LoadedDocument, MissingModelError
    Private Sub Class_Initialize()
        Set LoadedDocument = Nothing
        MissingModelError = False
    End Sub
    Public Function GetPathName()
        GetPathName = Path
    End Function
    Public Function GetModelDoc2()
        If MissingModelError Then
            Err.Raise 1001, , "This unloaded reference has no optional model pointer."
        End If
        Set GetModelDoc2 = LoadedDocument
    End Function
    Public Function GetSuppression2()
        Err.Raise 1002, , "The disabled filter inspected suppression."
    End Function
    Public Function SetSuppression2(ByVal state)
        Err.Raise 1003, , "Saved-reference inventory changed suppression."
    End Function
End Class
Class FakeDocument
    Public Path, DocumentType, Configurations, ConfigurationManager, Enumerations
    Private Sub Class_Initialize()
        Set Configurations = CreateObject("Scripting.Dictionary")
        Set ConfigurationManager = New FakeConfigurationManager
        Set ConfigurationManager.ActiveConfiguration = New FakeConfiguration
        ConfigurationManager.ActiveConfiguration.Name = "Default"
        Enumerations = 0
    End Sub
    Public Function GetPathName()
        GetPathName = Path
    End Function
    Public Function GetConfigurationByName(ByVal name)
        Dim result
        Set result = Nothing
        If Configurations.Exists(name) Then
            Set result = New FakeConfiguration
            result.Name = name
        End If
        Set GetConfigurationByName = result
    End Function
    Public Function ShowConfiguration2(ByVal name)
        ShowConfiguration2 = Configurations.Exists(name)
        If ShowConfiguration2 Then
            ConfigurationManager.ActiveConfiguration.Name = name
        End If
    End Function
    Public Function GetComponents(ByVal topLevelOnly)
        Check topLevelOnly, "Saved inventory counted descendants twice through a flat array."
        Enumerations = Enumerations + 1
        GetComponents = Configurations(ConfigurationManager.ActiveConfiguration.Name)
    End Function
End Class
Class FakeSolidWorks
    Public Models, Loaded, Opens
    Private Sub Class_Initialize()
        Set Models = CreateObject("Scripting.Dictionary")
        Models.CompareMode = 1
        Set Loaded = CreateObject("Scripting.Dictionary")
        Loaded.CompareMode = 1
        Opens = 0
    End Sub
    Public Function GetOpenDocumentByName(ByVal path)
        Set GetOpenDocumentByName = Nothing
        If Loaded.Exists(path) Then
            Set GetOpenDocumentByName = Loaded(path)
        End If
    End Function
    Public Function OpenDoc6(ByVal path, ByVal kind, ByVal options, ByVal cfg, ByRef error, ByRef warning)
        Set OpenDoc6 = Nothing
        If Models.Exists(path) Then
            Check kind = Models(path).DocumentType, "The copied reference used the wrong document type."
            Check options = SW_OPEN_SILENT, "An unexpected open mode was used."
            Set OpenDoc6 = Models(path)
            Loaded.Add path, Models(path)
            Opens = Opens + 1
        Else
            error = 1
        End If
    End Function
End Class
Function AsAssembly(ByVal doc)
    Set AsAssembly = doc
End Function
Sub Log(ByVal message)
End Sub
Sub Check(ByVal condition, ByVal message)
    If Not condition Then
        WScript.Echo message
        WScript.Quit 1
    End If
End Sub
Function Component(ByVal file, ByVal cfg)
    Dim result
    Set result = New FakeComponent
    result.Path = file
    result.ReferencedConfiguration = cfg
    Set Component = result
End Function
Function Model(ByVal name, ByVal kind)
    Dim result, stream
    Set result = New FakeDocument
    result.Path = runDir & MODELS_RELATIVE_DIRECTORY & prefix & name
    result.DocumentType = kind
    Set stream = fs.CreateTextFile(result.Path)
    stream.Close
    sw.Models.Add result.Path, result
    Set Model = result
End Function
Sub ResetQuantities()
    Set quantities = CreateObject("Scripting.Dictionary")
    quantities.CompareMode = 1
End Sub
Sub CheckQuantity(ByVal name, ByVal cfg, ByVal count)
    Dim key, job
    key = LCase(runDir & MODELS_RELATIVE_DIRECTORY & prefix & name & "|" & cfg)
    Check quantities.Exists(key), "A saved reference was omitted: " & key
    Set job = quantities(key)
    Check job("qty") = count, "Wrong nested occurrence quantity: " & key
End Sub
Sub Main()
    Dim root, branch, child, sharedPart, leaf, originalPath, number, cachedReference, missingReference
    Set fs = CreateObject("Scripting.FileSystemObject")
    runDir = fs.GetParentFolderName(WScript.ScriptFullName)
    prefix = "TEST_"
    fs.CreateFolder runDir & "\Models"
    Set sw = New FakeSolidWorks
    Set root = Model("root.sldasm", SW_DOC_ASSEMBLY)
    Set branch = Model("branch.sldasm", SW_DOC_ASSEMBLY)
    Set child = Model("child.sldasm", SW_DOC_ASSEMBLY)
    Set sharedPart = Model("shared.sldprt", SW_DOC_PART)
    Set leaf = Model("leaf.sldprt", SW_DOC_PART)
    Set cachedReference = Component("C:\old-snapshot\TEST_shared.sldprt", "Default")
    Set cachedReference.LoadedDocument = sharedPart
    Set missingReference = Component("C:\source\leaf.sldprt", "Default")
    missingReference.MissingModelError = True
    root.Configurations.Add "Default", Array(Component("C:\source\branch.sldasm", "A"), _
        Component("C:\source\branch.sldasm", "A"), Component("C:\source\branch.sldasm", "B"), _
        cachedReference)
    branch.Configurations.Add "A", Array(Component("C:\source\shared.sldprt", "Default"), Component("C:\source\child.sldasm", "Default"))
    branch.Configurations.Add "B", Array(Component("C:\source\shared.sldprt", "Alternate"), Component("C:\source\leaf.sldprt", "Default"))
    child.Configurations.Add "Default", Array(missingReference)
    sw.Loaded.Add root.Path, root
    ResetQuantities
    InventoryAllSavedComponents root, "Default"
    Check quantities.Count = 6, "Saved model/configuration grouping changed."
    CheckQuantity "branch.sldasm", "A", 2
    CheckQuantity "branch.sldasm", "B", 1
    CheckQuantity "child.sldasm", "Default", 2
    CheckQuantity "shared.sldprt", "Default", 3
    CheckQuantity "shared.sldprt", "Alternate", 1
    CheckQuantity "leaf.sldprt", "Default", 3
    Check root.Enumerations = 1 And branch.Enumerations = 2 And child.Enumerations = 1, "Assembly/configuration plans were not cached independently."
    Check sw.Opens = 3, "A loaded component was reopened instead of using its validated model pointer."
    Check root.ConfigurationManager.ActiveConfiguration.Name = "Default", "The root configuration changed."

    originalPath = root.Path
    root.Path = runDir & "\source.sldasm"
    On Error Resume Next
    InventoryAllSavedComponents root, "Default"
    number = Err.Number
    Err.Clear
    On Error GoTo 0
    Check number = vbObjectError + ERR_INVENTORY_REFERENCE, "A source outside Models was accepted."
    root.Path = originalPath

    root.Configurations("Default") = Array(Component("C:\source\missing.sldprt", "Default"))
    ResetQuantities
    On Error Resume Next
    InventoryAllSavedComponents root, "Default"
    number = Err.Number
    Err.Clear
    On Error GoTo 0
    Check number = vbObjectError + ERR_INVENTORY_REFERENCE, "A missing copied reference was silently omitted."

    root.Configurations("Default") = Array(Component(root.Path, "Default"))
    ResetQuantities
    On Error Resume Next
    InventoryAllSavedComponents root, "Default"
    number = Err.Number
    Err.Clear
    On Error GoTo 0
    Check number = vbObjectError + ERR_INVENTORY_REFERENCE, "A cyclic assembly did not stop recursion."

    root.Configurations("Default") = Array(Component(branch.Path, "Missing"))
    ResetQuantities
    On Error Resume Next
    InventoryAllSavedComponents root, "Default"
    number = Err.Number
    Err.Clear
    On Error GoTo 0
    Check number = vbObjectError + ERR_CONFIGURATION, "A missing referenced configuration was ignored."
    WScript.Echo "PASS: unloaded saved references; suppressed descendants; repeated quantities; configuration caches; copied paths; cycles"
End Sub
Main
'''

root = work_path('inventory_saved')
root.mkdir(parents=True, exist_ok=True)
with tempfile.TemporaryDirectory(dir=root) as folder:
    script = Path(folder) / 'inventory.vbs'
    script.write_text('Option Explicit\n' + '\n'.join(constants + functions) + '\n' + harness, encoding='ascii')
    result = subprocess.run([str(Path(os.environ['WINDIR']) / 'System32/cscript.exe'), '//nologo', str(script)],
                            capture_output=True, timeout=20)
    if result.returncode != 0 or b'PASS:' not in result.stdout:
        diagnostic = result.stderr.decode('oem', errors='replace')
        location = re.search(r'\((\d+),\s*\d+\)', diagnostic)
        if location:
            diagnostic += '\n' + script.read_text('ascii').splitlines()[int(location[1]) - 1]
        raise AssertionError((result.stdout, diagnostic))
    print(result.stdout.decode('ascii').strip())
