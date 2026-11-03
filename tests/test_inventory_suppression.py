"""Exercise actual VBA inventory with suppressed ancestors and shared occurrences."""
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
routines = []
for name in ['Setting', 'SuppressionFilterEnabled', 'RememberInventoryModel', 'ComponentOrParentSuppressed', 'Inventory']:
    routine = re.search(r'(?ms)^Private (?:Function|Sub) ' + name + r'\(.*?^End (?:Function|Sub)', source)[0]
    routine = routine.replace('Private ', '', 1)
    routine = re.sub(r' As (?:SldWorks\.ModelDoc2|Object|Variant|Long|String|Boolean)\b', '', routine)
    routine = re.sub(r'(?m)^(\s*)Next \w+\s*$', r'\1Next', routine)
    routine = routine.replace('LCase$(', 'LCase(')
    routines.append(routine)
constants = []
for name in ['SW_COMPONENT_SUPPRESSED',
             'SUPPRESSION_FILTER_DISABLED', 'SUPPRESSION_FILTER_ENABLED', 'MODELS_RELATIVE_DIRECTORY',
             'ERR_LIBRARY_COMPONENT', 'ERR_DRAWING_SETTINGS']:
    value = re.search(r'^Private Const ' + name + r' As \w+ = (.+)$', source, re.M)[1]
    constants.append('Const ' + name + ' = ' + value)

launcher = (base / 'launcher.vbs').read_text('utf8')
for name in ['ResolveSuppressionFilter', 'IniSetting']:
    routines.append(re.search(r'(?ms)^Function ' + name + r'\(.*?^End Function', launcher)[0])
value = re.search(r'^Const ERR_ARGUMENTS = (.+)$', launcher, re.M)[1]
constants.append('Const ERR_ARGUMENTS = ' + value)

harness = r'''
Dim fs, sw, runDir, prefix, quantities, skipped, warnings, settings, excludeSuppressed, named, resolutions

Class FakeDocument
    Public Path
    Public Function GetPathName()
        GetPathName = Path
    End Function
End Class

Class FakeComponent
    Public Name2, ReferencedConfiguration, State, ForbiddenAccess, ModelReads
    Public Parent, Document, ResolutionResult, RemainSuppressed
    Public Function GetSuppression2()
        GetSuppression2 = State
    End Function
    Public Function IsSuppressed()
        Err.Raise 1001, , "The ambiguous lightweight/suppressed flag was used."
    End Function
    Public Function GetParent()
        Set GetParent = Parent
    End Function
    Public Function SetSuppression2(ByVal value)
        Err.Raise 1003, , "The filtered inventory changed component suppression."
    End Function
    Public Function GetModelDoc2()
        If ForbiddenAccess Then
            Err.Raise 1002, , "A suppressed occurrence or descendant was opened: " & Name2
        End If
        ModelReads = ModelReads + 1
        Set GetModelDoc2 = Document
    End Function
End Class

Class FakeAssembly
    Public Components, CollapsedComponents, Gate, Path, Enumerations
    Private Sub Class_Initialize()
        Set Gate = Nothing
        Enumerations = 0
    End Sub
    Public Function GetPathName()
        GetPathName = Path
    End Function
    Public Function GetComponents(ByVal topLevelOnly)
        Check topLevelOnly = False, "Nested occurrences were omitted from the inventory."
        Enumerations = Enumerations + 1
        If Not Gate Is Nothing Then
            If Gate.State = SW_COMPONENT_SUPPRESSED Then
                GetComponents = CollapsedComponents
                Exit Function
            End If
        End If
        GetComponents = Components
    End Function
End Class

Function AsAssembly(ByVal model)
    Set AsAssembly = model
End Function

Sub Log(ByVal message)
    If InStr(message, "skipped suppressed") > 0 Then
        skipped = skipped + 1
    End If
End Sub

Sub Warn(ByVal message)
    warnings = warnings + 1
End Sub

Sub Check(ByVal condition, ByVal message)
    If Not condition Then
        WScript.Echo message
        WScript.Quit 1
    End If
End Sub

Function Component(ByVal name, ByVal file, ByVal configuration, ByVal state, ByVal parent)
    Dim result, doc
    Set result = New FakeComponent
    Set doc = New FakeDocument
    doc.Path = runDir & MODELS_RELATIVE_DIRECTORY & file
    result.Name2 = name
    result.ReferencedConfiguration = configuration
    result.State = state
    result.ForbiddenAccess = False
    result.ModelReads = 0
    Set result.Parent = parent
    Set result.Document = doc
    Set Component = result
End Function

Sub CheckQuantity(ByVal file, ByVal configuration, ByVal expected)
    Dim key, job
    key = LCase(runDir & MODELS_RELATIVE_DIRECTORY & file & "|" & configuration)
    Check quantities.Exists(key), "Missing active file/configuration: " & key
    Set job = quantities(key)
    Check job("qty") = expected, "Suppressed occurrences affected quantity: " & key
    Check job("cfg") = configuration, "A referenced configuration was lost."
End Sub

Sub OptionChecks()
    Dim value, errorNumber, invalid
    Set named = CreateObject("Scripting.Dictionary")
    Check ResolveSuppressionFilter("") = "0", "The launcher enabled filtering by default."
    Check ResolveSuppressionFilter("ExcludeSuppressed=1") = "1", "The launcher ignored the INI filter."
    named.Add "excludesuppressed", "0"
    Check ResolveSuppressionFilter("ExcludeSuppressed=1") = "0", "CLI 0 did not override INI 1."
    named("excludesuppressed") = "1"
    Check ResolveSuppressionFilter("ExcludeSuppressed=0") = "1", "CLI 1 did not override INI 0."
    Check ResolveSuppressionFilter("ExcludeSuppressed=invalid") = "1", "An explicit CLI override did not take priority."
    For Each invalid In Array("", "true", "false", "2", "-1")
        named("excludesuppressed") = invalid
        On Error Resume Next
        value = ResolveSuppressionFilter("")
        errorNumber = Err.Number
        Err.Clear
        On Error GoTo 0
        Check errorNumber = vbObjectError + ERR_ARGUMENTS, "The launcher accepted an invalid suppression filter."
    Next
    Set settings = CreateObject("Scripting.Dictionary")
    Check SuppressionFilterEnabled() = False, "Direct VBA enabled filtering by default."
    settings.Add "ExcludeSuppressed", "1"
    Check SuppressionFilterEnabled(), "Direct VBA ignored the enabled filter."
    settings("ExcludeSuppressed") = "0"
    Check SuppressionFilterEnabled() = False, "Direct VBA ignored the disabled filter."
    settings("ExcludeSuppressed") = "invalid"
    On Error Resume Next
    value = SuppressionFilterEnabled()
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    Check errorNumber = vbObjectError + ERR_DRAWING_SETTINGS, "Direct VBA accepted an invalid filter value."
End Sub

Sub Main()
    Dim assembly, blocked, middle, child, grandchild, direct, active, first, second
    Dim hidden, light, fullLight, alternate, item, emptyAssembly, missing
    Set fs = CreateObject("Scripting.FileSystemObject")
    runDir = fs.GetParentFolderName(WScript.ScriptFullName)
    prefix = "TEST_"
    OptionChecks
    excludeSuppressed = True
    resolutions = 0
    skipped = 0
    warnings = 0
    Set quantities = CreateObject("Scripting.Dictionary")
    quantities.CompareMode = 1
    Set blocked = Component("Blocked-1", "blocked.sldasm", "Default", 0, Nothing)
    Set middle = Component("Blocked-1/Middle-1", "middle.sldasm", "Default", 2, blocked)
    Set child = Component("Blocked-1/Child-1", "shared.sldprt", "Default", 2, blocked)
    Set grandchild = Component("Blocked-1/Middle-1/Child-1", "shared.sldprt", "Default", 2, middle)
    Set direct = Component("Direct-1", "direct.sldprt", "Default", 0, Nothing)
    For Each item In Array(blocked, middle, child, grandchild, direct)
        item.ForbiddenAccess = True
    Next
    Set active = Component("Active-1", "active.sldasm", "Default", 2, Nothing)
    Set first = Component("Active-1/Shared-1", "shared.sldprt", "Default", 2, active)
    Set second = Component("Active-1/Shared-2", "SHARED.sldprt", "Default", 2, active)
    Set hidden = Component("Hidden-1", "hidden.sldprt", "Default", 3, Nothing)
    Set light = Component("Light-1", "light.sldprt", "Default", 1, active)
    Set fullLight = Component("FullLight-1", "full-light.sldprt", "Default", 4, Nothing)
    Set alternate = Component("Alternate-1", "shared.sldprt", "Alternate", 2, Nothing)
    Set assembly = New FakeAssembly
    ' Deliberately put descendants before parents, as GetComponents has no tree order.
    assembly.Components = Array(grandchild, first, child, hidden, direct, second, light, _
                                middle, fullLight, blocked, active, alternate)
    Inventory assembly
    Check quantities.Count = 6, "Suppressed-only models entered the drawing queue."
    Check skipped = 5, "A suppressed occurrence or descendant was not logged."
    Check warnings = 0, "Suppression produced an unresolved-component warning."
    CheckQuantity "shared.sldprt", "Default", 2
    CheckQuantity "shared.sldprt", "Alternate", 1
    CheckQuantity "active.sldasm", "Default", 1
    CheckQuantity "hidden.sldprt", "Default", 1
    CheckQuantity "light.sldprt", "Default", 1
    CheckQuantity "full-light.sldprt", "Default", 1

    ' A new inventory observes changed assembly-context states, without a stale cache.
    blocked.State = 2
    For Each item In Array(blocked, middle, child, grandchild)
        item.ForbiddenAccess = False
    Next
    Set quantities = CreateObject("Scripting.Dictionary")
    quantities.CompareMode = 1
    skipped = 0
    Inventory assembly
    Check quantities.Count = 8, "Reactivated subassembly descendants stayed excluded."
    Check skipped = 1, "Direct suppression was lost when another branch was enabled."
    CheckQuantity "shared.sldprt", "Default", 4

    Set emptyAssembly = New FakeAssembly
    emptyAssembly.Components = Empty
    Inventory emptyAssembly
    Check quantities.Count = 8, "An empty assembly changed the inventory."

    ' Keep the existing unresolved-model warning for active components only.
    Set missing = Component("Missing-1", "missing.sldprt", "Default", 2, Nothing)
    Set missing.Document = Nothing
    assembly.Components = Array(missing)
    Inventory assembly
    Check warnings = 1, "An unresolved active component was silently ignored."
    Check quantities.Count = 8, "An unresolved model entered the drawing queue."
    WScript.Echo "PASS: optional suppression filter; CLI/INI precedence; filtered ancestors; quantities preserved"
End Sub

Main
'''

root = work_path('inventory_suppression')
root.mkdir(parents=True, exist_ok=True)
with tempfile.TemporaryDirectory(dir=root) as folder:
    script = Path(folder) / 'inventory.vbs'
    script.write_text('Option Explicit\n' + '\n'.join(constants + routines) + '\n' + harness,
                      encoding='ascii')
    result = subprocess.run([str(Path(os.environ['WINDIR']) / 'System32/cscript.exe'),
                             '//nologo', str(script)], capture_output=True, timeout=20)
    if result.returncode != 0 or b'PASS:' not in result.stdout:
        diagnostic = result.stderr.decode('oem', errors='replace')
        location = re.search(r'\((\d+),\s*\d+\)', diagnostic)
        if location:
            line = int(location[1])
            diagnostic += '\n' + script.read_text('ascii').splitlines()[line - 1]
        raise AssertionError((result.stdout.decode('oem', errors='replace'), diagnostic))
    print(result.stdout.decode('ascii').strip())
