"""Keep measured native label ink separate from dimension-line tails."""
from pathlib import Path
import re
import subprocess
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from project_paths import work_path
from vba_sources import read_source

source = read_source()
procedures = []
for name in ['NativeTextBounds', 'NativeExpandBounds', 'PlanarTextPolygons', 'MinD']:
    body = re.search(r'(?ms)^Private (?:Function|Sub) ' + name + r'\(.*?^End (?:Function|Sub)', source)[0]
    body = body.replace('Private ', '', 1)
    body = body.replace('Optional ByVal owner As Object = Nothing', 'ByVal owner As Object')
    body = body.replace('Optional ByVal printed As Boolean = False', 'ByVal printed As Boolean')
    body = body.replace('PlanarTextPolygons(owner, data, 0)', 'PlanarTextPolygons(owner, data, 0, False)')
    body = body.replace(' As New Collection', '')
    body = re.sub(r' As (?:SldWorks\.\w+|Object|Collection|Variant|Long|Double|Boolean|String)\b', '', body)
    body = re.sub(r'\b(\w+)\$\(', r'\1(', body)
    body = re.sub(r'(?m)^(\s*)Next \w+\s*$', r'\1Next', body)
    if name == 'PlanarTextPolygons':
        body = body.replace('    For index =', '    Set result = CreateObject("Scripting.Dictionary")\n    For index =', 1)
        body = body.replace('        result.Add polygon', '        result.Add result.Count, polygon')
        body = body.replace('    Set PlanarTextPolygons = result', '    PlanarTextPolygons = result.Items')
        body = body.replace('            GoTo NextPlanarText', '            Err.Raise 513, , "Fixture text must be nonempty"')
        body = body.replace('NextPlanarText:', '')
    procedures.append(body)
constants = ['Const ' + m[1] + ' = ' + re.sub(r'(?<=\d)#', '', m[2])
             for m in re.finditer(r'^Private Const (\w+) As \w+ = (.+)$', source, re.M)]
harness = r'''
Function PanelOwnsDimension(ByVal owner)
    PanelOwnsDimension = False
End Function
Function PlanarMeasureTextWidth(ByVal owner, ByVal text, ByVal height, ByVal printed, ByVal face)
    PlanarMeasureTextWidth = Len(text) * height * 0.4
End Function
Class DataStub
    Public Angle
    Public Text
    Private Sub Class_Initialize
        Text = " 6,39 "
    End Sub
    Public Function GetTextCount()
        GetTextCount = 1
    End Function
    Public Function GetTextAtIndex(ByVal index)
        GetTextAtIndex = Text
    End Function
    Public Function GetTextHeightAtIndex(ByVal index)
        GetTextHeightAtIndex = 0.0035
    End Function
    Public Function GetTextPositionAtIndex(ByVal index)
        GetTextPositionAtIndex = Array(0,0,0)
    End Function
    Public Function GetTextAngleAtIndex(ByVal index)
        GetTextAngleAtIndex = Angle
    End Function
    Public Function GetTextFontAtIndex(ByVal index)
        GetTextFontAtIndex = "Test Font"
    End Function
    Public Function GetLineCount()
        GetLineCount = 1
    End Function
    Public Function GetLineAtIndex2(ByVal index)
        GetLineAtIndex2 = Array(0,0,-1,-1,0,0,0,0.0093,0,0)
    End Function
End Class
Class OwnerStub
    Public Prefix, Suffix, HoleValue
    Public Function GetAnnotation()
        Set GetAnnotation = New AnnotationStub
    End Function
    Public Function IsHoleCallout()
        IsHoleCallout = HoleValue
    End Function
    Public Function GetText(ByVal part)
        GetText = ""
        If part = SW_DIMENSION_TEXT_PREFIX Then
            GetText = Prefix
        ElseIf part = SW_DIMENSION_TEXT_SUFFIX Then
            GetText = Suffix
        End If
    End Function
End Class
Class AnnotationStub
    Public Function GetTextFormat(ByVal index)
        Set GetTextFormat = New FormatStub
    End Function
    Public Function GetPosition()
        GetPosition = Array(0, 0, 0)
    End Function
End Class
Class FormatStub
    Public CharHeight, TypeFaceName
    Private Sub Class_Initialize
        CharHeight = 0.0035
        TypeFaceName = "Test Font"
    End Sub
End Class
Dim data, owner, bounds, printedOwner, polygons, polygon, firstWidth, width
Set data = New DataStub
Set owner = CreateObject("Scripting.Dictionary")
data.Angle = 0
bounds = NativeTextBounds(data,owner)
If Abs(bounds(0)-0.0014)>0.000001 Or Abs(bounds(2)-0.007)>0.000001 Then
    Err.Raise 513, , "A dimension tail or whitespace inflated measured panel text"
End If
data.Angle = PI/2
bounds = NativeTextBounds(data,owner)
If Abs(bounds(1)-0.0014)>0.000001 Or Abs(bounds(3)-0.007)>0.000001 Then
    Err.Raise 513, , "Rotated measured panel text has incorrect bounds"
End If
Set printedOwner = New OwnerStub
printedOwner.Prefix = ""
printedOwner.Suffix = " (2 places)"
printedOwner.HoleValue = False
data.Angle = 0
polygons = PlanarTextPolygons(printedOwner, data, 0, True)
polygon = polygons(0)
width = polygon(2)(0)-polygon(0)(0)
If Abs(width-Len(Trim(data.Text & printedOwner.Suffix))*0.0035*0.4)>0.000001 Then
    Err.Raise 513, , "A missing count suffix was excluded from printed ink"
End If
firstWidth = width
data.Text = data.Text & printedOwner.Suffix
polygons = PlanarTextPolygons(printedOwner, data, 0, True)
polygon = polygons(0)
If Abs(polygon(2)(0)-polygon(0)(0)-firstWidth)>0.000001 Then
    Err.Raise 513, , "A suffix already returned by DisplayData was added twice"
End If
printedOwner.HoleValue = 1
printedOwner.Prefix = "Thread text returned through the COM prefix accessor"
printedOwner.Suffix = ""
data.Text = "8 x"
polygons = PlanarTextPolygons(printedOwner, data, 0, True)
polygon = polygons(0)
If Abs(polygon(2)(0)-polygon(0)(0)-Len(data.Text)*0.0035*0.4)>0.000001 Then
    Err.Raise 513, , "A COM True value of 1 incorrectly expanded a hole callout"
End If
WScript.Echo "Panel text bounds passed: measured ink, whitespace, and rotated text"
'''
directory = work_path('panel_text_tests')
directory.mkdir(parents=True, exist_ok=True)
script = directory / 'bounds.vbs'
script.write_text('\n'.join(constants + procedures) + '\n' + harness, encoding='utf-16')
result = subprocess.run(['cscript.exe', '//nologo', str(script)], capture_output=True,
                        text=True, timeout=30)
print(result.stdout, end='')
print(result.stderr, end='', file=sys.stderr)
raise SystemExit(result.returncode)
