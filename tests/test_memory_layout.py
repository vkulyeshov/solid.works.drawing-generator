"""Execute workspace and reference geometry from the canonical VBA under WSH."""
from pathlib import Path
import os
import re
import subprocess
import sys
import tempfile
import unittest

sys.dont_write_bytecode = True
BASE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE / 'tools'))
from project_paths import work_path
from vba_sources import read_source


class MemoryLayoutTests(unittest.TestCase):
    def test_candidate_search_does_not_call_cad(self):
        source = read_source()
        for name in ['MemoryBuildCandidates', 'MemoryCandidateFits', 'MemoryLeaderCandidate',
                     'MemoryAngularCandidate', 'MemoryLinearTextCandidate', 'MemoryTranslateLinearCandidate', 'MemoryReferenceCandidates',
                     'MemoryOppositeAngularCandidate', 'MemorySolveCandidates', 'ReferenceLabelCandidate',
                     'ReferenceLeaderLines', 'ReferenceOvalBounds', 'WorkspaceFrame', 'WorkspaceFormatBlocks']:
            body = re.search(r'(?ms)^Private (?:Function|Sub) ' + name + r'\(.*?^End (?:Function|Sub)', source)[0]
            for forbidden in ['GraphicsRedraw', 'SetPosition2', 'ForceRebuild', 'AlignDimensions', 'GetDisplayData']:
                self.assertNotIn(forbidden, body, name)
        self.assertIn('ElseIf MemoryLayoutEnabled() Then', (BASE / 'modules/native_layout.bas').read_text('utf8'))
        self.assertIn('AddProfileDetailReferences d', (BASE / 'modules/profile_pages.bas').read_text('utf8'))
        self.assertIn('"|LayoutEngine="', (BASE / 'modules/checkpoint.bas').read_text('utf8'))
        reference = re.search(r'(?ms)^Private Sub RegisterMemoryReference\(.*?^End Sub', source)[0]
        self.assertLess(reference.index('ClearSelection2 True'), reference.index('InsertNote(label)'))
        self.assertIn('format.Escapement = 0', reference)
        self.assertIn('MemoryAppendReferences records, views, frame, obstacles', source)

    def test_native_workspace_and_reference_math(self):
        source = read_source()
        routines = []
        for name in ['WorkspaceFrame', 'WorkspaceInternalBox', 'WorkspaceTableBounds', 'MemoryWitnessFoot', 'MemoryAngularEndpoint',
                     'DetailReferenceLabel', 'MinD', 'MaxD', 'NativeBoundsInside', 'NativeBoxesOverlap',
                     'MemoryLinearTextCandidate', 'MemoryRebound', 'NativeExpandBounds', 'PlanarShiftPoint', 'PlanarGraphicLength',
                     'MemorySolveCandidates', 'ReferenceOvalBounds', 'ReferenceLeaderLines', 'SpatialClear', 'LayoutLinearDirection', 'MemoryCanSampleNatively', 'LayoutDimensionKind']:
            body = re.search(r'(?ms)^Private (?:Function|Sub) ' + name + r'\(.*?^End (?:Function|Sub)', source)[0]
            body = body.replace('Private ', '', 1)
            body = re.sub(r' As SldWorks\.\w+', '', body)
            body = body.replace('Mid$(', 'Mid(')
            body = body.replace(' As New Collection', '')
            body = re.sub(r' As (?:Collection|Object|Variant|String|Double|Long|Boolean)\b', '', body)
            body = re.sub(r'(?m)^(\s*)Next \w+\s*$', r'\1Next', body)
            body = re.sub(r'(For Each \w+ In (?:seed|candidate)\("(?:text|lines|heads)"\))', r'\1.Items', body)
            if name == 'PlanarGraphicLength':
                body = body.replace(' In lines\n', ' In lines.Items\n')
            if name == 'MemoryLinearTextCandidate':
                body = body.replace('\n', '\n    Set text = CreateObject("Scripting.Dictionary")\n    Set lines = CreateObject("Scripting.Dictionary")\n    Set bases = CreateObject("Scripting.Dictionary")\n', 1)
                for collection, value in [('text', 'moved'), ('lines', 'line'), ('bases', 'line')]:
                    body = body.replace(collection + '.Add ' + value, collection + '.Add ' + collection + '.Count + 1, ' + value)
            if name == 'MemorySolveCandidates':
                body = body.replace('\n', '\n    Set best = Nothing\n', 1)
                body = body.replace('New Collection', 'CreateObject("Scripting.Dictionary")')
                body = re.sub(r'(For Each \w+ In [^\n]+)', r'\1.Items', body)
                for collection, value in [('available', 'alternative'), ('nextRemaining', 'child'), ('chosen', 'candidate')]:
                    body = body.replace(collection + '.Add ' + value, collection + '.Add ' + collection + '.Count + 1, ' + value)
            if name == 'ReferenceLeaderLines':
                body = body.replace('\n', '\n    Set result = CreateObject("Scripting.Dictionary")\n', 1)
                body = body.replace('result.Add Array(', 'result.Add result.Count + 1, Array(')
            routines.append(body)
        definitions = dict(re.findall(r'^Private Const (\w+) As \w+ = (.+)$', source, re.M))
        wanted = set(re.findall(r'\b[A-Z][A-Z_]+\b', '\n'.join(routines))) & definitions.keys()
        wanted.add('SW_DIMENSION_ANGULAR')
        constants = ['Const ' + name + ' = ' + definitions[name] for name in sorted(wanted)]
        harness = '''
Sub Check(ByVal condition, ByVal message)
    If Not condition Then
        Err.Raise 1001, , message
    End If
End Sub
Sub CheckExecutionBudget()
End Sub
Function MemoryStrictSearchExpired()
    MemoryStrictSearchExpired=False
End Function
Dim lines, frame, point, seed, text, bases, heads, candidate, planarSearchLimit
Dim oval, leaders, source, anchor, shelf
Class FakeArrowData
    Public count, first, last
    Function GetArrowHeadCount()
        GetArrowHeadCount=count
    End Function
    Function GetArrowHeadAtIndex2(ByVal index)
        GetArrowHeadAtIndex2=first
        If index=1 Then
            GetArrowHeadAtIndex2=last
        End If
    End Function
End Class
Class FakeLinearDisplay
    Public data
    Function GetDisplayData()
        Set GetDisplayData=data
    End Function
End Class
Dim arrowData, linearDisplay, measuredAxis
Set arrowData=New FakeArrowData
Set linearDisplay=New FakeLinearDisplay
Set linearDisplay.data=arrowData
arrowData.count=1
arrowData.first=Array(0.04,0.16,0,-0.9396926208,-0.3420201433,0)
measuredAxis=LayoutLinearDirection(linearDisplay)
Check Abs(measuredAxis(0)+0.9396926208)<0.000000001,"A one-sided linear dimension lost its measured axis"
Check Abs(measuredAxis(1)+0.3420201433)<0.000000001,"A one-sided linear dimension was treated as a circular leader"
arrowData.count=2
arrowData.first=Array(0.04,0.16,0,0,0,0)
arrowData.last=Array(0.04,0.18,0,0,0,0)
measuredAxis=LayoutLinearDirection(linearDisplay)
Check Abs(measuredAxis(0))<0.000000001 And Abs(measuredAxis(1)-1)<0.000000001,"The two-arrow axis changed"
arrowData.count=0
Check IsEmpty(LayoutLinearDirection(linearDisplay)),"A dimension without arrows acquired an invented axis"
Dim probeRecord
Check LayoutDimensionKind(SW_DIMENSION_HORIZONTAL_LINEAR)=SW_DIMENSION_LINEAR,"An imported horizontal dimension entered circular routing"
Check LayoutDimensionKind(SW_DIMENSION_VERTICAL_LINEAR)=SW_DIMENSION_LINEAR,"An imported vertical dimension entered circular routing"
Check LayoutDimensionKind(SW_DIMENSION_DIAMETER)=SW_DIMENSION_DIAMETER,"A real diameter lost its type"
Set probeRecord=CreateObject("Scripting.Dictionary")
probeRecord.Add "type",SW_DIMENSION_LINEAR
Check Not MemoryCanSampleNatively(probeRecord),"A missing linear axis was replaced with a circular leader"
probeRecord.Add "direction",Array(1,0)
Check MemoryCanSampleNatively(probeRecord),"An unstable linear annotation cannot be verified"
probeRecord("type")=SW_DIMENSION_DIAMETER
Check Not MemoryCanSampleNatively(probeRecord),"A circular annotation acquired an invented anchor"
probeRecord.Add "anchor",Array(0.1,0.1)
Check MemoryCanSampleNatively(probeRecord),"A reattaching diameter cannot be verified"
probeRecord("type")=SW_DIMENSION_RADIAL
Check MemoryCanSampleNatively(probeRecord),"A reattaching radius cannot be verified"
probeRecord("type")=SW_DIMENSION_ANGULAR
Check Not MemoryCanSampleNatively(probeRecord),"Angular sectors entered the circular sampling path"
frame=Array(0.021,0.006,0.204,0.291)
oval=ReferenceOvalBounds(Array(0.08,0.12,0.10,0.13),frame)
Check NativeBoundsInside(oval,frame), "A source oval exceeded the actual drawing frame"
Check oval(2)-oval(0)>0.02 And oval(3)-oval(1)>0.01, "An oval failed to enclose its feature family"
oval=ReferenceOvalBounds(Array(0.08,0.12,0.084,0.124),frame)
Check Abs((oval(2)-oval(0))-(oval(3)-oval(1)))<0.000001, "A compact source did not receive a circle"
source=Array(0.06,0.10,0.12,0.12)
text=Array(0.14,0.15,0.145,0.155)
Set leaders=ReferenceLeaderLines(source,text)
anchor=leaders(1): shelf=leaders(2)
Check Abs(((anchor(0)-0.09)/0.03)^2+((anchor(1)-0.11)/0.01)^2-1)<0.000001, "The leader is detached from the source oval"
Check Abs(shelf(1)-shelf(3))<0.000001 And shelf(1)<text(1), "The reference shelf is not horizontal below the letter"
Check Abs(anchor(2)-shelf(0))<0.000001 And Abs(anchor(3)-shelf(1))<0.000001, "The leader does not meet its shelf"
Function PlanarCandidatesCompatibleCached(ByVal first, ByVal last)
    Dim a, b
    a = first("id"): b = last("id")
    PlanarCandidatesCompatibleCached = Not ((a="a1" And Left(b,1)="b") Or (a="b1" And Left(b,1)="c"))
End Function
Dim records, record, options, chosen, steps, group, index
Set records = CreateObject("Scripting.Dictionary")
For Each group In Array("a","b","c")
    Set record = CreateObject("Scripting.Dictionary")
    Set options = CreateObject("Scripting.Dictionary")
    For index = 1 To 2
        Set candidate = CreateObject("Scripting.Dictionary")
        candidate.Add "id", group & index
        options.Add index,candidate
    Next
    record.Add "name",group
    record.Add "candidates",options
    records.Add records.Count+1,record
Next
Set chosen = CreateObject("Scripting.Dictionary")
steps=0: planarSearchLimit=20
Check MemorySolveCandidates(records,chosen,steps), "Forward domain propagation lost a valid assignment"
Check chosen.Count=3 And chosen(1)("id")="a2" And chosen(2)("id")="b2", "Backtracking retained an incompatible candidate"
Check records(1)("candidates").Count=2, "A child search mutated its parent's candidate domain"
Set chosen = CreateObject("Scripting.Dictionary")
steps=0: planarSearchLimit=2
Check Not MemorySolveCandidates(records,chosen,steps) And chosen.Count=0, "A bounded failed search retained a partial assignment"
Check steps<=planarSearchLimit+1, "Exhausted search continued through parent candidate domains"
Set records=CreateObject("Scripting.Dictionary")
For Each group In Array("x","y")
    Set record=CreateObject("Scripting.Dictionary")
    Set options=CreateObject("Scripting.Dictionary")
    For index=1 To 3
        Set candidate=CreateObject("Scripting.Dictionary")
        candidate.Add "id",group & index
        options.Add index,candidate
    Next
    record.Add "name",group: record.Add "candidates",options
    records.Add records.Count+1,record
Next
Set chosen=CreateObject("Scripting.Dictionary")
steps=0: planarSearchLimit=1
Check Not MemorySolveCandidates(records,chosen,steps), "A search ignored its recursion budget"
Check steps=2 And chosen.Count=0, "Parent branches kept scanning after a child exhausted the budget"
Class FakeTable
    Public ColumnCount, RowCount, AnchorType
    Public Function GetAnnotation()
        Set GetAnnotation = Me
    End Function
    Public Function GetPosition()
        GetPosition = Array(0.1,0.2,0)
    End Function
    Public Function ColumnHidden(ByVal index)
        ColumnHidden = index=1
    End Function
    Public Function RowHidden(ByVal index)
        RowHidden = False
    End Function
    Public Function GetColumnWidth(ByVal index)
        GetColumnWidth = 0.03
    End Function
    Public Function GetRowHeight(ByVal index)
        GetRowHeight = 0.01
    End Function
End Class
Dim table
Set table = New FakeTable
table.ColumnCount=2: table.RowCount=3: table.AnchorType=1
point = WorkspaceTableBounds(table)
Check Abs(point(0)-0.1)<0.000001 And Abs(point(1)-0.17)<0.000001 And Abs(point(2)-0.13)<0.000001, "A BOM or hidden column occupied the wrong area"
table.AnchorType=4
point = WorkspaceTableBounds(table)
Check Abs(point(0)-0.07)<0.000001 And Abs(point(3)-0.23)<0.000001, "Bottom-right table anchoring was ignored"
lines = Array(Array(0.02,0.005,0.205,0.005), Array(0.02,0.292,0.205,0.292), Array(0.02,0.005,0.02,0.292), Array(0.205,0.005,0.205,0.292), Array(0.02,0.06,0.205,0.06))
frame = WorkspaceFrame(lines,0.21,0.297)
Check Abs(frame(0)-0.021)<0.000001, "The binding margin was ignored"
Check Abs(frame(1)-0.006)<0.000001, "The stamp became the entire bottom margin"
Check Abs(frame(2)-0.204)<0.000001, "The right frame was not read"
Check NativeBoundsInside(Array(0.025,0.008,0.03,0.015),frame), "Free space beside the stamp was lost"
Check IsEmpty(WorkspaceInternalBox(Array(0.008,0.005,0.02,0.15),frame)), "The external binding strip consumed interior paper"
point = WorkspaceInternalBox(Array(0.02,0.278,0.09,0.292),frame)
Check Abs(point(0)-frame(0))<0.000001 And point(1)>0.27, "The top stamp consumed free paper below it"
lines = Array(Array(0.02,0.005,0.292,0.005),Array(0.02,0.205,0.292,0.205),Array(0.02,0.005,0.02,0.205),Array(0.292,0.005,0.292,0.205))
frame = WorkspaceFrame(lines,0.297,0.21)
Check Abs(frame(2)-0.291)<0.000001 And Abs(frame(3)-0.204)<0.000001, "Landscape workspace retained portrait coordinates"
frame = WorkspaceFrame(Array(),0.42,0.297)
Check frame(2)>0.41 And frame(3)>0.29, "An empty format lost the A3 workspace"
point = MemoryWitnessFoot(Array(0.02,0.04),Array(Array(0.02,0.01,0.02,0.042),Array(0.02,0.04,0.06,0.04)),1,0)
Check Abs(point(1)-0.01)<0.000001, "A measured witness foot moved with the dimension line"
Check DetailReferenceLabel(1)="\u0410" And DetailReferenceLabel(2)="\u0411" And DetailReferenceLabel(3)="\u0412", "Source and detail letters are not Ukrainian"
Check DetailReferenceLabel(27)="\u0410\u0410", "Detail letters overflowed the alphabet"
Set seed = CreateObject("Scripting.Dictionary")
seed.Add "arcCenter",Array(0,0)
seed.Add "arcRadius",0.02
point = MemoryAngularEndpoint(Array(0.02,0),seed,2)
Check Abs(point(0)-0.04)<0.000001, "An angular arc did not move to its planned radius"
point = MemoryAngularEndpoint(Array(0.001,0),seed,2)
Check Abs(point(0)-0.001)<0.000001, "An angular witness foot was scaled"
Set seed = CreateObject("Scripting.Dictionary")
seed.Add "direction",Array(1,0)
seed.Add "first",Array(0.02,0.02)
seed.Add "last",Array(0.0235,0.02)
seed.Add "baseline",Array(0.02175,0.04)
seed.Add "position",Array(0.035,0.04,0)
seed.Add "controlPoint",Array(0.035,0.04,0)
seed.Add "value",0.0035
seed.Add "cost",0.1
Set text = CreateObject("Scripting.Dictionary")
text.Add 1,Array(Array(0.031,0.04),Array(0.039,0.04),Array(0.039,0.0425),Array(0.031,0.0425))
seed.Add "text",text
Set lines = CreateObject("Scripting.Dictionary")
lines.Add 1,Array(0.02,0.02,0.02,0.04)
lines.Add 2,Array(0.0235,0.02,0.0235,0.04)
lines.Add 3,Array(0.02,0.04,0.04,0.04)
seed.Add "lines",lines
Set bases = CreateObject("Scripting.Dictionary")
bases.Add 1,Array(0.02,0.04,0.04,0.04)
seed.Add "bases",bases
Set heads = CreateObject("Scripting.Dictionary")
seed.Add "heads",heads
Set candidate = MemoryLinearTextCandidate(seed,-1)
Check candidate("position")(0)<0.02, "Free text could not use the opposite end of a short dimension"
Check Abs(candidate("first")(0)-0.02)<0.000001 And candidate("value")=seed("value"), "Free-text routing changed a model reference or measurement"
Check candidate("centerText")=False And candidate("offsetText")=False, "A free-text candidate retained centering"
point = candidate("lines")(1)
Check Abs(point(0)-0.02)<0.000001 And Abs(point(1)-0.02)<0.000001, "Free-text routing moved a witness foot"
WScript.Echo "PASS: A4/A3 workspace, landscape borders, free stamp-side space, witness feet and Ukrainian reference letters"
'''
        root = work_path('memory_layout/unit')
        root.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=root) as temporary:
            script = Path(temporary) / 'memory-layout.vbs'
            script.write_text('Option Explicit\n' + '\n'.join(constants + routines) + '\n' + harness, 'utf-16')
            result = subprocess.run([str(Path(os.environ['WINDIR']) / 'System32/cscript.exe'),
                                     '//nologo', str(script)], capture_output=True, timeout=20)
            if result.returncode or b'PASS:' not in result.stdout:
                (root / 'failed-memory-layout.vbs').write_bytes(script.read_bytes())
            self.assertEqual(result.returncode, 0, (result.stdout, result.stderr))
            self.assertIn(b'PASS:', result.stdout, (result.stdout, result.stderr))


if __name__ == '__main__':
    unittest.main(verbosity=2)
