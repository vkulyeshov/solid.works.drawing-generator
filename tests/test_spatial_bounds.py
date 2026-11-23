"""Differentially test cached broad-phase checks against exact VBA predicates."""
from pathlib import Path
import os
import re
import subprocess
import sys
import tempfile
import unittest

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from project_paths import work_path
from vba_sources import read_source


def script_routine(source, name):
    body = re.search(r'(?ms)^Private (?:Function|Sub) ' + name + r'\(.*?^End (?:Function|Sub)', source)[0]
    collections = set(re.findall(r'ByVal (\w+) As Collection', body.split('\n', 1)[0]))
    local_collections = re.findall(r'(\w+) As New Collection', body)
    body = body.replace('Private ', '', 1).replace(' As New Collection', '')
    body = re.sub(r' As (?:Collection|Object|Variant|String|Double|Long|Boolean)\b', '', body)
    body = re.sub(r'(?m)^(\s*)Next \w+\s*$', r'\1Next', body)
    def enumeration(match):
        expression = match[2].strip()
        if expression in collections or expression == 'candidate("heads")':
            expression += '.Items'
        return match[1] + expression
    body = re.sub(r'(For Each \w+ In )([^\n]+)', enumeration, body)
    for collection in local_collections:
        body = body.replace('\n', '\n    Set ' + collection + ' = CreateObject("Scripting.Dictionary")\n', 1)
        body = body.replace(collection + '.Add ', collection + '.Add ' + collection + '.Count + 1, ')
    if name in ('ProfileLineCutsMaterial', 'PlanarPolygonsOverlap'):
        body = body.replace('\n', '\n    exactCalls = exactCalls + 1\n', 1)
    return body


HARNESS = '''
Dim planarStrictCrossings, exactCalls, expectedCalls, optimizedCalls
Sub Check(ByVal condition, ByVal message)
    If Not condition Then
        Err.Raise 1001,,message
    End If
End Sub
Function Rectangle(ByVal x, ByVal y, ByVal width, ByVal height, ByVal angle)
    Dim result(3), index, dx, dy
    For index=0 To 3
        dx=-width/2: dy=-height/2
        If index=1 Or index=2 Then
            dx=width/2
        End If
        If index>=2 Then
            dy=height/2
        End If
        result(index)=Array(x+dx*Cos(angle)-dy*Sin(angle),y+dx*Sin(angle)+dy*Cos(angle))
    Next
    Rectangle=result
End Function
Function Candidate(ByVal x, ByVal y)
    Dim result, text, lines, heads, bases, index, line
    Set result=CreateObject("Scripting.Dictionary")
    Set text=CreateObject("Scripting.Dictionary")
    Set lines=CreateObject("Scripting.Dictionary")
    Set heads=CreateObject("Scripting.Dictionary")
    Set bases=CreateObject("Scripting.Dictionary")
    text.Add 1,Rectangle(x,y,0.001+Rnd*0.014,0.002+Rnd*0.003,Rnd*3.14)
    For index=1 To 4
        line=Array(x+(Rnd-0.5)*0.04,y+(Rnd-0.5)*0.04,x+(Rnd-0.5)*0.04,y+(Rnd-0.5)*0.04)
        lines.Add index,line
        If index=1 Then
            bases.Add 1,line
        End If
    Next
    heads.Add 1,Array(Array(x,y),Array(x+0.0025,y-0.0005),Array(x+0.0025,y+0.0005))
    result.Add "text",text: result.Add "lines",lines
    result.Add "bases",bases: result.Add "heads",heads
    Set Candidate=result
End Function
Function ExactLinesHitText(ByVal lines, ByVal text)
    Dim line, polygon
    For Each line In lines.Items
        For Each polygon In text.Items
            If ProfileLineCutsMaterial(Array(line(0),line(1)),Array(line(2),line(3)),Array(polygon)) Then
                ExactLinesHitText=True
                Exit Function
            End If
        Next
    Next
End Function
' Frozen pre-optimization decision sequence; exact geometry routines are shared.
Function ExactCompatible(ByVal first, ByVal last)
    Dim a,b
    For Each a In first("text").Items
        For Each b In last("text").Items
            If PlanarPolygonsOverlap(a,b) Then
                Exit Function
            End If
        Next
    Next
    If ExactLinesHitText(first("lines"),last("text")) Or ExactLinesHitText(last("lines"),first("text")) Then
        Exit Function
    End If
    For Each a In first("bases").Items
        For Each b In last("bases").Items
            If PlanarDimensionLineConflict(a,b) Or PlanarDimensionLineConflict(b,a) Then
                Exit Function
            End If
        Next
    Next
    If planarStrictCrossings Then
        If PlanarBasesCrossLines(first("bases"),last("lines")) Or PlanarBasesCrossLines(last("bases"),first("lines")) Then
            Exit Function
        End If
    End If
    If PlanarHeadsOverlap(first("heads"),last("heads")) Then
        Exit Function
    End If
    ExactCompatible=True
End Function
Dim index, mode, first, last, expected, actual, cache, bounds, original, duplicate, key, text
Randomize 17
expectedCalls=0: optimizedCalls=0
For index=1 To 600
    Set first=Candidate((Rnd-0.5)*0.04,(Rnd-0.5)*0.04)
    Set last=Candidate((Rnd-0.5)*0.08,(Rnd-0.5)*0.08)
    For mode=0 To 1
        planarStrictCrossings=(mode=1)
        exactCalls=0: expected=ExactCompatible(first,last): expectedCalls=expectedCalls+exactCalls
        exactCalls=0: actual=SpatialCandidatesCompatible(first,last): optimizedCalls=optimizedCalls+exactCalls
        Check expected=actual,"Broad-phase changed a strict or relaxed collision decision: " & index
    Next
Next
Check optimizedCalls<expectedCalls,"The broad phase did not skip any exact geometry work"
Check SpatialBoundsOverlap(Array(0,0,1,0),Array(0.5,0,2,0)),"Touching zero-height bounds were discarded"
Check SpatialBoundsOverlap(Array(0,0,1,1),Array(1+PROFILE_MATERIAL_TOLERANCE_M/2,0,2,1)),"A tolerance-sized gap was discarded"
Check Not SpatialBoundsOverlap(Array(0,0,1,1),Array(1+PROFILE_MATERIAL_TOLERANCE_M*2,0,2,1)),"Separated bounds reached exact geometry"
Dim gap, a, b, loops, line
For Each gap In Array(0,PROFILE_MATERIAL_TOLERANCE_M/2,PROFILE_MATERIAL_TOLERANCE_M*2)
    Set a=CreateObject("Scripting.Dictionary"): Set b=CreateObject("Scripting.Dictionary")
    a.Add 1,Rectangle(0,0,0.01,0.005,0)
    b.Add 1,Rectangle(0.01+gap,0,0.01,0.005,0)
    Check SpatialPolygonsOverlap(a,b,SpatialPolygonBoxes(a),SpatialPolygonBoxes(b))=PlanarPolygonsOverlap(a(1),b(1)),"Near-tolerance text contact changed"
    Set a=CreateObject("Scripting.Dictionary"): Set b=CreateObject("Scripting.Dictionary")
    a.Add 1,Array(0,0,0.02,0): b.Add 1,Array(0.005,gap,0.015,gap)
    Check SpatialBasesConflict(a,b,SpatialLineBoxes(a),SpatialLineBoxes(b))=PlanarDimensionLineConflict(a(1),b(1)),"Near-collinear baseline conflict changed"
Next
loops=Array(Array(Array(0,0),Array(0.02,0),Array(0.02,0.005),Array(0.005,0.005),Array(0.005,0.02),Array(0,0.02)),Rectangle(0.0025,0.0025,0.001,0.001,0))
For index=1 To 500
    line=Array((Rnd-0.5)*0.08,(Rnd-0.5)*0.08,(Rnd-0.5)*0.08,(Rnd-0.5)*0.08)
    expected=ProfileLineCutsMaterial(Array(line(0),line(1)),Array(line(2),line(3)),loops)
    If expected Then
        Check SpatialBoundsOverlap(SpatialLineBounds(line),PanelMaterialBounds(loops(0))),"The material gate discarded a cut through a concave profile with a hole"
    End If
Next
Set first=Candidate(0,0)
Set cache=SpatialCandidate(first): original=cache("text")(1)
Set duplicate=CreateObject("Scripting.Dictionary")
For Each key In first.Keys
    duplicate.Add key,first(key)
Next
Set text=CreateObject("Scripting.Dictionary")
text.Add 1,Rectangle(1,1,0.01,0.005,0)
Set duplicate("text")=text
SpatialClear duplicate
Set cache=SpatialCandidate(duplicate): bounds=cache("text")(1)
Check bounds(0)>0.9 And first.Exists("spatial"),"A moved copy retained stale bounds or cleared its source cache"
Set cache=SpatialCandidate(first): bounds=cache("text")(1)
Check bounds(0)=original(0),"A derived candidate mutated cached source bounds"
WScript.Echo "PASS: 1200 exact-equivalence cases; exact calls " & expectedCalls & " -> " & optimizedCalls & "; touching bounds and cache isolation verified"
'''


class SpatialBoundsTests(unittest.TestCase):
    def test_exact_equivalence_and_cache_isolation(self):
        source = read_source()
        names = ['MinD', 'MaxD', 'PanelMaterialBounds', 'ProfilePolygonLocation',
                 'ProfilePointInMaterial', 'ProfileLineCutsMaterial', 'PlanarPolygonsOverlap',
                 'PlanarSegmentsCross', 'PlanarDimensionLineConflict', 'PlanarBasesCrossLines',
                 'PlanarShrinkTriangle', 'PlanarHeadsOverlap', 'SpatialBoundsOverlap',
                 'SpatialLineBounds', 'SpatialLineBoxes', 'SpatialPolygonBoxes', 'SpatialClear',
                 'SpatialCandidate', 'SpatialPolygonsOverlap', 'SpatialLinesHitText',
                 'SpatialBasesConflict', 'SpatialBasesCrossLines', 'SpatialCandidatesCompatible']
        routines = [script_routine(source, name) for name in names]
        definitions = dict(re.findall(r'^Private Const (\w+) As \w+ = (.+)$', source, re.M))
        wanted = set(re.findall(r'\b[A-Z][A-Z_]+\b', '\n'.join(routines))) & definitions.keys()
        constants = ['Const ' + name + ' = ' + definitions[name] for name in sorted(wanted)]
        root = work_path('spatial_bounds/unit')
        root.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=root) as temporary:
            script = Path(temporary) / 'spatial.vbs'
            script.write_text('Option Explicit\n' + '\n'.join(constants + routines) + '\n' + HARNESS, 'utf16')
            result = subprocess.run([str(Path(os.environ['WINDIR']) / 'System32/cscript.exe'),
                                     '//nologo', str(script)], capture_output=True, timeout=60)
            if result.returncode or b'PASS:' not in result.stdout:
                (root / 'failed-spatial.vbs').write_bytes(script.read_bytes())
            self.assertEqual(result.returncode, 0, (result.stdout, result.stderr))
            self.assertIn(b'PASS:', result.stdout, (result.stdout, result.stderr))
            print(result.stdout.decode('ascii').strip())


if __name__ == '__main__':
    unittest.main(verbosity=2)
