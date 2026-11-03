"""Run the VBA planar constraint solver against independent geometric fixtures."""
from pathlib import Path
import os
import re
import subprocess
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from project_paths import work_path
from vba_sources import read_source

source = read_source()
names = ['MinD', 'ProfileRoot', 'PlanarAngle', 'PlanarAcuteAngle', 'PlanarArrowBisector', 'PlanarDatumPlan', 'PlanarDistance', 'PlanarMirror', 'PlanarPointIndex',
         'PlanarSymmetries', 'PlanarJoin', 'PlanarEdgePlan', 'PlanarSegmentsCross',
         'ProfilePolygonLocation', 'ProfilePointInMaterial', 'ProfileLineCutsMaterial',
         'PlanarPolygonsOverlap', 'PlanarArrowTriangle', 'PlanarArrowSpanFits', 'PlanarInsideArrowsFit', 'PlanarClipPolygon', 'PlanarLinesHitText', 'PlanarGraphicLength', 'PlanarDimensionLineConflict', 'PlanarCandidatesCompatible', 'SolvePlanarCandidates']
functions = []
for name in names:
    body = re.search(r'(?ms)^Private (?:Function|Sub) ' + name + r'\(.*?^End (?:Function|Sub)', source)[0]
    body = body.replace('Private ', '', 1)
    body = re.sub(r' As New Collection\b', '', body)
    body = re.sub(r' As (?:Object|Collection|Variant|Long|Double|Boolean|String)\b', '', body)
    body = re.sub(r'(?<=\d)#', '', body)
    body = re.sub(r'(?m)^(\s*)Next \w+\s*$', r'\1Next', body)
    for collection in ['points', 'holes', 'symmetries', 'result', 'lines', 'polygons',
                       'firstText', 'lastText', 'bases', 'remaining', 'options', 'chosen', 'bestOptions']:
        body = re.sub(r'(For Each \w+ In ' + collection + r')\b', r'\1.Items', body)
    body = body.replace('result.Add symmetry', 'result.Add result.Count + 1, symmetry')
    body = body.replace('result.Add group', 'result.Add result.Count + 1, group')
    if name == 'PlanarSymmetries':
        body = body.replace('    For Each node In points.Items',
                            '    Set result = CreateObject("Scripting.Dictionary")\n    For Each node In points.Items', 1)
    if name == 'PlanarEdgePlan':
        body = body.replace('    Set parents =', '    Set result = CreateObject("Scripting.Dictionary")\n    Set parents =', 1)
        body = body.replace('    best = SEARCH_INITIAL_LIMIT', '    Set firstDrop = Nothing\n    Set lastDrop = Nothing\n    best = SEARCH_INITIAL_LIMIT', 1)
    if name == 'SolvePlanarCandidates':
        body = body.replace('PlanarCandidatesCompatibleCached', 'PlanarCandidatesCompatible')
        body = body.replace('    steps = steps + 1', '    Set best = Nothing\n    steps = steps + 1', 1)
        body = body.replace('Set available = New Collection', 'Set available = CreateObject("Scripting.Dictionary")')
        body = body.replace('Set nextRemaining = New Collection', 'Set nextRemaining = CreateObject("Scripting.Dictionary")')
        for collection, item in [('available', 'candidate'), ('nextRemaining', 'record'), ('chosen', 'candidate')]:
            body = body.replace(f'{collection}.Add {item}', f'{collection}.Add {collection}.Count + 1, {item}')
    if name == 'PlanarClipPolygon':
        body = body.replace('Set output = New Collection', 'Set output = CreateObject("Scripting.Dictionary")')
        body = body.replace('output.Add Array(', 'output.Add output.Count + 1, Array(')
        body = body.replace('output.Add last', 'output.Add output.Count + 1, last')
    functions.append(body)
constants = ['Const ' + m[1] + ' = ' + re.sub(r'(?<=\d)#', '', m[2])
             for m in re.finditer(r'^Private Const (\w+) As \w+ = (.+)$', source, re.M)]
harness = r'''
Sub Log(ByVal message)
End Sub
Sub Check(ByVal condition, ByVal message)
    If Not condition Then
        WScript.Echo message
        WScript.Quit 1
    End If
End Sub
Function Nodes(ByVal coordinates)
    Dim result, coordinate, node
    Set result = CreateObject("Scripting.Dictionary")
    For Each coordinate In coordinates
        Set node = CreateObject("Scripting.Dictionary")
        node.Add "x", coordinate(0)
        node.Add "y", coordinate(1)
        node.Add "point", Array(coordinate(0),coordinate(1),0)
        result.Add result.Count + 1, node
    Next
    Set Nodes = result
End Function
Function Candidate(ByVal textBox, ByVal segment)
    Dim result, polygons, lines, bases
    Set result = CreateObject("Scripting.Dictionary")
    Set polygons = CreateObject("Scripting.Dictionary")
    Set lines = CreateObject("Scripting.Dictionary")
    Set bases = CreateObject("Scripting.Dictionary")
    polygons.Add 1, textBox
    lines.Add 1, segment
    result.Add "text", polygons
    result.Add "lines", lines
    result.Add "bases", bases
    Set Candidate = result
End Function
Sub LayoutChecks
    Dim a, b, c, clipped, first, second, records, chosen, options, record, steps, datum, anchor, direction
    Dim firstArrow, lastArrow, triangle
    firstArrow = Array(0,0,0,1,0,0,0.0025,0.00088,1)
    lastArrow = Array(0.008,0,0,-1,0,0,0.0025,0.00088,1)
    Check PlanarInsideArrowsFit(firstArrow,lastArrow), "Inside arrows were rejected because the dimension text needs an external position"
    triangle = PlanarArrowTriangle(firstArrow)
    Check Abs(triangle(1)(0) - 0.0025) < 0.000000001, "The arrowhead body was drawn on the wrong side of its tip"
    lastArrow(0) = 0.004
    Check Not PlanarInsideArrowsFit(firstArrow,lastArrow), "Overlapping inside arrowheads were accepted"
    lastArrow(0) = 0.008
    lastArrow(3) = 1
    Check Not PlanarInsideArrowsFit(firstArrow,lastArrow), "An outside arrow was reported as inside"
    firstArrow = Array(0,0,0,0.6,0.8,0,0.0025,0.00088,1)
    lastArrow = Array(0.006,0.008,0,-0.6,-0.8,0,0.0025,0.00088,1)
    Check PlanarInsideArrowsFit(firstArrow,lastArrow), "Inclined inside arrows were tested only in horizontal paper coordinates"
    a = Array(Array(0,0),Array(0.01,0),Array(0.01,0.004),Array(0,0.004))
    b = Array(Array(0.005,0.002),Array(0.015,0.002),Array(0.015,0.006),Array(0.005,0.006))
    c = Array(Array(0.02,0),Array(0.03,0),Array(0.03,0.004),Array(0.02,0.004))
    Check PlanarPolygonsOverlap(a,b), "Overlapping text polygons were accepted"
    Check Not PlanarPolygonsOverlap(a,c), "Separate text polygons were rejected"
    Check Abs(PlanarAcuteAngle(-3,4) - Atn(0.75)) < 0.000000001, "The local datum angle differs from the model direction"
    Check PlanarAcuteAngle(0,-1) = 0, "An axis-aligned edge produced a spurious angle"
    Check Abs(PlanarArrowBisector(Array(0,0),Array(1,0),Array(0,2)) - PI / 4) < 0.000000001, "The measured angle bisector used an arc fragment or an arrow's radius"
    a = Array(Array(0,0),Array(0.03,0),Array(0.04,0.01),Array(0.03,0.02),Array(0,0.02))
    datum = PlanarDatumPlan(Array(0.03,0),Array(0.04,0.01),a)
    anchor = Array(0.03,0)
    If datum(0) = EDGE_ENDPOINT_LAST Then
        anchor = Array(0.04,0.01)
    End If
    direction = datum(2)
    Check Not ProfileLineCutsMaterial(anchor,Array(anchor(0)+direction(0)*PLANAR_ANGLE_DATUM_LENGTH_M,anchor(1)+direction(1)*PLANAR_ANGLE_DATUM_LENGTH_M),Array(a)), "The local angle datum crossed the convex material boundary"
    Check Abs(datum(3) - PI / 4) < 0.000000001, "The selected local axis changed the true inclined edge angle"
    a = Array(Array(0,0),Array(0.01,0),Array(0.01,0.004),Array(0,0.004))
    Set options = CreateObject("Scripting.Dictionary")
    options.Add 1, Array(0,0,0.003,0.004)
    options.Add 2, Array(0.003,0.004,0.003,0.006)
    Check Abs(PlanarGraphicLength(options) - 0.007) < 0.000000001, "Compact layout scoring did not include complete extension and leader lengths"
    Check Not PlanarPolygonsOverlap(Array(Array(0,0),Array(0.01,0.01),Array(0.009,0.011),Array(-0.001,0.001)),Array(Array(0,0.003),Array(0.01,0.013),Array(0.009,0.014),Array(-0.001,0.004))), "Empty corners of rotated text bounds blocked a clear placement"
    clipped = PlanarClipPolygon(a,Array(0.005,-0.001,0.02,0.003))
    Check UBound(clipped) = 3, "The material polygon was not clipped to the visible region"
    Check Not PlanarPolygonsOverlap(clipped,Array(Array(0,0),Array(0.003,0),Array(0.003,0.004),Array(0,0.004))), "Invisible material blocked a clear detail-view label"
    Check PlanarDimensionLineConflict(Array(0,0,0.03,0),Array(0.01,0,0.02,0)), "Collinear dimension baselines were allowed to overlap"
    Check Not PlanarDimensionLineConflict(Array(0,0,0.01,0),Array(0.01,0,0.02,0)), "Touching baselines were classified as overlapping"
    Set first = Candidate(a,Array(-0.01,0.002,0.04,0.002))
    Set second = Candidate(c,Array(0.025,-0.01,0.025,0.01))
    Check Not PlanarCandidatesCompatible(first,second), "A candidate line through another label was accepted"
    Set first = Candidate(a,Array(-0.01,-0.02,0.04,-0.02))
    Set second = Candidate(c,Array(0.02,-0.005,0.03,-0.005))
    Check PlanarCandidatesCompatible(first,second), "A clear native placement was rejected"
    Set records = CreateObject("Scripting.Dictionary")
    Set chosen = CreateObject("Scripting.Dictionary")
    Set options = CreateObject("Scripting.Dictionary")
    options.Add 1, Candidate(a,Array(0.025,-0.01,0.025,0.01))
    options.Add 2, first
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "name", "first"
    record.Add "candidates", options
    records.Add 1, record
    Set options = CreateObject("Scripting.Dictionary")
    options.Add 1, second
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "name", "second"
    record.Add "candidates", options
    records.Add 2, record
    steps = 0
    Check SolvePlanarCandidates(records,chosen,steps), "Forward checking failed to find the clear assignment"
    Check chosen.Count = 2, "The solver dropped a dimension"
End Sub
Sub Main
    Dim points, holes, plan, group, selected, repeated, hole, symmetries
    Set holes = CreateObject("Scripting.Dictionary")
    Set points = Nodes(Array(Array(0,0),Array(0.02,0),Array(0.02,0.01),Array(0,0.01)))
    Set plan = PlanarEdgePlan(points,holes)
    Check plan.Count = 2, "A reflected rectangle must use one width and one height"
    For Each group In plan.Items
        Check group("count") = 2 And group("selected"), "Reflected rectangle lengths lost their two physical occurrences"
    Next
    Set hole = CreateObject("Scripting.Dictionary")
    hole.Add "x", 0.003
    hole.Add "y", 0.007
    hole.Add "radius", 0.001
    holes.Add 1, hole
    Set symmetries = PlanarSymmetries(points,holes)
    Check symmetries.Count = 0, "An asymmetric hole must invalidate outline-only mirror evidence"
    Set plan = PlanarEdgePlan(points,holes)
    Check plan.Count = 4, "Equal lengths without verified feature symmetry were grouped"
    selected = 0
    For Each group In plan.Items
        If group("selected") Then
            selected = selected + 1
        End If
    Next
    Check selected = 2, "Closed contour directions must reconstruct two independent omitted lengths"
    Set holes = CreateObject("Scripting.Dictionary")
    Set points = Nodes(Array(Array(0,0),Array(0.01,0),Array(0.01,0.003),Array(0.02,0.003),Array(0.02,0),Array(0.03,0),Array(0.03,0.015),Array(0,0.012)))
    Set plan = PlanarEdgePlan(points,holes)
    repeated = 0
    For Each group In plan.Items
        If group("count") = 2 And Abs(group("value")-0.003)<0.000001 Then
            repeated = repeated + 1
        End If
    Next
    Check repeated = 1, "The two equal sides of a local U feature were not recognized"
    points(5)("y") = 0.00005
    Set plan = PlanarEdgePlan(points,holes)
    repeated = 0
    For Each group In plan.Items
        If group("count") = 2 And Abs(group("value")-0.003)<0.000001 Then
            repeated = repeated + 1
        End If
    Next
    Check repeated = 0, "A 0.05 mm difference was hidden by local symmetry grouping"
    Check PlanarSegmentsCross(Array(0,0),Array(1,1),Array(0,1),Array(1,0)), "A diagonal crossing was missed"
    Check Not PlanarSegmentsCross(Array(0,0),Array(1,0),Array(1,0),Array(1,1)), "A shared reference endpoint was classified as a crossing"
    LayoutChecks
    WScript.Echo "PASS: whole-profile and local symmetries, asymmetric holes, independent contour closure, unequal features, and diagonal crossings"
End Sub
Main
'''
target = work_path('planar_constraints', 'solver.vbs')
target.parent.mkdir(parents=True, exist_ok=True)
target.write_text('Option Explicit\n' + '\n'.join(constants + functions) + '\n' + harness, encoding='utf-16')
result = subprocess.run([str(Path(os.environ['WINDIR']) / 'System32/cscript.exe'), '//nologo', str(target)],
                        capture_output=True, timeout=25)
print(result.stdout.decode('utf-8', errors='replace').strip())
if result.returncode or b'PASS:' not in result.stdout:
    raise AssertionError((result.returncode, result.stdout, result.stderr))
