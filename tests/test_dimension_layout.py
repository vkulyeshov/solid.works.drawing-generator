"""Exercise the actual VBA marker and lane policies through Windows Script Host."""
from pathlib import Path
import os
import re
import subprocess
import sys
import tempfile

sys.dont_write_bytecode = True
BASE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE / 'tools'))
from project_paths import work_path

source = (BASE / 'generator.bas').read_text('utf8')
names = ['UseCoordinateDots', 'CoordinateMarkerPadding', 'CoordinateTextCenter', 'FindCoordinateTextCenter',
         'AddForbiddenCoordinate', 'TryCoordinatePlanRow', 'BalanceCoordinateSides',
         'UnionDetailBounds', 'DetailProfileRadius', 'MinD', 'SortedCoordinateKeys',
         'ReferencePairCost', 'LocalNeighborBands', 'LocalDatumPairs', 'ResetCoordinatePlans',
         'CompatibleRequiredCoordinateLanes', 'CompatibleBalanceCoordinateSides',
         'CompatiblePackCoordinateSide', 'CoordinateRegionEdge']
routines = []
for name in names:
    body = re.search(r'(?ms)^Private (Function|Sub) ' + name + r'\(.*?^End (?:Function|Sub)', source)[0]
    body = re.sub(r'^Private ', '', body)
    if name == 'LocalDatumPairs':
        body = body.replace('Dim result As New Collection,', 'Dim result,')
        body = body.replace('    count = coordinates.Count', '    Set result = CreateObject("Scripting.Dictionary")\n    count = coordinates.Count')
        body = body.replace('        result.Add pair', '        result.Add result.Count, pair')
    if name == 'ReferencePairCost':
        body = body.replace('Set a =', 'a =').replace('Set b =', 'b =')
    body = re.sub(r' As (?:Object|Variant|Long|Double|Boolean|String|Collection)\b', '', body)
    body = re.sub(r'(?m)^(\s*)ReDim (?!Preserve)(.+)$', r'\1ReDim \2', body)
    body = re.sub(r'\(0 To ([^)]+)\)', r'(\1)', body)
    body = re.sub(r'(?m)^(\s*)Next \w+\s*$', r'\1Next', body)
    body = re.sub(r'(?m)^(\s*)Exit (Sub|Function)\s*$', r'\1Exit \2', body)
    # A dictionary provides WSH with the object enumeration supplied by VBA Collection.
    body = body.replace('For Each plan In plans', 'For Each plan In plans.Items')
    body = body.replace('For Each other In plans', 'For Each other In plans.Items')
    routines.append(body)
constants = []
for match in re.finditer(r'^Private Const (\w+) As \w+ = (.+)$', source, re.M):
    constants.append('Const ' + match[1] + ' = ' + re.sub(r'(?<=\d)#', '', match[2]))

harness = r'''
Dim coordinateRegion, sheetWidth, sheetHeight
Sub Check(ByVal condition, ByVal message)
    If Not condition Then
        WScript.Echo message
        WScript.Quit 1
    End If
End Sub

Function Plans(ByVal count, ByVal spacing, ByVal span)
    Dim result, plan, i
    Set result = CreateObject("Scripting.Dictionary")
    For i = 0 To count - 1
        Set plan = CreateObject("Scripting.Dictionary")
        plan.Add "desired", 0.05 + i * spacing
        plan.Add "span", span
        plan.Add "width", 0.009
        plan.Add "lowCost", 0
        plan.Add "highCost", 0
        plan.Add "lane", -1
        plan.Add "side", 0
        plan.Add "preferredSide", 0
        plan.Add "geometricCenter", 0.05 + i * spacing
        plan.Add "firstCoordinate", 0.05 + i * spacing - span / 2
        plan.Add "lastCoordinate", 0.05 + i * spacing + span / 2
        plan.Add "lineStart", 0.05 + i * spacing - span / 2 - CoordinateMarkerPadding(span, 0.009)
        plan.Add "lineEnd", 0.05 + i * spacing + span / 2 + CoordinateMarkerPadding(span, 0.009)
        result.Add i + 1, plan
    Next
    Set Plans = result
End Function

Sub Main()
    Dim items, plan, used, bounds, number, coordinates, record, center, values, perps, i, pairs, pair
    Check Abs(GOST_ARROW_HEAD_LENGTH_M - 0.0025) < 0.0000001, "Arrow length"
    Check Abs(GOST_WITNESS_EXTENSION_M - 0.002) < 0.0000001, "Witness overrun"
    Check COORDINATE_LANE_PITCH_M >= 0.007, "Row pitch"
    Check COORDINATE_BASE_OFFSET_M >= 0.01, "Contour offset"
    Check Not UseCoordinateDots(0.00001), "Merged dots on sub-pixel intervals"
    Check UseCoordinateDots(0.001), "Short-chain dot fallback"
    Check Not UseCoordinateDots(0.01), "Ordinary arrowheads"
    Check CoordinateMarkerPadding(0.0001, 0.01) >= GOST_ARROW_LENGTH_M, "Outside arrow footprint"
    Set items = Plans(5, 0.025, 0.02)
    BalanceCoordinateSides items, 0, 0.2, 1, 1
    Check items.Count = 5, "Lost intervals"
    Set items = Plans(4, 0, 0.0002)
    BalanceCoordinateSides items, 0, 0.2, 1, 3
    Set used = CreateObject("Scripting.Dictionary")
    For Each plan In items.Items
        used.Add CStr(plan("side")) & ":" & CStr(plan("lane")), True
    Next
    Check used.Count = 4, "Overlapping arrowheads in one row"
    Set items = Plans(3, 0, 0.0002)
    On Error Resume Next
    BalanceCoordinateSides items, 0, 0.2, 1, 1
    number = Err.Number
    Err.Clear
    On Error GoTo 0
    Check number = vbObjectError + ERR_DIMENSION_LAYOUT, "Overfull band must fail without dropping dimensions"
    Set items = Plans(3, 0.015, 0.001)
    BalanceCoordinateSides items, 0, 0.2, 1, 0
    bounds = UnionDetailBounds(Array(0, 1, 2, 3), Array(-1, 2, 1, 4))
    Check bounds(0) = -1 And bounds(1) = 1 And bounds(2) = 2 And bounds(3) = 4, "Detail extent union"
    Check DetailProfileRadius(Array(0, 0, 0.006, 0.008)) >= 0.005, "Detail profile excludes a reference"
    Set coordinates = CreateObject("Scripting.Dictionary")
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "coordinate", 0.05
    coordinates.Add "first", record
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "coordinate", 0.0501
    coordinates.Add "last", record
    center = CoordinateTextCenter(coordinates, Array("first", "last"), 0.05, 0.0501, 0.009, 0, 0.1)
    Check center - 0.0045 > 0.055 Or center + 0.0045 < 0.045, "Tiny span label crosses its own extension lines"
    Set coordinates = CreateObject("Scripting.Dictionary")
    values = Array(0, 10, 10.15, 20, 20.15, 30)
    perps = Array(Array(0,1000), Array(0), Array(1000), Array(0), Array(1000), Array(0,1000))
    For i = 0 To UBound(values)
        Set record = CreateObject("Scripting.Dictionary")
        record.Add "coordinate", values(i)
        record.Add "candidatePositions", perps(i)
        coordinates.Add CStr(values(i)), record
    Next
    Set pairs = LocalDatumPairs(coordinates)
    Check pairs.Count = coordinates.Count - 1, "Local graph is not a spanning tree"
    Set used = CreateObject("Scripting.Dictionary")
    For Each pair In pairs.Items
        used(CStr(pair("firstRecord")("coordinate"))) = True
        used(CStr(pair("lastRecord")("coordinate"))) = True
        Check ReferencePairCost(pair("firstRecord"), pair("lastRecord")) < 100, "Local graph joins unrelated opposite-end features"
    Next
    Check used.Count = coordinates.Count, "A geometric coordinate was omitted"
    Set items = Plans(12, 0.005, 0.0002)
    On Error Resume Next
    BalanceCoordinateSides items, 0, 0.2, 1, 1
    number = Err.Number
    Err.Clear
    On Error GoTo 0
    Check number <> 0, "Crowded test did not exercise a partial strict assignment"
    ResetCoordinatePlans items
    For Each plan In items.Items
        Check plan("lane") = -1 And plan("side") = plan("preferredSide"), "Stale strict-plan state leaks into compatibility placement"
    Next
    CompatibleBalanceCoordinateSides items, 0, 0.2, 1, 1
    CompatiblePackCoordinateSide items, 0, 0, 0.2, 0.04, 1, 0
    CompatiblePackCoordinateSide items, 1, 0, 0.2, 0.15, 1, 0
    For Each plan In items.Items
        Check plan("lane") >= 0, "Compatibility fallback lost a dimension"
    Next
    WScript.Echo "PASS: GOST markers, spacing, lane capacity, and detail coverage bounds"
End Sub

Main
'''
root = work_path('dimension_layout')
root.mkdir(parents=True, exist_ok=True)
with tempfile.TemporaryDirectory(dir=root) as folder:
    script = Path(folder) / 'layout.vbs'
    script.write_text('Option Explicit\n' + '\n'.join(constants + routines) + '\n' + harness, 'ascii')
    result = subprocess.run([str(Path(os.environ['WINDIR']) / 'System32/cscript.exe'),
                             '//nologo', str(script)], capture_output=True, timeout=15)
    if result.returncode:
        lines = script.read_text('ascii').splitlines()
        location = re.search(rb'\((\d+),', result.stderr)
        line = int(location[1]) if location else 1
        raise AssertionError((result.stdout, result.stderr, list(enumerate(lines[max(0,line-3):line+2], max(1,line-2)))))
    if b'PASS:' not in result.stdout:
        diagnostic = root / 'failed-layout.vbs'
        diagnostic.write_bytes(script.read_bytes())
        raise AssertionError((result.stdout, result.stderr, str(diagnostic)))
    print(result.stdout.decode('ascii').strip())

fixture = root / 'geometry.json'
if fixture.is_file():
    from dimension_geometry_fixture import check_layout
    check_layout({'constants': constants, 'routines': routines}, fixture, root)
