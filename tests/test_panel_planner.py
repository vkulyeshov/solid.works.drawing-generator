"""Exercise the actual VBA panel recognizer with independent and native contours."""
from pathlib import Path
import argparse
import re
import subprocess
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from project_paths import work_path
from vba_sources import read_source
from panel_fixture import boundary

source = read_source()
names = ['MinD', 'ProfileSegmentAxis', 'ProfileNotch', 'ProfileCoordinateKey', 'ProfileRoot', 'ProfilePointClose', 'ProfileTopologyPointClose',
         'PanelNode', 'PanelDistance', 'PanelCleanFrame', 'PanelExcursion',
         'PanelStripExcursions', 'PanelAddRequest', 'PanelFrameDimensions', 'PanelOpeningDimensions', 'PlanPanelDimensions',
         'ProfilePageUnion', 'CompactGroup', 'CompactFrameBounds',
         'CompactMajorCuts', 'CompactFeatureFamily', 'CompactPlanGroups']
functions = []
for name in names:
    body = re.search(r'(?ms)^Private (?:Function|Sub) ' + name + r'\(.*?^End (?:Function|Sub)', source)[0]
    body = body.replace('Private ', '', 1)
    body = re.sub(r' As New Collection\b', '', body)
    body = re.sub(r' As (?:Object|Collection|Variant|Long|Double|Boolean|String)\b', '', body)
    body = body.replace('Optional ByVal overall = False', 'ByVal overall')
    body = re.sub(r'(?m)^(\s*)Next \w+\s*$', r'\1Next', body)
    body = body.replace('New Collection', 'CreateObject("Scripting.Dictionary")')
    for collection in ['outer', 'features', 'frame', 'profilePanelFrame', 'cuts', 'profilePagePlan', 'openings', 'points', 'result']:
        body = re.sub(r'(For Each \w+ In ' + collection + r')\b', r'\1.Items', body)
    for collection, expression in [('raw', 'node'), ('nextPoints', 'points(offset)'),
                                   ('features', 'feature'), ('requests', 'request')]:
        body = body.replace(f'{collection}.Add {expression}', f'{collection}.Add {collection}.Count + 1, {expression}')
    body = re.sub(r'(?m)^(\s*)raw.Add (a|b|c|last)$', r'\1raw.Add raw.Count + 1, \2', body)
    body = body.replace('points.Remove index', 'RemoveNode points, index')
    if name == 'PanelExcursion':
        body = body.replace('    Set a =', '    Set PanelExcursion = Nothing\n    Set raw = CreateObject("Scripting.Dictionary")\n    Set a =', 1)
    if name == 'PlanPanelDimensions':
        body = body.replace('    For Each raw In openings.Items', '    Set PlanPanelDimensions = Nothing\n    Set result = CreateObject("Scripting.Dictionary")\n    Set features = CreateObject("Scripting.Dictionary")\n    For Each raw In openings.Items', 1)
        # VBScript has no optional arguments; make every actual call explicit.
    if name == 'CompactGroup':
        body = body.replace('    Set result =', '    Set requests = CreateObject("Scripting.Dictionary")\n    Set result =', 1)
    if name in ['CompactMajorCuts', 'CompactPlanGroups']:
        body = body.replace('    Set ', '    Set result = CreateObject("Scripting.Dictionary")\n    Set ', 1)
    body = body.replace('result.Add group', 'result.Add result.Count + 1, group')
    body = body.replace('group("requests").Add request', 'group("requests").Add group("requests").Count + 1, request')
    body = body.replace('For Each request In group("requests")', 'For Each request In group("requests").Items')
    body = re.sub(r'(?<=\d)#', '', body)
    body = body.replace('Mid$(', 'Mid(')
    body = re.sub(r'(?m)^(\s*PanelAddRequest .*?, (?:True|False))$',
                  lambda match: match[0] if re.search(r', (?:True|False), (?:True|False)$', match[0]) else match[0] + ', False', body)
    functions.append(body)

constants = ['Const ' + m[1] + ' = ' + re.sub(r'(?<=\d)#', '', m[2])
             for m in re.finditer(r'^Private Const (\w+) As \w+ = (.+)$', source, re.M)]
harness = r'''
Sub Log(ByVal message)
    WScript.Echo message
End Sub
Function ProfilePagePoint(ByVal point)
    ProfilePagePoint = Array(point(0), point(1))
End Function
Sub Check(ByVal condition, ByVal message)
    If Not condition Then
        Err.Raise 513, , message
    End If
End Sub
Sub VerifyTopologyTolerance
    Dim first, nearby, coincident
    first = Array(0.0, 0.0, 0.0)
    nearby = Array(0.000001, 0.0, 0.0)
    coincident = Array(0.0000000001, 0.0, 0.0)
    Check ProfilePointClose(first, nearby), "Equivalent dimension coordinates should retain their tolerance"
    Check Not ProfileTopologyPointClose(first, nearby), "Distinct narrow-slot vertices must not be joined"
    Check ProfileTopologyPointClose(first, coincident), "Coincident model endpoints must still connect"
End Sub
Sub RemoveNode(ByVal points, ByVal index)
    Dim offset
    For offset = index To points.Count - 1
        Set points(offset) = points(offset + 1)
    Next
    points.Remove points.Count
End Sub
Function Nodes(ByVal values)
    Dim result, value, node
    Set result = CreateObject("Scripting.Dictionary")
    For Each value In values
        Set node = CreateObject("Scripting.Dictionary")
        node.Add "x", value(0)
        node.Add "y", value(1)
        node.Add "point", Array(value(0), value(1), 0)
        result.Add result.Count + 1, node
    Next
    Set Nodes = result
End Function
Sub Verify(ByVal points)
    Dim openings, plan, request, cuts, corners, overall, depth, groups, group, retained, major, count, slots
    Set openings = CreateObject("Scripting.Dictionary")
    Set plan = PlanPanelDimensions(points, openings)
    Check Not plan Is Nothing, "A supported clipped panel was not recognized"
    cuts = 0
    corners = 0
    overall = 0
    slots = 0
    For Each request In plan.Items
        WScript.Echo request("kind") & vbTab & request("axis") & vbTab & request("value") & vbTab & request("features").Count & vbTab & request("role")
        If request("kind") = "major-cut-width" Then
            cuts = cuts + 1
        End If
        If request("kind") = "corner-width" Then
            corners = corners + 1
        End If
        If request("overall") Then
            overall = overall + 1
        End If
        If request("kind") = "notch-width" Then
            slots = slots + 1
        End If
    Next
    Check overall = 2, "The frame/local-feature plan lost its overall dimensions"
    Set profilePagePlan = plan
    Set major = CompactMajorCuts()
    If major.Count = COMPACT_TARGET_SHEETS Then
        Set groups = CompactPlanGroups(major)
        Set retained = CreateObject("Scripting.Dictionary")
        For Each group In groups.Items
            Check group("page") >= 1 And group("page") <= 2, "A compact family created an extra page"
            For Each request In group("requests").Items
                Check Not retained.Exists(request("key")), "A compact measurement has two owners"
                retained.Add request("key"), True
            Next
        Next
        count = 0
        For Each request In plan.Items
            If request("role") = PROFILE_ROLE_SLOT Then
                count = count + 1
                Check retained.Exists(request("key")), "A compact family dropped a measurement"
            End If
        Next
        Check retained.Count = count, "Compact coverage differs from the geometry plan"
        WScript.Echo "COMPACT CHECK families=" & groups.Count & "; retained=" & count & "; sheets=2"
    End If
End Sub
Sub VerifyOpening()
    Dim openings, plan, request, sizes, locations
    Set openings = CreateObject("Scripting.Dictionary")
    openings.Add 1, Nodes(Array(Array(0.02, 0.03), Array(0.03, 0.03), Array(0.03, 0.05), Array(0.02, 0.05)))
    Set plan = PlanPanelDimensions(Nodes(Array(Array(0, 0), Array(0.2, 0), Array(0.2, 0.1), Array(0, 0.1))), openings)
    Check Not plan Is Nothing, "A rectangular opening was not recognized"
    sizes = 0
    locations = 0
    For Each request In plan.Items
        If request("kind") = "opening-size" Then
            sizes = sizes + 1
        ElseIf request("kind") = "opening-location" Then
            locations = locations + 1
        End If
    Next
    Check sizes = 2 And locations = 2 And plan.Count = 6, "Opening sizes or datum locations were lost"
    WScript.Echo "OPENING CHECK sizes=2; locations=2; dimensions=6"
End Sub
'''

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--geometry', type=Path)
parser.add_argument('--expect-rightside', action='store_true', help='Verify the captured RightSide feature sizes and counts')
args = parser.parse_args()
fixtures = [
    [(0, 0), (.2, 0), (.2, .1), (.3, .1), (.3, 0), (.9, 0), (.9, .3), (.8, .4), (.1, .4), (0, .3)],
    [(0, 0), (.2, 0), (.2, .1), (.3, .1), (.3, 0), (.5, 0), (.5, .0083), (.52, .0083), (.52, 0),
     (.6, 0), (.6, .00835), (.62, .00835), (.62, 0), (.9, 0), (.9, .3), (.8, .4), (.1, .4), (0, .3)],
    [(0, 0), (.02, 0), (.02, .03), (.03, .03), (.03, 0), (.09, 0), (.09, .008),
     (.1, .008), (.1, 0), (.2, 0), (.2, .1), (0, .1)],
]
if args.geometry:
    fixtures.append(boundary(args.geometry))
checks = []
for points in fixtures:
    literal = 'Array(' + ','.join('Array(' + ','.join(format(v, '.15g') for v in point[:2]) + ')' for point in points) + ')'
    checks.append('Verify Nodes(' + literal + ')')
directory = work_path('rightside_roles/planner')
directory.mkdir(parents=True, exist_ok=True)
script = directory / 'panel-planner.vbs'
script.write_text('Option Explicit\nDim profilePanelFrame, profilePanelFeatures, profilePanelOpenings, profilePagePlan\n' + '\n'.join(constants + functions) + harness + '\n'.join(checks) + '\nVerifyOpening\nVerifyTopologyTolerance\n', encoding='utf-16')
result = subprocess.run(['cscript.exe', '//nologo', str(script)], capture_output=True)
output = result.stdout.decode('oem', errors='replace') + result.stderr.decode('oem', errors='replace')
(directory / 'planner-output.txt').write_text(output, encoding='utf-8')
print(output)
assert result.returncode == 0, result.returncode
assert 'PANEL PLAN' in output
cases = []
for line in output.splitlines():
    if line.startswith('PANEL PLAN'):
        cases.append([])
    elif '\t' in line:
        kind, axis, value, count, role = line.split('\t')
        cases[-1].append((kind, int(axis), float(value.replace(',', '.')), int(count), role))
assert len(cases) == len(fixtures) + 1, 'A fixture did not reach its completed plan'
assert 'OPENING CHECK sizes=2; locations=2; dimensions=6' in output
depths = [row[2] for row in cases[1] if row[0] == 'notch-depth' and row[2] < .02]
assert len(depths) == 2 and abs(depths[0] - depths[1]) > 0.00004, 'A 0.05 mm depth difference was merged'
if args.expect_rightside:
    assert args.geometry, '--expect-rightside requires the captured geometry'
    rows = cases[len(fixtures) - 1]
    for kind, value, count in [('overall-width', 1.3, 1), ('overall-height', .348, 1),
                                ('major-cut-width', .14, 2), ('major-cut-depth', .113, 2),
                                ('corner-height', .048, 2), ('notch-depth', .0083, 22),
                                ('notch-width', .082147, 1), ('notch-width', .119968991, 1),
                                ('notch-depth', .004064706, 4), ('notch-depth', .002738206, 4),
                                ('notch-location', .3588, 1)]:
        matches = [row for row in rows if row[0] == kind and abs(row[2] - value) < 0.000005]
        assert len(matches) == 1 and matches[0][3] == count, (kind, value, matches)
    assert all(row[4] == 'outline' for row in rows if row[0].startswith(('overall-', 'major-', 'corner-')))
    assert all(row[4] == 'slot' for row in rows if row[0].startswith('notch-'))
    assert 'COMPACT CHECK families=9; retained=64; sheets=2' in output, 'The native panel no longer fits its two-page feature plan'
print('PASS: clipped-panel recognition and native-reference dimension planning')
