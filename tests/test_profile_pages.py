"""Execute the VBA page allocator and material clipping without starting CAD."""
from pathlib import Path
import re
import subprocess
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'tools'))
from project_paths import work_path
from vba_sources import read_source

source = read_source()
names = ['MinD', 'ProfilePageCell', 'ProfileClipPolygon', 'ProfilePolygonLocation',
         'ProfilePointClose', 'ProfileCoordinateKey', 'ProfilePagePoint',
         'ProfilePageUnion', 'ProfilePageBoundaryBounds', 'ProfilePageRequestBounds',
         'ProfilePageSupport', 'ProfilePagePlanGroups', 'SplitProfilePageGroup',
         'ProfileLabelBaseFits', 'NativeBoundsInside', 'PanelBoxTouchesMaterial',
         'ProfilePointInMaterial', 'ProfileLineCutsMaterial', 'ProfilePageBands', 'ProfileBasesConflict', 'ProfilePageMapBounds']
procedures = []
for name in names:
    body = re.search(r'(?ms)^Private Function ' + name + r'\(.*?^End Function', source)[0]
    body = body.replace('Private ', '', 1)
    body = body.replace(' As New Collection', '')
    body = re.sub(r' As (?:Object|Collection|Variant|Long|Double|Boolean|String)\b', '', body)
    body = body.replace('New Collection', 'New Sequence')
    body = re.sub(r'(?m)^(\s*)Next \w+\s*$', r'\1Next', body)
    body = re.sub(r'(For Each \w+ In (?:points|plan|groups|profileRawLoops|profilePageGroups))\b', r'\1.Items', body)
    body = body.replace('For Each request In group("requests")', 'For Each request In group("requests").Items')
    if name == 'ProfilePagePlanGroups':
        body = body.replace('    Set supports =', '    Set groups = New Sequence\n    Set supports =', 1)
    procedures.append(body)
constants = ['Const ' + m[1] + ' = ' + re.sub(r'(?<=\d)#', '', m[2])
             for m in re.finditer(r'^Private Const (\w+) As \w+ = (.+)$', source, re.M)]
harness = r'''
Class Sequence
    Private data
    Private Sub Class_Initialize()
        Set data = CreateObject("Scripting.Dictionary")
    End Sub
    Public Sub Add(ByVal item)
        data.Add data.Count + 1, item
    End Sub
    Public Property Get Count()
        Count = data.Count
    End Property
    Public Property Get Items()
        Items = data.Items
    End Property
    Public Default Property Get Item(ByVal index)
        If IsObject(data(index)) Then
            Set Item = data(index)
        Else
            Item = data(index)
        End If
    End Property
End Class
Sub Check(ByVal condition, ByVal message)
    If Not condition Then
        Err.Raise 513, , message
    End If
End Sub
Sub Log(ByVal message)
End Sub
Dim area, x, y, rowHeight, box, count, previous, number
area = Array(0.026, 0.1, 0.202, 0.192)
x = area(0)
y = area(3)
rowHeight = 0
For count = 1 To 4
    box = ProfilePageCell(0.07, 0.04, x, y, rowHeight, area)
    Check Not IsEmpty(box), "A fitting cell was rejected"
    Check box(0) >= area(0) And box(2) <= area(2), "A cell exceeds the horizontal frame"
    Check box(1) >= area(1) And box(3) <= area(3), "A cell exceeds the vertical frame"
Next
box = ProfilePageCell(0.07, 0.04, x, y, rowHeight, area)
Check IsEmpty(box), "Overflow must ask for another sheet"
x = area(0)
y = area(3)
rowHeight = 0
box = ProfilePageCell(0.07, 0.04, x, y, rowHeight, area)
Check Not IsEmpty(box), "Continuation sheet was not reusable"
On Error Resume Next
box = ProfilePageCell(1, 1, x, y, rowHeight, area)
number = Err.Number
Err.Clear
On Error GoTo 0
Check number = vbObjectError + ERR_DIMENSION_LAYOUT, "Oversized cells must fail rather than loop"
Dim polygon, clipped, point, a, b, index, signedArea
polygon = Array(Array(0,0),Array(10,0),Array(10,10),Array(6,10),Array(6,4),Array(4,4),Array(4,10),Array(0,10))
clipped = ProfileClipPolygon(polygon, Array(3,3,7,9))
signedArea = 0
For index = 0 To UBound(clipped)
    a = clipped(index)
    b = clipped((index + 1) Mod (UBound(clipped) + 1))
    Check a(0) >= 3 And a(0) <= 7 And a(1) >= 3 And a(1) <= 9, "Invisible material escaped the crop"
    signedArea = signedArea + a(0)*b(1) - b(0)*a(1)
Next
Check Abs(Abs(signedArea)/2 - 14) < 0.000001, "Clipping filled the open U-shaped recess"
Check ProfilePolygonLocation(Array(5,5), clipped) = PROFILE_POINT_OUTSIDE, "A real recess became material"
Check ProfilePolygonLocation(Array(3.5,5), clipped) = PROFILE_POINT_INSIDE, "Visible material was lost"
clipped = ProfileClipPolygon(polygon, Array(20,20,21,21))
Check UBound(clipped) = -1, "An invisible loop must be empty"
Dim panelLoops, labelRegion
box = ProfilePageMapBounds(Array(-2,-1,3,4),Array(1,2),Array(.1,.2),.005)
Check Abs(box(0)-.085)<.000001 And Abs(box(1)-.185)<.000001 And Abs(box(2)-.11)<.000001 And Abs(box(3)-.21)<.000001, "The true crop rectangle was replaced by padded view bounds"
panelLoops = Array(Array(Array(0,0),Array(.01,0),Array(.01,.02),Array(.02,.02),Array(.02,0),Array(.03,0),Array(.03,.03),Array(0,.03)))
labelRegion = Array(-.01,-.01,.04,.04)
Check Not ProfileLabelBaseFits(Array(.006,.004,.009,.012),1,.01,.011,labelRegion,panelLoops), "A clear baseline concealed text crossing a cutout wall"
Check ProfileLabelBaseFits(Array(.006,.004,.009,.012),1,.01,.018,labelRegion,panelLoops), "A centered label with real cutout clearance was rejected"
Check ProfileBasesConflict(Array("view",0,.02,0,.01,"notch-width"),"view",0,.02,.01,.03,"notch-location"), "Adjacent widths and gaps shared one dimension row"
Check Not ProfileBasesConflict(Array("view",0,.02,0,.01,"notch-width"),"view",0,.03,.01,.03,"notch-location"), "Separate dimension rows were rejected"
Check Not ProfileBasesConflict(Array("view",0,.02,0,.01,"notch-width"),"view",0,.02,.01,.03,"notch-width"), "Nonoverlapping widths could not share a row"
Dim profilePageMatrix, profileRawLoops, raw, bounds, plan, groups, group, request, owners, identifier, profilePageGroups, profilePageSplitSerial
profilePageMatrix = Array(1,0,0,0,1,0,0,0,1,0,0,0,1,0,0,0)
Set profileRawLoops = New Sequence
Set raw = New Sequence
For Each point In Array(Array(0,0,0),Array(.004,0,0),Array(.004,.006,0),Array(.006,.006,0),Array(.006,0,0),Array(.01,0,0),Array(.01,.01,0),Array(0,.01,0))
    raw.Add point
Next
profileRawLoops.Add raw
bounds = ProfilePageBoundaryBounds(Array(.004,0,0),Array(.006,0,0))
Check Abs(bounds(3) - .006) < 0.000001, "A gap across an open cut lost its material bridge"
Function MakeRequest(ByVal identifier, ByVal first, ByVal last)
    Dim item, features
    Set item = CreateObject("Scripting.Dictionary")
    Set features = CreateObject("Scripting.Dictionary")
    features.Add identifier, True
    item.Add "key", identifier
    item.Add "panel", True
    item.Add "kind", "notch-width"
    item.Add "role", PROFILE_ROLE_SLOT
    item.Add "features", features
    item.Add "first", first
    item.Add "last", last
    item.Add "value", Abs(last(0) - first(0))
    Set MakeRequest = item
End Function
Set plan = New Sequence
plan.Add MakeRequest("one",Array(.004,.006,0),Array(.006,.006,0))
plan.Add MakeRequest("two",Array(0,.01,0),Array(.002,.01,0))
plan.Add MakeRequest("three",Array(0,.006,0),Array(.002,.006,0))
Set groups = ProfilePagePlanGroups(plan,12)
Check groups.Count = 2, "Independent supporting edges were merged"
Set group = groups(1)
bounds = ProfilePageBands(group,.021)
Check bounds(0) = PROFILE_PAGE_CONTEXT_BAND_M And bounds(1) = .021, "Horizontal dimensions reserved unused side lanes"
Set profilePageGroups = groups
Set group = groups(1)
group.Add "view", "crowded"
group.Add "number", 1
Set group = groups(2)
group.Add "view", "stable"
group.Add "number", 2
profilePageSplitSerial = 0
Check SplitProfilePageGroup("crowded"), "A crowded group was not partitioned"
Set groups = ProfilePagePlanGroups(plan,12)
Check groups.Count = 3, "A local partition was merged again on rebuild"
Check SplitProfilePageGroup("stable"), "A single-owner group must reserve additional clearance"
Check SplitProfilePageGroup("stable"), "A second bounded clearance retry was rejected"
Check Not SplitProfilePageGroup("stable"), "A single-owner group must stop instead of retrying forever"
Set groups = ProfilePagePlanGroups(plan,1)
Check groups.Count = 3, "Crowded groups were not split"
Set owners = CreateObject("Scripting.Dictionary")
For Each group In groups.Items
    For Each request In group("requests").Items
        identifier = request("key")
        Check Not owners.Exists(identifier), "A split duplicated a dimension owner"
        owners.Add identifier, True
    Next
Next
Check owners.Count = plan.Count, "A split lost a planned dimension"
WScript.Echo "Profile page allocation and clipped-material tests passed"
'''
directory = work_path('profile_page_tests')
directory.mkdir(parents=True, exist_ok=True)
script = directory / 'allocator.vbs'
script.write_text('\n'.join(constants + procedures) + '\n' + harness, encoding='utf-16')
result = subprocess.run(['cscript.exe', '//nologo', str(script)], capture_output=True,
                        text=True, timeout=30)
print(result.stdout, end='')
print(result.stderr, end='', file=sys.stderr)
if result.returncode:
    raise SystemExit(result.returncode)
assert 'tests passed' in result.stdout
