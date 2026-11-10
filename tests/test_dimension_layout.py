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
from vba_sources import read_source

source = read_source()
names = ['UseCoordinateDots', 'CoordinateMarkerPadding', 'CoordinateTextCenter', 'FindCoordinateTextCenter',
         'AddForbiddenCoordinate', 'TryCoordinatePlanRow', 'BalanceCoordinateSides',
         'UnionDetailBounds', 'DetailProfileRadius', 'MinD', 'SortedCoordinateKeys',
         'ReferencePairCost', 'LocalNeighborBands', 'LocalDatumPairs', 'ResetCoordinatePlans',
         'CompatibleRequiredCoordinateLanes', 'CompatibleBalanceCoordinateSides',
         'CompatiblePackCoordinateSide', 'CoordinateRegionEdge', 'NativeBoundsInside',
         'NativeBoxesOverlap', 'NativeLeaderCrossesBox', 'NativeCandidateLeaderHitsBox', 'NativeGridIndex', 'NativeLabelRowCost', 'PanelLabelPriority', 'NativeIntervalKey',
         'NativePointKey', 'NativeReflectedPoint', 'NativePairKey', 'NativeDirectionKey',
         'NativeTextBounds', 'NativeExpandBounds', 'NativeProtectedSpan', 'NativeFixedLinearSegment', 'OppositeViewTransform', 'SixPrimaryOrientation', 'HolePatternShape', 'QuarterTurnPoint', 'OpenBoreSlotPlan', 'ProfileSegmentAxis', 'ProfileNotch', 'ProfileRoot', 'ProfilePolygonLocation', 'ProfilePointInMaterial', 'ProfileLineCutsMaterial', 'ProfileBaseClear', 'ProfileReferenceWitness', 'PanelMaterialBounds', 'PanelReferenceWitness', 'PanelBoxTouchesMaterial', 'PanelAlignedOffset']
routines = []
for name in names:
    body = re.search(r'(?ms)^Private (Function|Sub) ' + name + r'\(.*?^End (?:Function|Sub)', source)[0]
    body = re.sub(r'^Private ', '', body)
    if name == 'NativeTextBounds':
        body = body.replace(', Optional ByVal owner As Object = Nothing', '')
        body = body.replace('    Dim result As Variant,', '    Dim owner\n    Set owner = Nothing\n    Dim result As Variant,', 1)
    if name == 'LocalDatumPairs':
        body = body.replace('Dim result As New Collection,', 'Dim result,')
        body = body.replace('    count = coordinates.Count', '    Set result = CreateObject("Scripting.Dictionary")\n    count = coordinates.Count')
        body = body.replace('        result.Add pair', '        result.Add result.Count, pair')
    if name == 'ReferencePairCost':
        body = body.replace('Set a =', 'a =').replace('Set b =', 'b =')
    body = re.sub(r' As (?:Object|Variant|Long|Double|Boolean|String|Collection)\b', '', body)
    body = body.replace(' As SldWorks.DisplayData', '').replace('Trim$(', 'Trim(')
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

Sub CheckGridOrder()
    Dim count, preferred, center, order, index, previousDistance, seen, y, height, anchor, edge, bound
    For count = 1 To 25
        For preferred = -2 To count + 1
            center = preferred
            If center < 0 Then
                center = 0
            End If
            If center >= count Then
                center = count - 1
            End If
            previousDistance = -1
            Set seen = CreateObject("Scripting.Dictionary")
            For order = 0 To count - 1
                index = NativeGridIndex(order, count, preferred)
                Check index >= 0 And index < count, "Grid index outside its bounds"
                Check Not seen.Exists(index), "Grid position visited twice"
                Check Abs(index - center) >= previousDistance, "Grid did not search nearest positions first"
                seen.Add index, True
                previousDistance = Abs(index - center)
            Next
            Check seen.Count = count, "Grid search lost a candidate"
        Next
    Next
    For y = -20 To 20
        For height = 1 To 7
            For anchor = -20 To 20
                bound = NativeLabelRowCost(y, height, anchor)
                For edge = 0 To height Step height / 2
                    Check bound <= (y + edge - anchor) ^ 2 + 0.000000001, "Row pruning can discard a better label"
                Next
            Next
        Next
    Next
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

Function ProfileNode(ByVal x, ByVal y)
    Dim node
    Set node = CreateObject("Scripting.Dictionary")
    node.Add "x", x
    node.Add "y", y
    Set ProfileNode = node
End Function

Class NativeTextDataStub
    Public Text, Segment
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
        GetTextPositionAtIndex = Array(0, 0, 0)
    End Function
    Public Function GetTextAngleAtIndex(ByVal index)
        GetTextAngleAtIndex = 0
    End Function
    Public Function GetLineCount()
        If IsEmpty(Segment) Then
            GetLineCount = 0
        Else
            GetLineCount = 1
        End If
    End Function
    Public Function GetLineAtIndex2(ByVal index)
        GetLineAtIndex2 = Segment
    End Function
End Class

Sub Main()
    Dim items, plan, used, bounds, number, coordinates, record, center, values, perps, i, pairs, pair, textData, shape, polygon, loops
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
    Check NativeBoundsInside(Array(0.03, 0.09, 0.20, 0.28), Array(0.026, 0.082, 0.202, 0.289)), "In-frame graphics rejected"
    Check Not NativeBoundsInside(Array(0.02, 0.09, 0.20, 0.28), Array(0.026, 0.082, 0.202, 0.289)), "Out-of-frame leader accepted"
    Check NativeBoxesOverlap(Array(0, 0, 2, 2), Array(1, 1, 3, 3), 0), "Overlapping text boxes missed"
    Check Not NativeBoxesOverlap(Array(0, 0, 1, 1), Array(2, 2, 3, 3), 0), "Separate text boxes rejected"
    Check NativeLeaderCrossesBox(0, 0, 4, 4, Array(1, 1, 2, 2)), "Diagonal through text missed"
    Check NativeLeaderCrossesBox(4, 4, 0, 0, Array(1, 1, 2, 2)), "Reversed diagonal through text missed"
    Check Not NativeLeaderCrossesBox(0, 0, 4, 0, Array(1, 1, 2, 2)), "Horizontal leader outside text rejected"
    Check Not NativeLeaderCrossesBox(0, 0, 0, 4, Array(1, 1, 2, 2)), "Vertical leader outside text rejected"
    Check Not NativeCandidateLeaderHitsBox(0,0,Array(2,2,4,3),Array(2.9,1.4,3.1,1.6),True), "An unused alternate attachment blocked a clear panel leader"
    Check NativeCandidateLeaderHitsBox(0,0,Array(2,2,4,3),Array(0.9,0.9,1.1,1.1),True), "The active panel leader was allowed through text"
    Check NativeCandidateLeaderHitsBox(0,0,Array(2,2,4,3),Array(2.9,1.4,3.1,1.6),False), "The ordinary two-path policy changed"
    Check NativeIntervalKey(0.01, 0.02, 0, True) = NativeIntervalKey(-0.02, -0.01, 0, True), "Verified mirror intervals did not group"
    Check NativeIntervalKey(0.01, 0.02, 0, False) <> NativeIntervalKey(-0.02, -0.01, 0, False), "Intervals grouped without symmetry evidence"
    Check NativeIntervalKey(0.01, 0.02, 0, False) <> NativeIntervalKey(0.03, 0.04, 0, False), "Equal lengths at different source coordinates grouped"
    Check NativeIntervalKey(0, 0.001354, 0, True) = NativeIntervalKey(-0.001354, 0, 0, True), "Near-origin mirror intervals did not group"
    Check NativePairKey(Array(1, 2, 3), Array(4, 5, 6)) = NativePairKey(Array(4, 5, 6), Array(1, 2, 3)), "Reference order changes geometric identity"
    shape = OpenBoreSlotPlan(Array(-0.006, 0.0154, 0.006, 0.0234), Array(0, 0, 0), 0.0194)
    Check ProfileNotch(ProfileNode(0.02, 0), ProfileNode(0.02, 0.01), ProfileNode(0.04, 0.01), ProfileNode(0.04, 0), 1), "A rectangular recess was not classified as a slot"
    Check ProfileNotch(ProfileNode(-0.02, 0), ProfileNode(-0.02, 0.01), ProfileNode(-0.04, 0.01), ProfileNode(-0.04, 0), -1), "Mirroring changed the geometric role of a slot"
    Check ProfileNotch(ProfileNode(0, 0.02), ProfileNode(-0.01, 0.02), ProfileNode(-0.01, 0.04), ProfileNode(0, 0.04), 1), "A quarter turn changed the geometric role of a slot"
    Check Not ProfileNotch(ProfileNode(0.02, 0), ProfileNode(0.02, -0.01), ProfileNode(0.04, -0.01), ProfileNode(0.04, 0), 1), "An outward tab was classified as a slot"
    Check Not ProfileNotch(ProfileNode(0.02, 0), ProfileNode(0.02, 0.01), ProfileNode(0.04, 0.01), ProfileNode(0.04, 0.002), 1), "An open step was incorrectly treated as a closed rectangular recess"
    Check Not IsEmpty(shape), "Open circular keyway was not recognized"
    Check Abs(shape(5) - 0.012) < 0.000001, "Open keyway width is incorrect"
    Check Abs(shape(6) - (0.0234 - Sqr(0.0194 ^ 2 - 0.006 ^ 2))) < 0.000001, "Keyway depth still includes the rectangle inside the bore"
    shape = OpenBoreSlotPlan(Array(-0.006, -0.0234, 0.006, -0.0154), Array(0, 0, 0), 0.0194)
    Check Not IsEmpty(shape) And shape(2) < shape(1), "Negative radial keyway direction was lost"
    shape = OpenBoreSlotPlan(Array(0.002, 0.0154, 0.014, 0.0234), Array(0, 0, 0), 0.0194)
    Check IsEmpty(shape), "An off-center rectangle was incorrectly recognized as a radial keyway"
    values = OppositeViewTransform(Array(1, 0, 0, 0, 1, 0, 0, 0, 1, 0.1, 0.2, 0.3, 0.4, 0, 0, 0))
    shape = HolePatternShape(Array(Array(0, -0.035355339, -0.035355339), Array(0, -0.035355339, 0.035355339), Array(0, 0.035355339, -0.035355339), Array(0, 0.035355339, 0.035355339)), Array(1, 0, 0))
    Check shape(PATTERN_REGULAR_INDEX) And Abs(shape(PATTERN_RADIUS_INDEX) - 0.05) < 0.000001, "A square bolt pattern was not recognized as a 100 mm pitch circle"
    shape = HolePatternShape(Array(Array(0.042, 0, -0.034), Array(0.07, 0, -0.034), Array(0.042, 0, 0.034), Array(0.07, 0, 0.034)), Array(0, 1, 0))
    Check Not shape(PATTERN_REGULAR_INDEX), "An unequal rectangular pattern was incorrectly called equally spaced on a circle"
    shape = QuarterTurnPoint(Array(0.03, 0.02, 0), Array(0.01, 0.01, 0))
    Check Abs(shape(0)) < 0.000001 And Abs(shape(1) - 0.03) < 0.000001, "Local sketch rotation did not preserve the feature center"
    Check values(0) = -1 And values(4) = 1 And values(8) = -1, "The opposite view must reverse right and depth while preserving up"
    Check values(9) = 0 And values(10) = 0 And values(11) = 0 And values(12) = 1, "Drawing translation and scale leaked into the named model orientation"
    values = OppositeViewTransform(values)
    Check values(0) = 1 And values(4) = 1 And values(8) = 1, "Opposite projection is not a reversible proper rotation"
    Check SixPrimaryOrientation("*Right", False) = "*Top" And SixPrimaryOrientation("*Right", True) = "*Front", "Six-view primary directions repeat the principal view"
    bounds = NativeProtectedSpan(Array(0, 0), Array(0.01, 0), GOST_ARROW_LENGTH_M)
    Check Abs(bounds(0) + GOST_ARROW_LENGTH_M) < MIN_PROJECTED_EXTENT_M And Abs(bounds(2) - 0.01 - GOST_ARROW_LENGTH_M) < MIN_PROJECTED_EXTENT_M, "Future outside arrow tails were not reserved"
    bounds = NativeProtectedSpan(Array(0.076, 0.124), Array(0.076, 0.1516), GOST_WITNESS_EXTENSION_M)
    Check NativeLeaderCrossesBox(bounds(0), bounds(1), bounds(2), bounds(3), Array(0.072, 0.152, 0.0834, 0.1555)), "Future witness extension was not reserved"
    Set textData = New NativeTextDataStub
    textData.Text = " <MOD-DIAM> "
    bounds = NativeTextBounds(textData)
    Check bounds(2) - bounds(0) < 0.01, "A native diameter symbol was counted as a ten-character token"
    textData.Text = " 6 "
    textData.Segment = Array(0, 0, -1, -1, 0, 0, 0, 0.0093, 0, 0)
    bounds = NativeTextBounds(textData)
    Check Abs(bounds(2) - bounds(0) - 0.0093) < 0.000001, "Native leader shoulder width was ignored"
    textData.Text = " 6,39 "
    bounds = NativeTextBounds(textData)
    Check bounds(2) - bounds(0) > 0.0093, "A short shoulder truncated the longer text bounds"
    polygon = Array(Array(0,0),Array(0.004,0),Array(0.004,0.004),Array(0.006,0.004),Array(0.006,0),Array(0.01,0),Array(0.01,0.01),Array(0,0.01))
    loops = Array(polygon)
    Check Not PanelBoxTouchesMaterial(Array(0.0042,0.001,0.0058,0.003),loops), "A label in an open cutout was rejected as bounding-box material"
    Check PanelBoxTouchesMaterial(Array(0.0035,0.001,0.0045,0.003),loops), "A label crossing a cutout wall was accepted"
    Check PanelBoxTouchesMaterial(Array(-0.001,-0.001,0.011,0.011),loops), "Material enclosed by a label was missed"
    Check Not PanelBoxTouchesMaterial(Array(0.0045,-0.005,0.0055,-0.001),loops), "An external panel label was rejected"
    Check ProfilePointInMaterial(Array(0.002,0.002),loops), "Profile solid land was missed"
    Check Not ProfilePointInMaterial(Array(0.005,0.002),loops), "An open notch was classified as material"
    Check Not ProfilePointInMaterial(Array(0.004,0.002),loops), "A boundary was classified as solid interior"
    Check Not ProfileLineCutsMaterial(Array(0.004,0.004),Array(0.004,-0.003),loops), "A lower witness along the slot wall was rejected"
    Check ProfileLineCutsMaterial(Array(0.004,0.004),Array(0.004,0.013),loops), "An upper witness crossing the part was accepted"
    Check ProfileLineCutsMaterial(Array(0.003,0.002),Array(0.007,0.002),loops), "A line with its midpoint in a void concealed two material crossings"
    Check Not ProfileLineCutsMaterial(Array(0.005,0.002),Array(0.005,-0.003),loops), "A leader through the open mouth was rejected"
    Check ProfileLineCutsMaterial(Array(0.005,0.002),Array(0.009,-0.003),loops), "A diagonal leader cutting the slot wall was accepted"
    Check ProfileLineCutsMaterial(Array(0.004,0.00400000000001),Array(0.007,0.00400000000001),loops), "Native floor rounding concealed an overrun beyond the opposite wall"
    Check Not ProfileBaseClear(Array(0.004,0.004),Array(0.004,0),1,0.0055,loops), "Witness overrun into the opposite slot wall was accepted"
    Check ProfileReferenceWitness(Array(0.004,0.004),Array(-0.01,0.004),Array(0.004,0.004),Array(0.004,0),1), "External extension at its model coordinate was rejected"
    Check Not ProfileReferenceWitness(Array(0.004,0.002),Array(-0.01,0.002),Array(0.004,0.004),Array(0.004,0),1), "An arbitrary horizontal leader was treated as an extension"
    Check Not ProfileReferenceWitness(Array(0.004,0.004),Array(-0.01,0.003),Array(0.004,0.004),Array(0.004,0),1), "A diagonal leader was treated as an extension"
    Check Not ProfileReferenceWitness(Array(0.004,0.004),Array(0.004,0),Array(0.004,0.004),Array(0.004,0),1), "A dimension baseline was treated as an extension"
    Check PanelReferenceWitness(Array(0.0012,-0.0009),Array(0.004,-0.003),Array(0,0),Array(0.03,0.04)), "An aligned extension at its native reference was rejected"
    Check Not PanelReferenceWitness(Array(0.015,0.02),Array(0.019,0.017),Array(0,0),Array(0.03,0.04)), "An aligned leader through material was treated as an extension"
    Check Not PanelReferenceWitness(Array(0,0),Array(0.03,0.04),Array(0,0),Array(0.03,0.04)), "An aligned baseline was treated as an extension"
    Check NativeFixedLinearSegment(Array(-0.005,0,0.025,0),Array(0,0),Array(0.02,0)), "A fixed measurement baseline was lost"
    Check NativeFixedLinearSegment(Array(0,0.03,0,-0.0015),Array(0,0),Array(0.02,0)), "A model-linked witness was lost"
    Check Not NativeFixedLinearSegment(Array(0.01,0,0.03,-0.01),Array(0,0),Array(0.02,0)), "An old movable leader remained a fixed obstacle"
    Check Not NativeFixedLinearSegment(Array(0.03,-0.01,0.04,-0.01),Array(0,0),Array(0.02,0)), "An old text shoulder remained a fixed obstacle"
    Check Not NativeFixedLinearSegment(Array(0.01,0,0.01,-0.01),Array(0,0),Array(0.02,0)), "A perpendicular text leader was confused with an endpoint witness"
    polygon = Array(Array(-1,-0.5),Array(0.8,-0.5),Array(1,-0.3),Array(1,0.5),Array(-1,0.5))
    Check Abs(PanelAlignedOffset(Array(0.8,-0.5),1/Sqr(2),-1/Sqr(2),polygon)-GOST_OBJECT_OFFSET_M)<0.000000001, "An outward chamfer dimension was pushed across the whole panel"
    Check PanelAlignedOffset(Array(0.8,-0.5),-1/Sqr(2),1/Sqr(2),polygon)>1, "The inward chamfer normal was not distinguished"
    Check PanelLabelPriority(0.0001433,False,False)+0.007>PanelLabelPriority(0.000406,False,False)+0.013, "A repeated wide label consumed the smallest measurement corridor first"
    Check PanelLabelPriority(0.0001433,False,False)>PanelLabelPriority(0.012,False,False), "A short dimension did not receive placement priority"
    Check PanelLabelPriority(0.004,True,False)>PanelLabelPriority(0.0001433,False,False), "An inclined corner lost its constrained placement priority"
    polygon = Array(Array(0.00498,0),Array(0.005,0),Array(0.005,0.01),Array(0.00498,0.01))
    Check ProfileLineCutsMaterial(Array(0,0.005),Array(0.02,0.005),Array(polygon)), "A narrow land between sampling positions was missed"
    WScript.Echo "PASS: GOST markers, spacing, lane capacity, detail bounds, leaders, and exact profile-material intersections"
End Sub

CheckGridOrder
Main
'''
root = work_path('dimension_layout')
root.mkdir(parents=True, exist_ok=True)
with tempfile.TemporaryDirectory(dir=root) as folder:
    script = Path(folder) / 'layout.vbs'
    script.write_text('Option Explicit\n' + '\n'.join(constants + routines) + '\n' + harness, 'utf-16')
    result = subprocess.run([str(Path(os.environ['WINDIR']) / 'System32/cscript.exe'),
                             '//nologo', str(script)], capture_output=True, timeout=15)
    if result.returncode:
        lines = script.read_text('utf-16').splitlines()
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
