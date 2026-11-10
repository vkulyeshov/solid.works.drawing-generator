Attribute VB_Name = "profile_routing"
Option Explicit

Private Function ProfilePolygonLocation(ByVal point As Variant, ByVal polygon As Variant) As Long
    Dim index As Long, a As Variant, b As Variant, cross As Double, length As Double, inside As Boolean
    For index = LBound(polygon) To UBound(polygon)
        a = polygon(index)
        b = polygon((index + 1) Mod (UBound(polygon) + 1))
        length = Sqr((b(0) - a(0)) ^ 2 + (b(1) - a(1)) ^ 2)
        cross = (point(0) - a(0)) * (b(1) - a(1)) - (point(1) - a(1)) * (b(0) - a(0))
        If Abs(cross) <= PROFILE_MATERIAL_TOLERANCE_M * length Then
            If point(0) >= MinD(a(0), b(0)) - PROFILE_MATERIAL_TOLERANCE_M And point(0) <= -MinD(-a(0), -b(0)) + PROFILE_MATERIAL_TOLERANCE_M And point(1) >= MinD(a(1), b(1)) - PROFILE_MATERIAL_TOLERANCE_M And point(1) <= -MinD(-a(1), -b(1)) + PROFILE_MATERIAL_TOLERANCE_M Then
                ProfilePolygonLocation = PROFILE_POINT_BOUNDARY
                Exit Function
            End If
        End If
        If (a(1) > point(1)) <> (b(1) > point(1)) Then
            If point(0) < a(0) + (point(1) - a(1)) * (b(0) - a(0)) / (b(1) - a(1)) Then
                inside = Not inside
            End If
        End If
    Next index
    If inside Then
        ProfilePolygonLocation = PROFILE_POINT_INSIDE
    End If
End Function

Private Function ProfilePointInMaterial(ByVal point As Variant, ByVal loops As Variant) As Boolean
    Dim index As Long
    If ProfilePolygonLocation(point, loops(0)) <> PROFILE_POINT_INSIDE Then
        Exit Function
    End If
    For index = 1 To UBound(loops)
        If ProfilePolygonLocation(point, loops(index)) <> PROFILE_POINT_OUTSIDE Then
            Exit Function
        End If
    Next index
    ProfilePointInMaterial = True
End Function

Private Function ProfileLineCutsMaterial(ByVal first As Variant, ByVal last As Variant, ByVal loops As Variant) As Boolean
    Dim parameters() As Double, count As Long, polygon As Variant, a As Variant, b As Variant, index As Long, other As Long
    Dim dx As Double, dy As Double, ex As Double, ey As Double, denominator As Double, t As Double, u As Double, swap As Double
    dx = last(0) - first(0)
    dy = last(1) - first(1)
    If Sqr(dx * dx + dy * dy) <= PROFILE_MATERIAL_TOLERANCE_M Then
        Exit Function
    End If
    ReDim parameters(1)
    parameters(1) = 1
    count = 2
    For Each polygon In loops
        For index = LBound(polygon) To UBound(polygon)
            a = polygon(index)
            b = polygon((index + 1) Mod (UBound(polygon) + 1))
            ex = b(0) - a(0)
            ey = b(1) - a(1)
            denominator = dx * ey - dy * ex
            If Abs(denominator) > PROFILE_MATERIAL_TOLERANCE_M ^ 2 Then
                t = ((a(0) - first(0)) * ey - (a(1) - first(1)) * ex) / denominator
                u = ((a(0) - first(0)) * dy - (a(1) - first(1)) * dx) / denominator
                If t > 0 And t < 1 And u >= -DIRECTION_ZERO_TOLERANCE And u <= 1 + DIRECTION_ZERO_TOLERANCE Then
                    ReDim Preserve parameters(count)
                    parameters(count) = t
                    count = count + 1
                End If
            End If
        Next index
    Next polygon
    For index = 0 To count - 2
        For other = index + 1 To count - 1
            If parameters(other) < parameters(index) Then
                swap = parameters(index)
                parameters(index) = parameters(other)
                parameters(other) = swap
            End If
        Next other
    Next index
    For index = 0 To count - 2
        If parameters(index + 1) - parameters(index) > DIRECTION_ZERO_TOLERANCE Then
            t = (parameters(index) + parameters(index + 1)) / 2
            If ProfilePointInMaterial(Array(first(0) + dx * t, first(1) + dy * t), loops) Then
                ProfileLineCutsMaterial = True
                Exit Function
            End If
        End If
    Next index
End Function

Private Function ProfileOwnsDimension(ByVal display As SldWorks.DisplayDimension) As Boolean
    If Not profileDimensionRequests Is Nothing Then
        ProfileOwnsDimension = profileDimensionRequests.Exists(display.GetNameForSelection)
    End If
End Function

Private Function ProfileLocalBase(ByVal display As SldWorks.DisplayDimension) As Boolean
    Dim request As Object
    If ProfileOwnsDimension(display) Then
        Set request = profileDimensionRequests(display.GetNameForSelection)
        If request.Exists("localBase") Then
            ProfileLocalBase = CBool(request("localBase"))
        End If
    End If
End Function

Private Sub CacheProfileContours(ByVal views As Collection)
    Dim item As Variant, view As SldWorks.View, loops() As Variant, polygon() As Variant, raw As Collection
    Dim loopIndex As Long, index As Long, point As Variant, audit As Object
    If profileGeometryViews Is Nothing Then
        Exit Sub
    End If
    For Each item In views
        Set view = item
        If profileGeometryViews.Exists(view.Name) Then
            ReDim loops(profileRawLoops.Count - 1)
            Set audit = fs.OpenTextFile(fs.BuildPath(runDir, PROFILE_CONTOURS_FILENAME), FILE_FOR_APPENDING, True, FILE_UNICODE)
            For loopIndex = 1 To profileRawLoops.Count
                Set raw = profileRawLoops(loopIndex)
                ReDim polygon(raw.Count - 1)
                For index = 1 To raw.Count
                    point = PatternSheetPoint(view, raw(index))
                    polygon(index - 1) = point
                    audit.WriteLine CStr(jobNumber) & vbTab & view.Name & vbTab & loopIndex - 1 & vbTab & index - 1 & vbTab & CStr(point(0)) & vbTab & CStr(point(1))
                Next index
                If profilePagedMode Then
                    If profilePageCrops.Exists(view.Name) Then
                        loops(loopIndex - 1) = ProfileClipPolygon(polygon, profilePageCrops(view.Name))
                    Else
                        loops(loopIndex - 1) = polygon
                    End If
                Else
                    loops(loopIndex - 1) = polygon
                End If
            Next loopIndex
            audit.Close
            profileSheetContours(view.Name) = loops
        End If
    Next item
End Sub

Private Function ProfileDimensionCutsMaterial(ByVal display As SldWorks.DisplayDimension, Optional ByVal diagnose As Boolean = False) As Boolean
    Dim data As SldWorks.DisplayData, request As Object, loops As Variant, line As Variant, index As Long
    Dim first As Variant, last As Variant, a As Variant, b As Variant, externalWitness As Boolean
    If Not ProfileOwnsDimension(display) Then
        Exit Function
    End If
    Set request = profileDimensionRequests(display.GetNameForSelection)
    loops = profileSheetContours(CStr(request("view")))
    If request.Exists("externalWitness") Then
        externalWitness = CBool(request("externalWitness"))
    End If
    If externalWitness Then
        a = request("sheetFirst")
        b = request("sheetLast")
    End If
    Set data = display.GetDisplayData
    For index = 0 To data.GetLineCount - 1
        line = data.GetLineAtIndex2(index)
        first = Array(line(DISPLAY_LINE_START_X), line(DISPLAY_LINE_START_Y))
        last = Array(line(DISPLAY_LINE_END_X), line(DISPLAY_LINE_END_Y))
        If ProfileLineCutsMaterial(first, last, loops) Then
            If Not externalWitness Then
                ProfileDimensionCutsMaterial = True
            Else
                If CLng(request("axis")) = PANEL_ALIGNED_AXIS Then
                    If Not PanelReferenceWitness(first, last, a, b) Then
                        ProfileDimensionCutsMaterial = True
                    End If
                ElseIf Not ProfileReferenceWitness(first, last, a, b, CLng(request("axis"))) Then
                    ProfileDimensionCutsMaterial = True
                End If
            End If
            If ProfileDimensionCutsMaterial Then
                If diagnose Then
                    Log "PROFILE MATERIAL COLLISION " & display.GetNameForSelection & "; segment=" & Join(first, ",") & " -> " & Join(last, ",") & "; external witnesses=" & externalWitness
                End If
                Exit Function
            End If
        End If
    Next index
End Function

' Only straight extension lines at the two measured model coordinates may cross material.
' The dimension baseline and the text leader must remain outside the material.
Private Function ProfileReferenceWitness(ByVal first As Variant, ByVal last As Variant, ByVal a As Variant, ByVal b As Variant, ByVal axis As Long) As Boolean
    If Abs(first(axis) - last(axis)) > PROFILE_MATERIAL_TOLERANCE_M Then
        Exit Function
    End If
    ProfileReferenceWitness = Abs(first(axis) - a(axis)) <= PROFILE_MATERIAL_TOLERANCE_M Or Abs(first(axis) - b(axis)) <= PROFILE_MATERIAL_TOLERANCE_M
End Function

Private Function ProfileLabelConnection(ByVal display As SldWorks.DisplayDimension, ByVal box As Variant) As Variant
    If NativePlacementSide(display) = NATIVE_LABEL_SIDE_LEFT Then
        ProfileLabelConnection = Array(box(OUTLINE_RIGHT_INDEX), box(OUTLINE_BOTTOM_INDEX))
    Else
        ProfileLabelConnection = Array(box(OUTLINE_LEFT_INDEX), box(OUTLINE_BOTTOM_INDEX))
    End If
End Function

Private Function ProfileCandidateLeaderClear(ByVal display As SldWorks.DisplayDimension, ByVal x As Double, ByVal y As Double, ByVal box As Variant) As Boolean
    Dim request As Object, loops As Variant
    ProfileCandidateLeaderClear = True
    If Not ProfileLocalBase(display) Then
        Exit Function
    End If
    Set request = profileDimensionRequests(display.GetNameForSelection)
    loops = profileSheetContours(CStr(request("view")))
    ProfileCandidateLeaderClear = Not ProfileLineCutsMaterial(Array(x, y), ProfileLabelConnection(display, box), loops)
End Function

Private Function ProfileBaseClear(ByVal a As Variant, ByVal b As Variant, ByVal axis As Long, ByVal coordinate As Double, ByVal loops As Variant) As Boolean
    Dim first As Variant, last As Variant, witnessA As Variant, witnessB As Variant
    first = Array(a(0), a(1))
    last = Array(b(0), b(1))
    first(1 - axis) = coordinate
    last(1 - axis) = coordinate
    witnessA = first
    witnessB = last
    witnessA(1 - axis) = coordinate + Sgn(coordinate - a(1 - axis)) * GOST_WITNESS_EXTENSION_M
    witnessB(1 - axis) = coordinate + Sgn(coordinate - b(1 - axis)) * GOST_WITNESS_EXTENSION_M
    ProfileBaseClear = Not ProfileLineCutsMaterial(a, witnessA, loops) And Not ProfileLineCutsMaterial(b, witnessB, loops) And Not ProfileLineCutsMaterial(first, last, loops)
End Function

Private Function ProfileLabelBaseFits(ByVal bounds As Variant, ByVal axis As Long, ByVal baseline As Double, ByVal coordinate As Double, ByVal region As Variant, ByVal loops As Variant) As Boolean
    Dim box As Variant, expanded As Variant, offset As Double
    box = Array(bounds(0), bounds(1), bounds(2), bounds(3))
    offset = coordinate - baseline
    box(1 - axis) = box(1 - axis) + offset
    box(3 - axis) = box(3 - axis) + offset
    If Not NativeBoundsInside(box, region) Then
        Exit Function
    End If
    expanded = Array(box(0) - LEADER_GEOMETRY_CLEARANCE_M, box(1) - LEADER_GEOMETRY_CLEARANCE_M, box(2) + LEADER_GEOMETRY_CLEARANCE_M, box(3) + LEADER_GEOMETRY_CLEARANCE_M)
    ProfileLabelBaseFits = Not PanelBoxTouchesMaterial(expanded, loops)
End Function

Private Function ProfileBasesConflict(ByVal placed As Variant, ByVal viewName As String, ByVal axis As Long, ByVal coordinate As Double, ByVal lower As Double, ByVal upper As Double, ByVal kind As String) As Boolean
    If placed(0) <> viewName Or CLng(placed(1)) <> axis Then
        Exit Function
    End If
    If Abs(coordinate - CDbl(placed(2))) >= COORDINATE_LANE_PITCH_M - PROFILE_MATERIAL_TOLERANCE_M Then
        Exit Function
    End If
    ' Feature widths and the gaps locating them need distinct dimension rows,
    ' even when their measured intervals only touch at a shared endpoint.
    If UBound(placed) >= PROFILE_BASE_KIND_INDEX Then
        If (kind = "notch-location") <> (placed(PROFILE_BASE_KIND_INDEX) = "notch-location") Then
            ProfileBasesConflict = True
            Exit Function
        End If
    End If
    ProfileBasesConflict = lower < placed(4) - PROFILE_MATERIAL_TOLERANCE_M And upper > placed(3) + PROFILE_MATERIAL_TOLERANCE_M
End Function

Private Sub PlaceProfileDimensionBase(ByVal d As SldWorks.ModelDoc2, ByVal display As SldWorks.DisplayDimension, ByVal view As SldWorks.View, ByVal region As Variant, ByVal bases As Collection)
    Dim request As Object, a As Variant, b As Variant, outline As Variant, loops As Variant, polygon As Variant
    Dim candidates As New Collection, candidate As Variant, point As Variant, axis As Long, side As Long, bestSide As Long, lane As Long
    Dim coordinate As Double, best As Double, cost As Double, bestCoordinate As Double, bestLocal As Boolean, boundary As Boolean
    Dim sharedDepth As Boolean, clear As Boolean, externalWitness As Boolean, bestExternalWitness As Boolean, audit As Object
    Dim index As Long, other As Long, position As Variant, arrow As Variant, ann As SldWorks.Annotation, data As SldWorks.DisplayData
    Dim levels As Object, levelKeys As Variant
    Dim panel As Boolean, occupied As Boolean, placedBase As Variant, lower As Double, upper As Double
    Dim textBounds As Variant, textBaseline As Double, baseKind As String
    Set request = profileDimensionRequests(display.GetNameForSelection)
    request("view") = view.Name
    If CLng(request("axis")) = PANEL_ALIGNED_AXIS Then
        PlacePanelAlignedBase d, display, view, request, region
        Exit Sub
    End If
    loops = profileSheetContours(view.Name)
    outline = view.GetOutline
    a = PatternSheetPoint(view, request("first"))
    b = PatternSheetPoint(view, request("last"))
    axis = CLng(request("axis"))
    sharedDepth = request("role") = PROFILE_ROLE_SLOT And axis = SCREEN_Y_AXIS_INDEX And request("features").Count > 1
    panel = request.Exists("panel")
    If panel And profilePagedMode Then
        baseKind = CStr(request("kind"))
    End If
    If panel Then
        outline = PanelMaterialBounds(loops(0))
        If request("kind") = "notch-depth" And request("features").Count > 1 Then
            sharedDepth = True
        End If
    End If
    side = NativePlacementSide(display)
    If featureDimensionLanes.Exists(display.GetNameForSelection) Then
        lane = CLng(featureDimensionLanes(display.GetNameForSelection))
    End If
    If axis = SCREEN_X_AXIS_INDEX Then
        candidates.Add Array(outline(1) - GOST_OBJECT_OFFSET_M - lane * COORDINATE_LANE_PITCH_M, NATIVE_LABEL_SIDE_BELOW, False)
        candidates.Add Array(outline(3) + GOST_OBJECT_OFFSET_M + lane * COORDINATE_LANE_PITCH_M, NATIVE_LABEL_SIDE_ABOVE, False)
    Else
        candidates.Add Array(outline(0) - GOST_OBJECT_OFFSET_M, NATIVE_LABEL_SIDE_LEFT, False)
        candidates.Add Array(outline(2) + GOST_OBJECT_OFFSET_M, NATIVE_LABEL_SIDE_RIGHT, False)
    End If
    If panel Then
        For index = 0 To LEADER_BASE_ROW_LIMIT
            If axis = SCREEN_X_AXIS_INDEX Then
                candidates.Add Array(outline(1) - GOST_OBJECT_OFFSET_M - index * COORDINATE_LANE_PITCH_M, NATIVE_LABEL_SIDE_BELOW, False)
                candidates.Add Array(outline(3) + GOST_OBJECT_OFFSET_M + index * COORDINATE_LANE_PITCH_M, NATIVE_LABEL_SIDE_ABOVE, False)
            Else
                candidates.Add Array(outline(0) - GOST_OBJECT_OFFSET_M - index * COORDINATE_LANE_PITCH_M, NATIVE_LABEL_SIDE_LEFT, False)
                candidates.Add Array(outline(2) + GOST_OBJECT_OFFSET_M + index * COORDINATE_LANE_PITCH_M, NATIVE_LABEL_SIDE_RIGHT, False)
            End If
        Next index
    End If
    polygon = loops(0)
    candidates.Add Array(a(1 - axis) + GOST_WITNESS_EXTENSION_M / 2, side, True)
    candidates.Add Array(a(1 - axis) - GOST_WITNESS_EXTENSION_M / 2, side, True)
    ' Material occupancy is constant between adjacent vertex levels. Testing
    ' every pair repeats the same clear strips quadratically on detailed panels.
    Set levels = CreateObject("Scripting.Dictionary")
    For Each polygon In loops
        For index = LBound(polygon) To UBound(polygon)
            point = polygon(index)
            levels(CStr(point(1 - axis))) = point(1 - axis)
        Next index
    Next polygon
    levelKeys = SortedCoordinateKeys(levels)
    For index = LBound(levelKeys) To UBound(levelKeys) - 1
        coordinate = (CDbl(levelKeys(index)) + CDbl(levelKeys(index + 1))) / 2
        candidates.Add Array(coordinate, side, True)
    Next index
    best = SEARCH_INITIAL_LIMIT
    lower = MinD(a(axis), b(axis)) - GOST_ARROW_LENGTH_M
    upper = -MinD(-a(axis), -b(axis)) + GOST_ARROW_LENGTH_M
    If panel Then
        lower = MinD(a(axis), b(axis))
        upper = -MinD(-a(axis), -b(axis))
        display.ArrowSide = SW_DIMENSION_ARROWS_INSIDE
        If upper - lower < 2 * GOST_ARROW_HEAD_LENGTH_M Then
            display.SetArrowHeadStyle2 False, SW_ARROW_DOT, SW_ARROW_DOT
            ApplyGostDimensionStyle display
        End If
    End If
    If panel And profilePagedMode Then
        display.CenterText = True
        d.GraphicsRedraw2
        Set data = display.GetDisplayData
        textBounds = NativeTextBounds(data, display)
        arrow = data.GetArrowHeadAtIndex2(0)
        textBaseline = arrow(1 - axis)
    End If
    For Each candidate In candidates
        coordinate = candidate(0)
        occupied = False
        If panel Then
            occupied = coordinate < region(1 - axis) Or coordinate > region(3 - axis)
            If axis = SCREEN_X_AXIS_INDEX Then
                If coordinate + GOST_ARROW_HEAD_WIDTH_M >= sheetHeight - LEADER_STAMP_TOP_MARGIN_M And lower < LEADER_STAMP_WIDTH_M Then
                    occupied = True
                End If
            ElseIf coordinate <= LEADER_STAMP_WIDTH_M And upper >= sheetHeight - LEADER_STAMP_TOP_MARGIN_M Then
                occupied = True
            End If
            For Each placedBase In bases
                If ProfileBasesConflict(placedBase, view.Name, axis, coordinate, lower, upper, baseKind) Then
                    occupied = True
                End If
            Next placedBase
        End If
        If occupied Then
            GoTo NextProfileCandidate
        End If
        If panel And profilePagedMode Then
            If Not ProfileLabelBaseFits(textBounds, axis, textBaseline, coordinate, region, loops) Then
                GoTo NextProfileCandidate
            End If
        End If
        boundary = False
        If CBool(candidate(2)) Then
            point = Array((a(0) + b(0)) / 2, (a(1) + b(1)) / 2)
            point(1 - axis) = coordinate
            For Each polygon In loops
                If ProfilePolygonLocation(point, polygon) = PROFILE_POINT_BOUNDARY Then
                    boundary = True
                End If
            Next polygon
        End If
        clear = ProfileBaseClear(a, b, axis, coordinate, loops)
        externalWitness = False
        If (sharedDepth Or panel) And Not CBool(candidate(2)) And Not clear Then
            point = Array(a(0), a(1))
            position = Array(b(0), b(1))
            point(1 - axis) = coordinate
            position(1 - axis) = coordinate
            clear = Not ProfileLineCutsMaterial(point, position, loops)
            externalWitness = clear
        End If
        If Not boundary And clear And Not occupied Then
            cost = Abs(a(1 - axis) - coordinate) + Abs(b(1 - axis) - coordinate)
            If CLng(candidate(1)) <> side Then
                cost = cost + PROFILE_WRONG_SIDE_COST_M
            End If
            If CBool(candidate(2)) Then
                If Not profilePagedMode Then
                    cost = cost + PROFILE_LOCAL_BASE_COST_M
                End If
            End If
            If externalWitness And panel And Not sharedDepth Then
                If request("role") <> PROFILE_ROLE_OUTER Then
                    cost = cost + PANEL_EXTERNAL_WITNESS_COST_M
                End If
            End If
            If cost < best Then
                best = cost
                bestCoordinate = coordinate
                bestLocal = CBool(candidate(2))
                bestSide = CLng(candidate(1))
                bestExternalWitness = externalWitness
            End If
        End If
NextProfileCandidate:
    Next candidate
    If best = SEARCH_INITIAL_LIMIT Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "No material-free profile dimension base: " & display.GetNameForSelection
    End If
    If panel Then
        bases.Add Array(view.Name, axis, bestCoordinate, lower, upper, baseKind)
    End If
    request("localBase") = bestLocal
    request("externalWitness") = bestExternalWitness
    request("sheetFirst") = a
    request("sheetLast") = b
    If bestExternalWitness Then
        Set audit = fs.OpenTextFile(fs.BuildPath(runDir, PROFILE_WITNESSES_FILENAME), FILE_FOR_APPENDING, True, FILE_UNICODE)
        audit.WriteLine view.Name & vbTab & display.GetNameForSelection & vbTab & axis & vbTab & a(0) & vbTab & a(1) & vbTab & b(0) & vbTab & b(1)
        audit.Close
    End If
    featureDimensionPlacement(display.GetNameForSelection) = bestSide
    If bestLocal Then
        point = Array((a(0) + b(0)) / 2, (a(1) + b(1)) / 2)
        point(1 - axis) = bestCoordinate
        If Not ProfileLineCutsMaterial(point, Array(point(0), outline(1) - GOST_OBJECT_OFFSET_M), loops) Then
            featureDimensionPlacement(display.GetNameForSelection) = NATIVE_LABEL_SIDE_BELOW
        ElseIf Not ProfileLineCutsMaterial(point, Array(outline(2) + GOST_OBJECT_OFFSET_M, point(1)), loops) Then
            featureDimensionPlacement(display.GetNameForSelection) = NATIVE_LABEL_SIDE_RIGHT
        ElseIf Not ProfileLineCutsMaterial(point, Array(outline(0) - GOST_OBJECT_OFFSET_M, point(1)), loops) Then
            featureDimensionPlacement(display.GetNameForSelection) = NATIVE_LABEL_SIDE_LEFT
        Else
            featureDimensionPlacement(display.GetNameForSelection) = NATIVE_LABEL_SIDE_ABOVE
        End If
        display.ArrowSide = SW_DIMENSION_ARROWS_INSIDE
        If CDbl(request("value")) * view.ScaleDecimal < 2 * GOST_ARROW_HEAD_LENGTH_M Then
            display.SetArrowHeadStyle2 False, SW_ARROW_DOT, SW_ARROW_DOT
            ApplyGostDimensionStyle display
        End If
    End If
    display.OffsetText = False
    display.CenterText = True
    Set ann = display.GetAnnotation
    For index = 1 To LEADER_CLAMP_ATTEMPTS
        d.GraphicsRedraw2
        Set data = display.GetDisplayData
        arrow = data.GetArrowHeadAtIndex2(0)
        If Abs(arrow(1 - axis) - bestCoordinate) < PROFILE_MATERIAL_TOLERANCE_M Then
            Exit For
        End If
        position = ann.GetPosition
        position(1 - axis) = position(1 - axis) + bestCoordinate - arrow(1 - axis)
        ann.SetPosition2 position(0), position(1), position(2)
    Next index
    Log "PROFILE BASE " & display.GetNameForSelection & "; local=" & bestLocal & "; side=" & NativePlacementSide(display) & "; coordinate=" & bestCoordinate
End Sub
