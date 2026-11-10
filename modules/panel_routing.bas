Attribute VB_Name = "panel_routing"
Option Explicit

Private Function PanelLabelPriority(ByVal span As Double, ByVal aligned As Boolean, ByVal localBase As Boolean) As Double
    If aligned Then
        PanelLabelPriority = PANEL_ALIGNED_LABEL_PRIORITY_M
        Exit Function
    End If
    If localBase Then
        PanelLabelPriority = PANEL_LOCAL_LABEL_PRIORITY_M
    End If
    If span < 2 * GOST_ARROW_LENGTH_M Then
        PanelLabelPriority = PANEL_SHORT_LABEL_PRIORITY_M + (1 - span / (2 * GOST_ARROW_LENGTH_M)) * PANEL_SHORT_LABEL_SPREAD_M
    End If
End Function

' DisplayData does not promise that the movable leader is the last two lines.
' Fixed lines lie on the measured baseline or on an endpoint's normal witness.
Private Function NativeFixedLinearSegment(ByVal segment As Variant, ByVal first As Variant, ByVal last As Variant) As Boolean
    Dim dx As Double, dy As Double, length As Double, stationA As Double, stationB As Double, normalA As Double, normalB As Double
    dx = last(0) - first(0)
    dy = last(1) - first(1)
    length = Sqr(dx * dx + dy * dy)
    If length <= DIRECTION_ZERO_TOLERANCE Then
        Exit Function
    End If
    dx = dx / length
    dy = dy / length
    stationA = (segment(0) - first(0)) * dx + (segment(1) - first(1)) * dy
    stationB = (segment(2) - first(0)) * dx + (segment(3) - first(1)) * dy
    normalA = (segment(0) - first(0)) * dy - (segment(1) - first(1)) * dx
    normalB = (segment(2) - first(0)) * dy - (segment(3) - first(1)) * dx
    If Abs(normalA) <= PROFILE_MATERIAL_TOLERANCE_M And Abs(normalB) <= PROFILE_MATERIAL_TOLERANCE_M Then
        NativeFixedLinearSegment = True
    ElseIf Abs(stationA - stationB) <= PROFILE_MATERIAL_TOLERANCE_M Then
        NativeFixedLinearSegment = Abs(stationA) <= PROFILE_MATERIAL_TOLERANCE_M Or Abs(stationA - length) <= PROFILE_MATERIAL_TOLERANCE_M
    End If
End Function

Private Function PanelMaterialBounds(ByVal polygon As Variant) As Variant
    Dim result As Variant, point As Variant
    result = Array(SEARCH_INITIAL_LIMIT, SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT)
    For Each point In polygon
        result(0) = MinD(result(0), point(0))
        result(1) = MinD(result(1), point(1))
        result(2) = -MinD(-result(2), -point(0))
        result(3) = -MinD(-result(3), -point(1))
    Next point
    PanelMaterialBounds = result
End Function

Private Function PanelOwnsDimension(ByVal display As SldWorks.DisplayDimension) As Boolean
    If ProfileOwnsDimension(display) Then
        PanelOwnsDimension = profileDimensionRequests(display.GetNameForSelection).Exists("panel")
    End If
End Function

' A view's rectangular outline includes its empty cutouts. Label ink is checked
' against the filled contour instead, with the normal graphical clearance.
Private Function PanelLabelTouchesMaterial(ByVal display As SldWorks.DisplayDimension, ByVal bounds As Variant) As Boolean
    Dim request As Object, loops As Variant, box As Variant
    Set request = profileDimensionRequests(display.GetNameForSelection)
    loops = profileSheetContours(request("view"))
    box = Array(bounds(0) - LEADER_GEOMETRY_CLEARANCE_M, bounds(1) - LEADER_GEOMETRY_CLEARANCE_M, bounds(2) + LEADER_GEOMETRY_CLEARANCE_M, bounds(3) + LEADER_GEOMETRY_CLEARANCE_M)
    If Not request.Exists("materialBounds") Then
        request.Add "materialBounds", PanelMaterialBounds(loops(0))
    End If
    If Not NativeBoxesOverlap(box, request("materialBounds"), 0) Then
        Exit Function
    End If
    PanelLabelTouchesMaterial = PanelBoxTouchesMaterial(box, loops)
End Function

Private Function PanelBoxTouchesMaterial(ByVal box As Variant, ByVal loops As Variant) As Boolean
    Dim corners As Variant, corner As Variant, polygon As Variant, index As Long
    If ProfilePointInMaterial(Array((box(0) + box(2)) / 2, (box(1) + box(3)) / 2), loops) Then
        PanelBoxTouchesMaterial = True
        Exit Function
    End If
    corners = Array(Array(box(0), box(1)), Array(box(2), box(1)), Array(box(2), box(3)), Array(box(0), box(3)))
    For index = LBound(corners) To UBound(corners)
        If ProfilePointInMaterial(corners(index), loops) Or ProfileLineCutsMaterial(corners(index), corners((index + 1) Mod (UBound(corners) + 1)), loops) Then
            PanelBoxTouchesMaterial = True
            Exit Function
        End If
    Next index
    For Each polygon In loops
        For Each corner In polygon
            If corner(0) >= box(0) And corner(0) <= box(2) And corner(1) >= box(1) And corner(1) <= box(3) Then
                PanelBoxTouchesMaterial = True
                Exit Function
            End If
        Next corner
    Next polygon
End Function

Private Function PanelReferenceWitness(ByVal first As Variant, ByVal last As Variant, ByVal a As Variant, ByVal b As Variant) As Boolean
    Dim dx As Double, dy As Double, length As Double, firstStation As Double, lastStation As Double
    dx = b(0) - a(0)
    dy = b(1) - a(1)
    length = Sqr(dx * dx + dy * dy)
    firstStation = ((first(0) - a(0)) * dx + (first(1) - a(1)) * dy) / length
    lastStation = ((last(0) - a(0)) * dx + (last(1) - a(1)) * dy) / length
    If Abs(firstStation - lastStation) <= PROFILE_MATERIAL_TOLERANCE_M Then
        PanelReferenceWitness = Abs(firstStation) <= PROFILE_MATERIAL_TOLERANCE_M Or Abs(firstStation - length) <= PROFILE_MATERIAL_TOLERANCE_M
    End If
End Function

Private Function PanelAlignedOffset(ByVal origin As Variant, ByVal nx As Double, ByVal ny As Double, ByVal polygon As Variant) As Double
    Dim point As Variant, station As Double
    For Each point In polygon
        station = (point(0) - origin(0)) * nx + (point(1) - origin(1)) * ny
        If station > PanelAlignedOffset Then
            PanelAlignedOffset = station
        End If
    Next point
    PanelAlignedOffset = PanelAlignedOffset + GOST_OBJECT_OFFSET_M
End Function

Private Sub PlacePanelAlignedBase(ByVal d As SldWorks.ModelDoc2, ByVal display As SldWorks.DisplayDimension, ByVal view As SldWorks.View, ByVal request As Object, ByVal region As Variant)
    Dim a As Variant, b As Variant, loops As Variant, outline As Variant, corners As Variant, corner As Variant
    Dim dx As Double, dy As Double, length As Double, nx As Double, ny As Double, station As Double, offset As Double
    Dim side As Long, row As Long, first As Variant, last As Variant, position As Variant, arrow As Variant, data As SldWorks.DisplayData
    Dim annotation As SldWorks.Annotation, index As Long, found As Boolean
    Dim best As Double, bestFirst As Variant, bestNx As Double, bestNy As Double
    a = PatternSheetPoint(view, request("first"))
    b = PatternSheetPoint(view, request("last"))
    dx = b(0) - a(0)
    dy = b(1) - a(1)
    length = Sqr(dx * dx + dy * dy)
    dx = dx / length
    dy = dy / length
    loops = profileSheetContours(view.Name)
    outline = PanelMaterialBounds(loops(0))
    best = SEARCH_INITIAL_LIMIT
    For row = 0 To LEADER_BASE_ROW_LIMIT
        For side = -1 To 1 Step 2
            nx = -dy * side
            ny = dx * side
            offset = PanelAlignedOffset(a, nx, ny, loops(0)) + row * COORDINATE_LANE_PITCH_M
            If profilePagedMode And request("kind") = "notch-location" Then
                offset = offset + COORDINATE_LANE_PITCH_M
            End If
            first = Array(a(0) + nx * offset, a(1) + ny * offset)
            last = Array(b(0) + nx * offset, b(1) + ny * offset)
            If NativeBoundsInside(Array(MinD(first(0), last(0)), MinD(first(1), last(1)), -MinD(-first(0), -last(0)), -MinD(-first(1), -last(1))), region) Then
                If Not ProfileLineCutsMaterial(first, last, loops) And offset < best Then
                    best = offset
                    bestFirst = first
                    bestNx = nx
                    bestNy = ny
                    found = True
                End If
            End If
        Next side
        If found Then
            Exit For
        End If
    Next row
    If Not found Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "No external aligned panel dimension base."
    End If
    first = bestFirst
    nx = bestNx
    ny = bestNy
    request("localBase") = False
    request("externalWitness") = True
    request("sheetFirst") = a
    request("sheetLast") = b
    If length >= 2 * GOST_ARROW_HEAD_LENGTH_M Then
        display.ArrowSide = SW_DIMENSION_ARROWS_INSIDE
    Else
        display.ArrowSide = SW_DIMENSION_ARROWS_OUTSIDE
    End If
    If Abs(nx) >= Abs(ny) Then
        side = NATIVE_LABEL_SIDE_LEFT
        If nx > 0 Then
            side = NATIVE_LABEL_SIDE_RIGHT
        End If
    Else
        side = NATIVE_LABEL_SIDE_BELOW
        If ny > 0 Then
            side = NATIVE_LABEL_SIDE_ABOVE
        End If
    End If
    featureDimensionPlacement(display.GetNameForSelection) = side
    display.OffsetText = False
    display.CenterText = True
    Set annotation = display.GetAnnotation
    For index = 1 To LEADER_CLAMP_ATTEMPTS
        d.GraphicsRedraw2
        Set data = display.GetDisplayData
        arrow = data.GetArrowHeadAtIndex2(0)
        station = (first(0) - arrow(0)) * nx + (first(1) - arrow(1)) * ny
        If Abs(station) < PROFILE_MATERIAL_TOLERANCE_M Then
            Exit For
        End If
        position = annotation.GetPosition
        annotation.SetPosition2 position(0) + nx * station, position(1) + ny * station, position(2)
    Next index
End Sub
