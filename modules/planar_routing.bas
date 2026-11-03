Attribute VB_Name = "planar_routing"
Option Explicit

Private Function PlanarArrowBisector(ByVal center As Variant, ByVal first As Variant, ByVal last As Variant) As Double
    Dim ax As Double, ay As Double, bx As Double, by As Double, aLength As Double, bLength As Double
    ax = first(0) - center(0)
    ay = first(1) - center(1)
    bx = last(0) - center(0)
    by = last(1) - center(1)
    aLength = Sqr(ax * ax + ay * ay)
    bLength = Sqr(bx * bx + by * by)
    PlanarArrowBisector = PlanarAngle(ax / aLength + bx / bLength, ay / aLength + by / bLength)
End Function

' Choose a local principal axis and endpoint whose angular sector is outside visible material.
Private Function PlanarDatumPlan(ByVal p As Variant, ByVal q As Variant, ByVal polygon As Variant) As Variant
    Dim endpoint As Long, axis As Long, row As Long, index As Long, steps As Long
    Dim anchor As Variant, direction As Variant, first As Variant, last As Variant
    Dim dx As Double, dy As Double, length As Double, ux As Double, uy As Double, edgeAngle As Double, axisAngle As Double, sweep As Double
    Dim radius As Double, angle As Double, cost As Double, penalty As Double, best As Double
    dx = q(0) - p(0)
    dy = q(1) - p(1)
    length = Sqr(dx * dx + dy * dy)
    best = SEARCH_INITIAL_LIMIT
    For endpoint = EDGE_ENDPOINT_FIRST To EDGE_ENDPOINT_LAST
        anchor = p
        ux = -dx / length
        uy = -dy / length
        If endpoint = EDGE_ENDPOINT_LAST Then
            anchor = q
            ux = dx / length
            uy = dy / length
        End If
        edgeAngle = PlanarAngle(ux, uy)
        For axis = SCREEN_X_AXIS_INDEX To SCREEN_Y_AXIS_INDEX
            direction = Array(Sgn(ux), 0)
            If axis = SCREEN_Y_AXIS_INDEX Then
                direction = Array(0, Sgn(uy))
            End If
            axisAngle = PlanarAngle(direction(0), direction(1))
            sweep = edgeAngle - axisAngle
            Do While sweep > PI
                sweep = sweep - 2 * PI
            Loop
            Do While sweep < -PI
                sweep = sweep + 2 * PI
            Loop
            penalty = 0
            last = Array(anchor(0) + direction(0) * PLANAR_ANGLE_DATUM_LENGTH_M, anchor(1) + direction(1) * PLANAR_ANGLE_DATUM_LENGTH_M)
            If ProfileLineCutsMaterial(anchor, last, Array(polygon)) Then
                penalty = PLANAR_DATUM_MATERIAL_PENALTY_M
            End If
            For row = 0 To PLANAR_DATUM_SEARCH_ROWS
                radius = GOST_OBJECT_OFFSET_M + row * COORDINATE_LANE_PITCH_M
                steps = Int(Abs(sweep) / PLANAR_ARC_STEP_RAD) + 1
                first = Array(anchor(0) + radius * direction(0), anchor(1) + radius * direction(1))
                cost = radius + penalty
                For index = 1 To steps
                    angle = axisAngle + sweep * index / steps
                    last = Array(anchor(0) + radius * Cos(angle), anchor(1) + radius * Sin(angle))
                    If ProfileLineCutsMaterial(first, last, Array(polygon)) Then
                        cost = cost + PLANAR_DATUM_MATERIAL_PENALTY_M
                    End If
                    first = last
                Next index
                If cost < best Then
                    best = cost
                    PlanarDatumPlan = Array(endpoint, axis, direction, Abs(sweep))
                End If
            Next row
        Next axis
    Next endpoint
End Function

Private Function PlanarOwnsDimension(ByVal display As SldWorks.DisplayDimension) As Boolean
    If Not planarDimensionRequests Is Nothing Then
        PlanarOwnsDimension = planarDimensionRequests.Exists(display.GetNameForSelection)
    End If
End Function

Private Sub StabilizePlanarDimensionText(ByVal d As SldWorks.ModelDoc2, ByVal display As SldWorks.DisplayDimension, ByVal view As SldWorks.View)
    Dim data As SldWorks.DisplayData, bounds As Variant, textLength As Double, span As Double
    Set data = display.GetDisplayData
    bounds = NativeTextBounds(data)
    textLength = Sqr((bounds(2) - bounds(0)) ^ 2 + (bounds(3) - bounds(1)) ^ 2)
    span = display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue * view.ScaleDecimal
    display.OffsetText = False
    display.CenterText = True
    d.GraphicsRedraw2
    bounds = NativeTextBounds(display.GetDisplayData)
    If span < textLength + 2 * GOST_ARROW_HEAD_LENGTH_M Or NativeBoxesOverlap(bounds, view.GetOutline, LEADER_GEOMETRY_CLEARANCE_M) Then
        display.CenterText = False
        display.OffsetText = True
    End If
End Sub

' Touching at an endpoint is allowed; a proper interior crossing is counted.
Private Function PlanarSegmentsCross(ByVal a As Variant, ByVal b As Variant, ByVal c As Variant, ByVal last As Variant) As Boolean
    Dim dx As Double, dy As Double, ex As Double, ey As Double, determinant As Double, t As Double, u As Double
    dx = b(0) - a(0)
    dy = b(1) - a(1)
    ex = last(0) - c(0)
    ey = last(1) - c(1)
    determinant = dx * ey - dy * ex
    If Abs(determinant) <= PROFILE_MATERIAL_TOLERANCE_M ^ 2 Then
        Exit Function
    End If
    t = ((c(0) - a(0)) * ey - (c(1) - a(1)) * ex) / determinant
    u = ((c(0) - a(0)) * dy - (c(1) - a(1)) * dx) / determinant
    PlanarSegmentsCross = t > PLANAR_ENDPOINT_PARAMETER_MARGIN And t < 1 - PLANAR_ENDPOINT_PARAMETER_MARGIN And u > PLANAR_ENDPOINT_PARAMETER_MARGIN And u < 1 - PLANAR_ENDPOINT_PARAMETER_MARGIN
End Function

Private Function PlanarCrossingCount(ByVal a As Variant, ByVal b As Variant, ByVal lines As Collection) As Long
    Dim item As Variant
    For Each item In lines
        If PlanarSegmentsCross(a, b, Array(item(0), item(1)), Array(item(2), item(3))) Then
            PlanarCrossingCount = PlanarCrossingCount + 1
        End If
    Next item
End Function

Private Sub PlacePlanarDimensionBase(ByVal d As SldWorks.ModelDoc2, ByVal display As SldWorks.DisplayDimension, ByVal view As SldWorks.View, ByVal region As Variant, ByVal lines As Collection, ByVal rows As Collection)
    Dim request As Object, a As Variant, b As Variant, first As Variant, last As Variant, point As Variant, polygon() As Variant
    Dim dx As Double, dy As Double, length As Double, nx As Double, ny As Double, distance As Double, winding As Double
    Dim row As Long, side As Long, index As Long, cost As Double, best As Double, chosen As Variant, ann As SldWorks.Annotation
    Dim bounds As Variant, arrows As Variant, other As Variant, data As SldWorks.DisplayData, position As Variant, target As Variant
    Dim emptyObstacles As New Collection
    Set request = planarDimensionRequests(display.GetNameForSelection)
    If request("kind") = "angle" Then
        a = PatternSheetPoint(view, request("first"))
        b = PatternSheetPoint(view, request("last"))
        dx = b(0) - a(0)
        dy = b(1) - a(1)
        length = Sqr(dx * dx + dy * dy)
        winding = request("winding")
        display.SetBrokenLeader2 False, SW_SOLID_LEADER_ALIGNED_TEXT
        display.OffsetText = False
        display.CenterText = True
        Set ann = display.GetAnnotation
        ann.SetPosition2 (a(0) + b(0)) / 2 + dy / length * winding * GOST_OBJECT_OFFSET_M, (a(1) + b(1)) / 2 - dx / length * winding * GOST_OBJECT_OFFSET_M, 0
        d.GraphicsRedraw2
        NativeCollectLines display, lines, True
        Exit Sub
    End If
    If request("kind") <> "edge" Then
        PlaceNativeDimensionBase display, view, region, emptyObstacles, rows
        ClampNativeDimension display, region
        NativeCollectLines display, lines, True
        Exit Sub
    End If
    a = PatternSheetPoint(view, request("first"))
    b = PatternSheetPoint(view, request("last"))
    dx = b(0) - a(0)
    dy = b(1) - a(1)
    length = Sqr(dx * dx + dy * dy)
    winding = request("winding")
    nx = dy / length * winding
    ny = -dx / length * winding
    ReDim polygon(planarContourPoints.Count - 1)
    For index = 1 To planarContourPoints.Count
        polygon(index - 1) = PatternSheetPoint(view, planarContourPoints(index))
    Next index
    best = SEARCH_INITIAL_LIMIT
    For row = 0 To PLANAR_BASE_ROW_LIMIT
        For side = 0 To 1
            distance = (GOST_OBJECT_OFFSET_M + row * COORDINATE_LANE_PITCH_M) * (1 - 2 * side)
            first = Array(a(0) + nx * distance, a(1) + ny * distance)
            last = Array(b(0) + nx * distance, b(1) + ny * distance)
            bounds = Array(MinD(first(0), last(0)) - GOST_ARROW_LENGTH_M, MinD(first(1), last(1)) - GOST_ARROW_LENGTH_M, -MinD(-first(0), -last(0)) + GOST_ARROW_LENGTH_M, -MinD(-first(1), -last(1)) + GOST_ARROW_LENGTH_M)
            If NativeBoundsInside(bounds, region) And Not ProfileLineCutsMaterial(first, last, Array(polygon)) Then
                cost = Abs(distance) + side * PLANAR_WRONG_SIDE_COST_M
                cost = cost + PLANAR_CROSSING_COST_M * (PlanarCrossingCount(first, last, lines) + PlanarCrossingCount(a, first, lines) + PlanarCrossingCount(b, last, lines))
                If cost < best Then
                    best = cost
                    chosen = Array((first(0) + last(0)) / 2, (first(1) + last(1)) / 2)
                End If
            End If
        Next side
    Next row
    If best = SEARCH_INITIAL_LIMIT Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "A planar length has no clear dimension baseline on this sheet."
    End If
    display.OffsetText = False
    display.CenterText = True
    Set ann = display.GetAnnotation
    For index = 1 To LEADER_CLAMP_ATTEMPTS
        d.GraphicsRedraw2
        Set data = display.GetDisplayData
        arrows = data.GetArrowHeadAtIndex2(0)
        other = data.GetArrowHeadAtIndex2(1)
        target = Array(chosen(0) - (arrows(0) + other(0)) / 2, chosen(1) - (arrows(1) + other(1)) / 2)
        If Sqr(target(0) ^ 2 + target(1) ^ 2) < PROFILE_MATERIAL_TOLERANCE_M Then
            Exit For
        End If
        position = ann.GetPosition
        ann.SetPosition2 position(0) + target(0), position(1) + target(1), position(2)
    Next index
    NativeCollectLines display, lines, True
    Log "PLANAR BASE " & display.GetNameForSelection & "; cost=" & best
End Sub
