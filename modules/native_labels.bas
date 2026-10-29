Attribute VB_Name = "native_labels"
Option Explicit
Private Function NativeTextBounds(ByVal data As SldWorks.DisplayData) As Variant
    Dim result As Variant, p As Variant, i As Long, j As Long
    Dim w As Double, h As Double, a As Double, x As Double, y As Double, dx As Double, dy As Double
    Dim text As String, line As Variant, lineIndex As Long, firstAlong As Double, lastAlong As Double
    Dim firstNormal As Double, lastNormal As Double, tail As Double, measured As Double
    result = Array(SEARCH_INITIAL_LIMIT, SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT)
    For i = 0 To data.GetTextCount - 1
        p = data.GetTextPositionAtIndex(i)
        h = data.GetTextHeightAtIndex(i)
        text = Replace(data.GetTextAtIndex(i), SW_DIMENSION_TOKEN_DIAMETER, "D")
        text = Replace(text, SW_DIMENSION_TOKEN_DEGREE, "d")
        text = Replace(text, SW_DIMENSION_TOKEN_PLUS_MINUS, "+")
        w = Len(text) * h * LEADER_GLYPH_WIDTH_RATIO
        a = data.GetTextAngleAtIndex(i)
        measured = SEARCH_INITIAL_LIMIT
        For lineIndex = 0 To data.GetLineCount - 1
            line = data.GetLineAtIndex2(lineIndex)
            dx = line(DISPLAY_LINE_START_X) - p(0)
            dy = line(DISPLAY_LINE_START_Y) - p(1)
            firstAlong = dx * Cos(a) + dy * Sin(a)
            firstNormal = -dx * Sin(a) + dy * Cos(a)
            dx = line(DISPLAY_LINE_END_X) - p(0)
            dy = line(DISPLAY_LINE_END_Y) - p(1)
            lastAlong = dx * Cos(a) + dy * Sin(a)
            lastNormal = -dx * Sin(a) + dy * Cos(a)
            If Abs(firstNormal) < DIRECTION_ZERO_TOLERANCE And Abs(lastNormal) < DIRECTION_ZERO_TOLERANCE Then
                If MinD(firstAlong, lastAlong) <= DIRECTION_ZERO_TOLERANCE Then
                    tail = -MinD(-firstAlong, -lastAlong)
                    If tail >= Len(Trim$(text)) * h * TEXT_ESTIMATED_GLYPH_WIDTH_RATIO / 2 And tail <= w + LEADER_SHOULDER_LENGTH_M Then
                        measured = MinD(measured, tail)
                    End If
                End If
            End If
        Next lineIndex
        If measured < SEARCH_INITIAL_LIMIT Then
            w = measured
        End If
        For j = 0 To 3
            dx = 0
            dy = 0
            If j Mod 2 = 1 Then
                dx = w
            End If
            If j >= 2 Then
                dy = h
            End If
            x = p(0) + dx * Cos(a) - dy * Sin(a)
            y = p(1) + dx * Sin(a) + dy * Cos(a)
            NativeExpandBounds result, x, y
        Next j
    Next i
    NativeTextBounds = result
End Function

Private Sub NativeExpandBounds(ByRef bounds As Variant, ByVal x As Double, ByVal y As Double)
    bounds(0) = MinD(bounds(0), x)
    bounds(1) = MinD(bounds(1), y)
    bounds(2) = -MinD(-bounds(2), -x)
    bounds(3) = -MinD(-bounds(3), -y)
End Sub

Private Function NativeDimensionBounds(ByVal display As SldWorks.DisplayDimension) As Variant
    Dim data As SldWorks.DisplayData, bounds As Variant, line As Variant, arrow As Variant, i As Long
    Dim dx As Double, dy As Double, length As Double, width As Double
    Set data = display.GetDisplayData
    bounds = NativeTextBounds(data)
    For i = 0 To data.GetLineCount - 1
        line = data.GetLineAtIndex2(i)
        NativeExpandBounds bounds, line(DISPLAY_LINE_START_X), line(DISPLAY_LINE_START_Y)
        NativeExpandBounds bounds, line(DISPLAY_LINE_END_X), line(DISPLAY_LINE_END_Y)
    Next i
    For i = 0 To data.GetArrowHeadCount - 1
        arrow = data.GetArrowHeadAtIndex2(i)
        length = arrow(DISPLAY_ARROW_WIDTH)
        width = arrow(DISPLAY_ARROW_HEIGHT) / 2
        dx = arrow(DISPLAY_ARROW_DIRECTION_X)
        dy = arrow(DISPLAY_ARROW_DIRECTION_Y)
        NativeExpandBounds bounds, arrow(0), arrow(1)
        NativeExpandBounds bounds, arrow(0) + length * dx - width * dy, arrow(1) + length * dy + width * dx
        NativeExpandBounds bounds, arrow(0) + length * dx + width * dy, arrow(1) + length * dy - width * dx
    Next i
    NativeDimensionBounds = bounds
End Function

Private Function NativeBoundsInside(ByVal bounds As Variant, ByVal region As Variant) As Boolean
    NativeBoundsInside = bounds(0) >= region(0) - DIMENSION_GEOMETRY_GAP_M And bounds(1) >= region(1) - DIMENSION_GEOMETRY_GAP_M And bounds(2) <= region(2) + DIMENSION_GEOMETRY_GAP_M And bounds(3) <= region(3) + DIMENSION_GEOMETRY_GAP_M
End Function

Private Function NativeBoxesOverlap(ByVal a As Variant, ByVal b As Variant, ByVal clearance As Double) As Boolean
    NativeBoxesOverlap = a(0) < b(2) + clearance And a(2) > b(0) - clearance And a(1) < b(3) + clearance And a(3) > b(1) - clearance
End Function

Private Sub PlaceNativeDimensionBase(ByVal display As SldWorks.DisplayDimension, ByVal v As SldWorks.View, ByVal region As Variant, ByVal obstacles As Collection, ByVal rows As Collection)
    Dim data As SldWorks.DisplayData, a As Variant, b As Variant, outline As Variant, box As Variant, item As Variant, p As Variant
    Dim horizontal As Boolean, row As Long, side As Long, coordinate As Double, lower As Double, upper As Double
    Dim best As Double, cost As Double, bestCoordinate As Double, valid As Boolean, found As Boolean, ann As SldWorks.Annotation
    If display.Type2 <> SW_DIMENSION_LINEAR Then
        Exit Sub
    End If
    display.OffsetText = False
    display.CenterText = True
    Set data = display.GetDisplayData
    If data.GetArrowHeadCount <> 2 Then
        Exit Sub
    End If
    a = data.GetArrowHeadAtIndex2(0)
    b = data.GetArrowHeadAtIndex2(1)
    horizontal = Abs(a(1) - b(1)) < DIRECTION_ZERO_TOLERANCE
    If Not horizontal And Abs(a(0) - b(0)) >= DIRECTION_ZERO_TOLERANCE Then
        Exit Sub
    End If
    If horizontal Then
        lower = MinD(a(0), b(0))
        upper = -MinD(-a(0), -b(0))
    Else
        lower = MinD(a(1), b(1))
        upper = -MinD(-a(1), -b(1))
    End If
    outline = v.GetOutline
    best = SEARCH_INITIAL_LIMIT
    For row = 0 To LEADER_BASE_ROW_LIMIT
        For side = 0 To 1
            If horizontal Then
                coordinate = outline(1) - GOST_OBJECT_OFFSET_M - row * COORDINATE_LANE_PITCH_M
                If side = 1 Then
                    coordinate = outline(3) + GOST_OBJECT_OFFSET_M + row * COORDINATE_LANE_PITCH_M
                End If
                box = Array(lower - GOST_ARROW_LENGTH_M, coordinate - GOST_ARROW_HEAD_WIDTH_M, upper + GOST_ARROW_LENGTH_M, coordinate + GOST_ARROW_HEAD_WIDTH_M)
            Else
                coordinate = outline(0) - GOST_OBJECT_OFFSET_M - row * COORDINATE_LANE_PITCH_M
                If side = 1 Then
                    coordinate = outline(2) + GOST_OBJECT_OFFSET_M + row * COORDINATE_LANE_PITCH_M
                End If
                box = Array(coordinate - GOST_ARROW_HEAD_WIDTH_M, lower - GOST_ARROW_LENGTH_M, coordinate + GOST_ARROW_HEAD_WIDTH_M, upper + GOST_ARROW_LENGTH_M)
            End If
            valid = NativeBoundsInside(box, region)
            If valid Then
                For Each item In obstacles
                    If NativeBoxesOverlap(box, item, LEADER_LABEL_CLEARANCE_M) Then
                        valid = False
                        Exit For
                    End If
                Next item
            End If
            If valid Then
                cost = row * COORDINATE_LANE_PITCH_M + side * DIMENSION_GEOMETRY_GAP_M
                For Each item In rows
                    If NativeBoxesOverlap(box, item, LEADER_LABEL_CLEARANCE_M) Then
                        cost = cost + LEADER_OCCUPIED_ROW_COST
                    End If
                Next item
                If cost < best Then
                    best = cost
                    bestCoordinate = coordinate
                    found = True
                End If
            End If
        Next side
    Next row
    If Not found Then
        Exit Sub
    End If
    Set ann = display.GetAnnotation
    p = ann.GetPosition
    If horizontal Then
        ann.SetPosition2 (lower + upper) / 2, bestCoordinate, p(2)
        rows.Add Array(lower - GOST_ARROW_LENGTH_M, bestCoordinate - GOST_ARROW_HEAD_WIDTH_M, upper + GOST_ARROW_LENGTH_M, bestCoordinate + GOST_ARROW_HEAD_WIDTH_M)
    Else
        ann.SetPosition2 bestCoordinate, (lower + upper) / 2, p(2)
        rows.Add Array(bestCoordinate - GOST_ARROW_HEAD_WIDTH_M, lower - GOST_ARROW_LENGTH_M, bestCoordinate + GOST_ARROW_HEAD_WIDTH_M, upper + GOST_ARROW_LENGTH_M)
    End If
End Sub

Private Sub ClampNativeDimension(ByVal display As SldWorks.DisplayDimension, ByVal region As Variant)
    Dim bounds As Variant, p As Variant, ann As SldWorks.Annotation, dx As Double, dy As Double, attempt As Long
    Set ann = display.GetAnnotation
    display.SetBrokenLeader2 False, SW_BROKEN_LEADER_HORIZONTAL_TEXT
    If display.Type2 = SW_DIMENSION_RADIAL Or display.Type2 = SW_DIMENSION_DIAMETER Then
        display.SolidLeader = False
        display.SetBentLeaderLength False, LEADER_SHOULDER_LENGTH_M
    End If
    display.OffsetText = False
    For attempt = 1 To LEADER_CLAMP_ATTEMPTS
        bounds = NativeDimensionBounds(display)
        If NativeBoundsInside(bounds, region) Then
            Exit Sub
        End If
        dx = 0
        dy = 0
        If bounds(0) < region(0) Then
            dx = region(0) - bounds(0)
        ElseIf bounds(2) > region(2) Then
            dx = region(2) - bounds(2)
        End If
        If bounds(1) < region(1) Then
            dy = region(1) - bounds(1)
        ElseIf bounds(3) > region(3) Then
            dy = region(3) - bounds(3)
        End If
        p = ann.GetPosition
        ann.SetPosition2 p(0) + dx, p(1) + dy, p(2)
    Next attempt
End Sub

Private Function NativeLeaderCrossesBox(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal bounds As Variant) As Boolean
    Dim low As Double, high As Double, a As Double, b As Double, delta As Double, i As Long
    low = 0
    high = 1
    For i = 0 To 1
        If i = 0 Then
            delta = x2 - x1
            a = bounds(0) - x1
            b = bounds(2) - x1
        Else
            delta = y2 - y1
            a = bounds(1) - y1
            b = bounds(3) - y1
        End If
        If Abs(delta) < DIRECTION_ZERO_TOLERANCE Then
            If a > 0 Or b < 0 Then
                Exit Function
            End If
        Else
            If delta < 0 Then
                delta = -delta
                a = -a
                b = -b
            End If
            low = -MinD(-low, -MinD(a / delta, b / delta))
            high = MinD(high, -MinD(-a / delta, -b / delta))
            If low > high Then
                Exit Function
            End If
        End If
    Next i
    NativeLeaderCrossesBox = True
End Function

Private Function NativeProtectedSpan(ByVal first As Variant, ByVal last As Variant, ByVal margin As Double) As Variant
    Dim dx As Double, dy As Double, length As Double, factor As Double
    dx = last(0) - first(0)
    dy = last(1) - first(1)
    length = Sqr(dx * dx + dy * dy)
    If length > MIN_PROJECTED_EXTENT_M Then
        factor = margin / length
    End If
    NativeProtectedSpan = Array(first(0) - dx * factor, first(1) - dy * factor, last(0) + dx * factor, last(1) + dy * factor)
End Function

Private Sub NativeCollectLines(ByVal display As SldWorks.DisplayDimension, ByVal lines As Collection, Optional ByVal fixedOnly As Boolean = False)
    Dim data As SldWorks.DisplayData, i As Long, line As Variant, first As Variant, last As Variant, count As Long
    Set data = display.GetDisplayData
    count = data.GetLineCount
    If fixedOnly Then
        If display.Type2 = SW_DIMENSION_LINEAR And display.OffsetText Then
            count = count - NATIVE_OFFSET_LEADER_SEGMENTS
        ElseIf display.Type2 = SW_DIMENSION_RADIAL Or display.Type2 = SW_DIMENSION_DIAMETER Then
            count = 0
        End If
    End If
    For i = 0 To count - 1
        line = data.GetLineAtIndex2(i)
        lines.Add Array(line(DISPLAY_LINE_START_X), line(DISPLAY_LINE_START_Y), line(DISPLAY_LINE_END_X), line(DISPLAY_LINE_END_Y))
    Next i
    If display.Type2 = SW_DIMENSION_LINEAR And data.GetArrowHeadCount = 2 Then
        first = data.GetArrowHeadAtIndex2(0)
        last = data.GetArrowHeadAtIndex2(1)
        lines.Add NativeProtectedSpan(first, last, GOST_ARROW_LENGTH_M)
    End If
End Sub

Private Function NativeLinesHitLabels(ByVal display As SldWorks.DisplayDimension, ByVal labels As Collection, Optional ByVal diagnose As Boolean = False) As Boolean
    Dim data As SldWorks.DisplayData, index As Long, line As Variant, item As Variant, interior As Variant
    Set data = display.GetDisplayData
    For Each item In labels
        interior = Array(item(0) + DIMENSION_GEOMETRY_GAP_M, item(1) + DIMENSION_GEOMETRY_GAP_M, item(2) - DIMENSION_GEOMETRY_GAP_M, item(3) - DIMENSION_GEOMETRY_GAP_M)
        For index = 0 To data.GetLineCount - 1
            line = data.GetLineAtIndex2(index)
            If NativeLeaderCrossesBox(line(DISPLAY_LINE_START_X), line(DISPLAY_LINE_START_Y), line(DISPLAY_LINE_END_X), line(DISPLAY_LINE_END_Y), interior) Then
                If diagnose Then
                    Log "NATIVE COLLISION " & display.GetNameForSelection & "; line=" & CStr(line(DISPLAY_LINE_START_X)) & "," & CStr(line(DISPLAY_LINE_START_Y)) & "," & CStr(line(DISPLAY_LINE_END_X)) & "," & CStr(line(DISPLAY_LINE_END_Y)) & "; label=" & Join(item, ",")
                End If
                NativeLinesHitLabels = True
                Exit Function
            End If
        Next index
    Next item
End Function

Private Sub PlaceNativeDimensionLabel(ByVal d As SldWorks.ModelDoc2, ByVal display As SldWorks.DisplayDimension, ByVal region As Variant, ByVal obstacles As Collection, ByVal labels As Collection, ByVal lines As Collection, ByVal ownOutline As Variant)
    Dim data As SldWorks.DisplayData, ann As SldWorks.Annotation, bounds As Variant, p As Variant, box As Variant, item As Variant
    Dim x As Double, y As Double, w As Double, h As Double, cx As Double, cy As Double, bestX As Double, bestY As Double
    Dim cost As Double, best As Double, valid As Boolean, found As Boolean, anchorX As Double, anchorY As Double, pass As Long
    Dim arrow As Variant, arrow2 As Variant, labelSide As Long
    Dim rejected As New Collection, oldPosition As Variant, retries As Long, textFormat As SldWorks.TextFormat
    Dim requiredWidth As Double, requiredHeight As Double, actual As Variant, resizeAttempts As Long
    Set ann = display.GetAnnotation
    display.SetBrokenLeader2 False, SW_BROKEN_LEADER_HORIZONTAL_TEXT
    If display.Type2 = SW_DIMENSION_RADIAL Or display.Type2 = SW_DIMENSION_DIAMETER Then
        display.SolidLeader = False
        display.SetBentLeaderLength False, LEADER_SHOULDER_LENGTH_M
    Else
        display.CenterText = False
        display.OffsetText = True
    End If
    oldPosition = ann.GetPosition
SearchNativeLabel:
    found = False
    Set data = display.GetDisplayData
    bounds = NativeTextBounds(data)
    w = bounds(2) - bounds(0)
    h = bounds(3) - bounds(1)
    w = -MinD(-w, -requiredWidth)
    h = -MinD(-h, -requiredHeight)
    anchorX = (bounds(0) + bounds(2)) / 2
    anchorY = (bounds(1) + bounds(3)) / 2
    If data.GetArrowHeadCount > 0 Then
        arrow = data.GetArrowHeadAtIndex2(0)
        anchorX = arrow(0)
        anchorY = arrow(1)
        If display.Type2 = SW_DIMENSION_LINEAR And data.GetArrowHeadCount = 2 Then
            arrow2 = data.GetArrowHeadAtIndex2(1)
            anchorX = (arrow(0) + arrow2(0)) / 2
            anchorY = (arrow(1) + arrow2(1)) / 2
            If Abs(arrow(1) - arrow2(1)) < DIRECTION_ZERO_TOLERANCE Then
                If anchorY > ownOutline(3) Then
                    labelSide = NATIVE_LABEL_SIDE_ABOVE
                ElseIf anchorY < ownOutline(1) Then
                    labelSide = NATIVE_LABEL_SIDE_BELOW
                End If
            ElseIf Abs(arrow(0) - arrow2(0)) < DIRECTION_ZERO_TOLERANCE Then
                If anchorX < ownOutline(0) Then
                    labelSide = NATIVE_LABEL_SIDE_LEFT
                ElseIf anchorX > ownOutline(2) Then
                    labelSide = NATIVE_LABEL_SIDE_RIGHT
                End If
            End If
        End If
    End If
    best = SEARCH_INITIAL_LIMIT
    For pass = 1 To LEADER_LABEL_SEARCH_PASSES
        For y = region(1) + LEADER_LABEL_CLEARANCE_M To region(3) - h - LEADER_LABEL_CLEARANCE_M Step LEADER_GRID_PITCH_M
            For x = region(0) + LEADER_LABEL_CLEARANCE_M To region(2) - w - LEADER_LABEL_CLEARANCE_M Step LEADER_GRID_PITCH_M
                cx = x + w / 2
                cy = y + h / 2
                If Abs(cx - anchorX) >= LEADER_MIN_DIAGONAL_M And Abs(cy - anchorY) >= LEADER_MIN_DIAGONAL_M Then
                    box = Array(x, y, x + w, y + h)
                    cost = (cx - anchorX) ^ 2 + (cy - anchorY) ^ 2
                    valid = cost < best
                    If pass = 1 Then
                        If labelSide = NATIVE_LABEL_SIDE_ABOVE And box(1) < ownOutline(3) + LEADER_GEOMETRY_CLEARANCE_M Then
                            valid = False
                        ElseIf labelSide = NATIVE_LABEL_SIDE_BELOW And box(3) > ownOutline(1) - LEADER_GEOMETRY_CLEARANCE_M Then
                            valid = False
                        ElseIf labelSide = NATIVE_LABEL_SIDE_LEFT And box(2) > ownOutline(0) - LEADER_GEOMETRY_CLEARANCE_M Then
                            valid = False
                        ElseIf labelSide = NATIVE_LABEL_SIDE_RIGHT And box(0) < ownOutline(2) + LEADER_GEOMETRY_CLEARANCE_M Then
                            valid = False
                        End If
                    End If
                    If display.Type2 = SW_DIMENSION_LINEAR Then
                        If anchorX < ownOutline(0) Or anchorX > ownOutline(2) Or anchorY < ownOutline(1) Or anchorY > ownOutline(3) Then
                            If NativeLeaderCrossesBox(anchorX, anchorY, cx, cy, ownOutline) Then
                                valid = False
                            End If
                        End If
                    End If
                    If Not valid Then
                        GoTo NextLabelCandidate
                    End If
                    For Each item In rejected
                        If Abs(x - item(0)) < DIRECTION_ZERO_TOLERANCE And Abs(y - item(1)) < DIRECTION_ZERO_TOLERANCE Then
                            valid = False
                            Exit For
                        End If
                    Next item
                    If Not valid Then
                        GoTo NextLabelCandidate
                    End If
                    For Each item In obstacles
                        If NativeBoxesOverlap(box, item, LEADER_GEOMETRY_CLEARANCE_M) Then
                            valid = False
                            Exit For
                        End If
                        If Not NativeSameBox(item, ownOutline) Then
                            If NativeLeaderCrossesBox(anchorX, anchorY, box(0), box(1), item) Or NativeLeaderCrossesBox(anchorX, anchorY, box(2), box(1), item) Then
                                valid = False
                                Exit For
                            End If
                        End If
                    Next item
                    If valid Then
                        For Each item In labels
                            If NativeBoxesOverlap(box, item, LEADER_LABEL_CLEARANCE_M) Then
                                valid = False
                                Exit For
                            End If
                        Next item
                    End If
                    If valid Then
                        For Each item In labels
                            If NativeLeaderCrossesBox(anchorX, anchorY, box(0), box(1), item) Or NativeLeaderCrossesBox(anchorX, anchorY, box(2), box(1), item) Then
                                valid = False
                                Exit For
                            End If
                        Next item
                    End If
                    If valid Then
                        For Each item In lines
                            If NativeLeaderCrossesBox(item(0), item(1), item(2), item(3), box) Then
                                valid = False
                                Exit For
                            End If
                        Next item
                    End If
                    If valid Then
                        If cost < best Then
                            best = cost
                            bestX = x
                            bestY = y
                            found = True
                        End If
                    End If
                End If
    NextLabelCandidate:
            Next x
        Next y
        If found Then
            Exit For
        End If
    Next pass
    If Not found Then
        Set textFormat = ann.GetTextFormat(SW_ANNOTATION_TEXT_FORMAT_FIRST)
        If textFormat.CharHeight > TEXT_HEIGHT_SMALL_M + DIMENSION_GEOMETRY_GAP_M Then
            textFormat.CharHeight = TEXT_HEIGHT_SMALL_M
            ann.SetTextFormat SW_ANNOTATION_TEXT_FORMAT_FIRST, False, textFormat
            d.GraphicsRedraw2
            Set rejected = New Collection
            retries = 0
            requiredWidth = 0
            requiredHeight = 0
            GoTo SearchNativeLabel
        End If
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "No free in-frame position for dimension " & display.GetNameForSelection
    End If
    p = ann.GetPosition
    ann.SetPosition2 p(0) + bestX - bounds(0), p(1) + bestY - bounds(1), p(2)
    Set data = display.GetDisplayData
    actual = NativeTextBounds(data)
    If actual(2) - actual(0) > w + DIMENSION_GEOMETRY_GAP_M Or actual(3) - actual(1) > h + DIMENSION_GEOMETRY_GAP_M Then
        ' Native text may rotate or acquire a longer shoulder on its first move.
        resizeAttempts = resizeAttempts + 1
        Log "NATIVE TEXT REFIT " & display.GetNameForSelection & "; attempt=" & resizeAttempts & "; planned=" & CStr(w) & "," & CStr(h) & "; actual=" & CStr(actual(2) - actual(0)) & "," & CStr(actual(3) - actual(1))
        If resizeAttempts > LEADER_CLAMP_ATTEMPTS Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Native text bounds did not stabilize: " & display.GetNameForSelection
        End If
        requiredWidth = -MinD(-requiredWidth, -(actual(2) - actual(0)))
        requiredHeight = -MinD(-requiredHeight, -(actual(3) - actual(1)))
        ann.SetPosition2 oldPosition(0), oldPosition(1), oldPosition(2)
        GoTo SearchNativeLabel
    End If
    p = ann.GetPosition
    ann.SetPosition2 p(0) + bestX - actual(0), p(1) + bestY - actual(1), p(2)
    If NativeLinesHitLabels(display, labels) Then
        If retries = 0 Then
            Call NativeLinesHitLabels(display, labels, True)
        End If
        rejected.Add Array(bestX, bestY)
        retries = retries + 1
        ann.SetPosition2 oldPosition(0), oldPosition(1), oldPosition(2)
        If retries < NATIVE_LABEL_RETRY_LIMIT Then
            GoTo SearchNativeLabel
        End If
        Set textFormat = ann.GetTextFormat(SW_ANNOTATION_TEXT_FORMAT_FIRST)
        If textFormat.CharHeight > TEXT_HEIGHT_SMALL_M + DIMENSION_GEOMETRY_GAP_M Then
            textFormat.CharHeight = TEXT_HEIGHT_SMALL_M
            ann.SetTextFormat SW_ANNOTATION_TEXT_FORMAT_FIRST, False, textFormat
            d.GraphicsRedraw2
            Set rejected = New Collection
            retries = 0
            requiredWidth = 0
            requiredHeight = 0
            GoTo SearchNativeLabel
        End If
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Native leader cannot avoid the existing labels: " & display.GetNameForSelection
    End If
    Set data = display.GetDisplayData
    labels.Add NativeTextBounds(data)
    NativeCollectLines display, lines
End Sub

Private Function NativeLabelsByWidth(ByVal records As Collection) As Collection
    Dim result As New Collection, widths As Object, item As Variant, display As SldWorks.DisplayDimension
    Dim bounds As Variant, data As SldWorks.DisplayData, width As Double, index As Long, inserted As Boolean
    Set widths = CreateObject("Scripting.Dictionary")
    For Each item In records
        Set display = item
        Set data = display.GetDisplayData
        bounds = NativeTextBounds(data)
        width = bounds(2) - bounds(0)
        widths.Add display.GetNameForSelection, width
        inserted = False
        For index = 1 To result.Count
            If width > CDbl(widths(result(index).GetNameForSelection)) Then
                result.Add display, , index
                inserted = True
                Exit For
            End If
        Next index
        If Not inserted Then
            result.Add display
        End If
    Next item
    Set NativeLabelsByWidth = result
End Function

Private Function KeepReadableNativeText(ByVal display As SldWorks.DisplayDimension, ByVal region As Variant, ByVal obstacles As Collection, ByVal labels As Collection, ByVal lines As Collection) As Boolean
    Dim data As SldWorks.DisplayData, bounds As Variant, interior As Variant, item As Variant
    Set data = display.GetDisplayData
    bounds = NativeTextBounds(data)
    If Not NativeBoundsInside(bounds, region) Then
        Exit Function
    End If
    For Each item In obstacles
        If NativeBoxesOverlap(bounds, item, LEADER_GEOMETRY_CLEARANCE_M) Then
            Exit Function
        End If
    Next item
    For Each item In labels
        If NativeBoxesOverlap(bounds, item, LEADER_LABEL_CLEARANCE_M) Then
            Exit Function
        End If
    Next item
    interior = Array(bounds(0) + DIMENSION_GEOMETRY_GAP_M, bounds(1) + DIMENSION_GEOMETRY_GAP_M, bounds(2) - DIMENSION_GEOMETRY_GAP_M, bounds(3) - DIMENSION_GEOMETRY_GAP_M)
    For Each item In lines
        If NativeLeaderCrossesBox(item(0), item(1), item(2), item(3), interior) Then
            Exit Function
        End If
    Next item
    labels.Add bounds
    KeepReadableNativeText = True
End Function

