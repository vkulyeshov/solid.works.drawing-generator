Attribute VB_Name = "planar_layout"
Option Explicit

Private Sub RejectPlanarCandidate(ByVal reason As String)
    If Not planarLayoutRejections Is Nothing Then
        If Not planarLayoutRejections.Exists(reason) Then
            planarLayoutRejections.Add reason, 0
        End If
        planarLayoutRejections(reason) = CLng(planarLayoutRejections(reason)) + 1
    End If
End Sub

Private Function PlanarAuditCoordinates(ByVal values As Variant) As String
    Dim value As Variant
    For Each value In values
        If Len(PlanarAuditCoordinates) > 0 Then
            PlanarAuditCoordinates = PlanarAuditCoordinates & vbTab
        End If
        PlanarAuditCoordinates = PlanarAuditCoordinates & CStr(value)
    Next value
End Function

' A candidate contains the complete native dimension, not an independently moved label.
Private Function PlanarDisplayLines(ByVal data As SldWorks.DisplayData) As Collection
    Dim result As New Collection, index As Long, segment As Variant, arc As Variant
    Dim angle As Double, finish As Double, sweep As Double, radius As Double, first As Variant, last As Variant, stepIndex As Long, steps As Long
    For index = 0 To data.GetLineCount - 1
        segment = data.GetLineAtIndex2(index)
        result.Add Array(segment(DISPLAY_LINE_START_X), segment(DISPLAY_LINE_START_Y), segment(DISPLAY_LINE_END_X), segment(DISPLAY_LINE_END_Y))
    Next index
    For index = 0 To data.GetArcCount - 1
        arc = data.GetArcAtIndex2(index)
        angle = PlanarAngle(arc(DISPLAY_LINE_START_X) - arc(DISPLAY_ARC_CENTER_X), arc(DISPLAY_LINE_START_Y) - arc(DISPLAY_ARC_CENTER_Y))
        finish = PlanarAngle(arc(DISPLAY_LINE_END_X) - arc(DISPLAY_ARC_CENTER_X), arc(DISPLAY_LINE_END_Y) - arc(DISPLAY_ARC_CENTER_Y))
        sweep = finish - angle
        If (arc(DISPLAY_ARC_COUNTERCLOCKWISE) <> 0) = (arc(DISPLAY_ARC_NORMAL_Z) >= 0) Then
            Do While sweep < 0
                sweep = sweep + 2 * PI
            Loop
        Else
            Do While sweep > 0
                sweep = sweep - 2 * PI
            Loop
        End If
        radius = Sqr((arc(DISPLAY_LINE_START_X) - arc(DISPLAY_ARC_CENTER_X)) ^ 2 + (arc(DISPLAY_LINE_START_Y) - arc(DISPLAY_ARC_CENTER_Y)) ^ 2)
        steps = Int(Abs(sweep) / PLANAR_ARC_STEP_RAD) + 1
        first = Array(arc(DISPLAY_LINE_START_X), arc(DISPLAY_LINE_START_Y))
        For stepIndex = 1 To steps
            finish = angle + sweep * stepIndex / steps
            last = Array(arc(DISPLAY_ARC_CENTER_X) + radius * Cos(finish), arc(DISPLAY_ARC_CENTER_Y) + radius * Sin(finish))
            result.Add Array(first(0), first(1), last(0), last(1))
            first = last
        Next stepIndex
    Next index
    Set PlanarDisplayLines = result
End Function

Private Function PlanarTextPolygons(ByVal display As SldWorks.DisplayDimension, ByVal data As SldWorks.DisplayData, ByVal margin As Double, Optional ByVal printed As Boolean = False) As Collection
    Dim result As New Collection, index As Long, corner As Long, p As Variant, polygon(3) As Variant
    Dim height As Double, width As Double, angle As Double, x As Double, y As Double, text As String, leading As Long, offset As Double
    Dim face As String, originalWidth As Double, prefix As String, suffix As String
    Dim format As SldWorks.TextFormat, textRatio As Double, anchor As Variant
    textRatio = 1
    If printed Then
        Set format = display.GetAnnotation.GetTextFormat(SW_ANNOTATION_TEXT_FORMAT_FIRST)
        anchor = display.GetAnnotation.GetPosition
        If data.GetTextCount > 0 Then
            If data.GetTextHeightAtIndex(0) > DIRECTION_ZERO_TOLERANCE Then
                textRatio = format.CharHeight / data.GetTextHeightAtIndex(0)
            End If
        End If
    End If
    For index = 0 To data.GetTextCount - 1
        p = data.GetTextPositionAtIndex(index)
        height = data.GetTextHeightAtIndex(index)
        text = data.GetTextAtIndex(index)
        angle = data.GetTextAngleAtIndex(index)
        face = ""
        If printed Then
            ' Imported display data may retain the source font height until
            ' reload, while its text positions already follow the drawing.
            height = height * textRatio
            If display.IsHoleCallout Then
                p(0) = anchor(0) + (p(0) - anchor(0)) * textRatio
                p(1) = anchor(1) + (p(1) - anchor(1)) * textRatio
            End If
            face = format.TypeFaceName
            ' Some COM Boolean results use 1 rather than VBA's -1. Do not
            ' apply bitwise Not to that result: a callout is already composed.
            If index = 0 And display.IsHoleCallout = False Then
                originalWidth = PlanarMeasureTextWidth(display, text, height, True, face)
                prefix = display.GetText(SW_DIMENSION_TEXT_PREFIX)
                suffix = display.GetText(SW_DIMENSION_TEXT_SUFFIX)
                If Len(Trim$(prefix)) > 0 And InStr(text, Trim$(prefix)) = 0 Then
                    text = prefix & text
                End If
                If Len(Trim$(suffix)) > 0 And InStr(text, Trim$(suffix)) = 0 Then
                    text = text & suffix
                End If
                width = PlanarMeasureTextWidth(display, text, height, True, face)
                p(0) = p(0) - (width - originalWidth) / 2 * Cos(angle)
                p(1) = p(1) - (width - originalWidth) / 2 * Sin(angle)
            End If
        End If
        text = Replace(text, SW_DIMENSION_TOKEN_DIAMETER, ChrW$(DIAMETER_GLYPH_CODEPOINT))
        text = Replace(text, SW_DIMENSION_TOKEN_DEGREE, "°")
        leading = Len(text) - Len(LTrim$(text))
        offset = PlanarMeasureTextWidth(display, Left$(text, leading), height, printed, face)
        text = Trim$(text)
        If Len(text) = 0 Then
            GoTo NextPlanarText
        End If
        width = PlanarMeasureTextWidth(display, text, height, printed, face)
        p(0) = p(0) + offset * Cos(angle)
        p(1) = p(1) + offset * Sin(angle)
        For corner = 0 To 3
            x = -margin
            y = -margin
            If corner = 1 Or corner = 2 Then
                x = width + margin
            End If
            If corner >= 2 Then
                y = height + margin
            End If
            polygon(corner) = Array(p(0) + x * Cos(angle) - y * Sin(angle), p(1) + x * Sin(angle) + y * Cos(angle))
        Next corner
        result.Add polygon
NextPlanarText:
    Next index
    Set PlanarTextPolygons = result
End Function

Private Function PlanarPolygonsOverlap(ByVal a As Variant, ByVal b As Variant) As Boolean
    Dim index As Long, other As Long
    For index = LBound(a) To UBound(a)
        If ProfilePolygonLocation(a(index), b) <> PROFILE_POINT_OUTSIDE Then
            PlanarPolygonsOverlap = True
            Exit Function
        End If
        For other = LBound(b) To UBound(b)
            If PlanarSegmentsCross(a(index), a((index + 1) Mod (UBound(a) + 1)), b(other), b((other + 1) Mod (UBound(b) + 1))) Then
                PlanarPolygonsOverlap = True
                Exit Function
            End If
        Next other
    Next index
    PlanarPolygonsOverlap = ProfilePolygonLocation(b(LBound(b)), a) <> PROFILE_POINT_OUTSIDE
End Function

Private Function PlanarLinesHitText(ByVal lines As Collection, ByVal polygons As Collection) As Boolean
    Dim segment As Variant, polygon As Variant
    For Each segment In lines
        For Each polygon In polygons
            If ProfileLineCutsMaterial(Array(segment(0), segment(1)), Array(segment(2), segment(3)), Array(polygon)) Then
                PlanarLinesHitText = True
                Exit Function
            End If
        Next polygon
    Next segment
End Function

Private Function PlanarGraphicLength(ByVal lines As Collection) As Double
    Dim segment As Variant
    For Each segment In lines
        PlanarGraphicLength = PlanarGraphicLength + Sqr((segment(2) - segment(0)) ^ 2 + (segment(3) - segment(1)) ^ 2)
    Next segment
End Function

Private Function PlanarArrowTriangle(ByVal arrow As Variant) As Variant
    Dim dx As Double, dy As Double, magnitude As Double, length As Double, halfWidth As Double
    dx = arrow(DISPLAY_ARROW_DIRECTION_X)
    dy = arrow(DISPLAY_ARROW_DIRECTION_Y)
    magnitude = Sqr(dx * dx + dy * dy)
    If magnitude < PROFILE_MATERIAL_TOLERANCE_M Then
        PlanarArrowTriangle = Empty
        Exit Function
    End If
    dx = dx / magnitude
    dy = dy / magnitude
    length = arrow(DISPLAY_ARROW_WIDTH)
    halfWidth = arrow(DISPLAY_ARROW_HEIGHT) / 2
    PlanarArrowTriangle = Array(Array(arrow(0), arrow(1)), _
        Array(arrow(0) + length * dx - halfWidth * dy, arrow(1) + length * dy + halfWidth * dx), _
        Array(arrow(0) + length * dx + halfWidth * dy, arrow(1) + length * dy - halfWidth * dx))
End Function

Private Function PlanarArrowSpanFits(ByVal first As Variant, ByVal last As Variant) As Boolean
    Dim dx As Double, dy As Double
    dx = last(0) - first(0)
    dy = last(1) - first(1)
    PlanarArrowSpanFits = Sqr(dx * dx + dy * dy) >= first(DISPLAY_ARROW_WIDTH) + last(DISPLAY_ARROW_WIDTH) + PLANAR_ARROW_CLEARANCE_M
End Function

Private Function PlanarInsideArrowsFit(ByVal first As Variant, ByVal last As Variant) As Boolean
    Dim dx As Double, dy As Double, firstProjection As Double, lastProjection As Double
    Dim firstTriangle As Variant, lastTriangle As Variant
    If Not PlanarArrowSpanFits(first, last) Then
        Exit Function
    End If
    dx = last(0) - first(0)
    dy = last(1) - first(1)
    firstProjection = dx * first(DISPLAY_ARROW_DIRECTION_X) + dy * first(DISPLAY_ARROW_DIRECTION_Y)
    lastProjection = -dx * last(DISPLAY_ARROW_DIRECTION_X) - dy * last(DISPLAY_ARROW_DIRECTION_Y)
    If firstProjection <= 0 Or lastProjection <= 0 Then
        Exit Function
    End If
    firstTriangle = PlanarArrowTriangle(first)
    lastTriangle = PlanarArrowTriangle(last)
    If IsEmpty(firstTriangle) Or IsEmpty(lastTriangle) Then
        Exit Function
    End If
    PlanarInsideArrowsFit = Not PlanarPolygonsOverlap(firstTriangle, lastTriangle)
End Function

Private Function PlanarClipPolygon(ByVal polygon As Variant, ByVal bounds As Variant) As Variant
    Dim result As Variant, output As Collection, first As Variant, last As Variant, index As Long, edge As Long, axis As Long
    Dim coordinate As Double, firstInside As Boolean, lastInside As Boolean, t As Double, side As Long
    result = polygon
    For edge = OUTLINE_LEFT_INDEX To OUTLINE_TOP_INDEX
        axis = edge Mod 2
        side = 1
        If edge = OUTLINE_RIGHT_INDEX Or edge = OUTLINE_TOP_INDEX Then
            side = -1
        End If
        coordinate = bounds(edge)
        Set output = New Collection
        first = result(UBound(result))
        firstInside = (first(axis) - coordinate) * side >= 0
        For index = LBound(result) To UBound(result)
            last = result(index)
            lastInside = (last(axis) - coordinate) * side >= 0
            If firstInside <> lastInside Then
                t = (coordinate - first(axis)) / (last(axis) - first(axis))
                output.Add Array(first(0) + (last(0) - first(0)) * t, first(1) + (last(1) - first(1)) * t)
            End If
            If lastInside Then
                output.Add last
            End If
            first = last
            firstInside = lastInside
        Next index
        If output.Count = 0 Then
            PlanarClipPolygon = Empty
            Exit Function
        End If
        ReDim result(output.Count - 1)
        For index = 1 To output.Count
            result(index - 1) = output(index)
        Next index
    Next edge
    PlanarClipPolygon = result
End Function

Private Function PlanarPointBoundaryDistance(ByVal point As Variant, ByVal polygon As Variant) As Double
    Dim index As Long, a As Variant, b As Variant, dx As Double, dy As Double, lengthSquared As Double, t As Double, distance As Double
    PlanarPointBoundaryDistance = SEARCH_INITIAL_LIMIT
    For index = LBound(polygon) To UBound(polygon)
        a = polygon(index)
        b = polygon((index + 1) Mod (UBound(polygon) + 1))
        dx = b(0) - a(0)
        dy = b(1) - a(1)
        lengthSquared = dx * dx + dy * dy
        t = 0
        If lengthSquared > 0 Then
            t = ((point(0) - a(0)) * dx + (point(1) - a(1)) * dy) / lengthSquared
            If t < 0 Then
                t = 0
            ElseIf t > 1 Then
                t = 1
            End If
        End If
        distance = Sqr((a(0) + t * dx - point(0)) ^ 2 + (a(1) + t * dy - point(1)) ^ 2)
        If distance < PlanarPointBoundaryDistance Then
            PlanarPointBoundaryDistance = distance
        End If
    Next index
End Function

' A native sample costs a placement and a full redraw. Skip it only when the
' text placed at the requested point is certain to fail the same region or
' material check in PlanarCaptureCandidate: every printed text box lies within
' PLANAR_SAMPLE_TEXT_REACH_M of that point. Only non-centered text is placed
' at the requested point, so callers must not use this for centered text.
Private Function PlanarSampleCannotFit(ByVal view As SldWorks.View, ByVal region As Variant, ByVal x As Double, ByVal y As Double) As Boolean
    Dim reach As Double, polygon As Variant, point As Variant
    reach = DIMENSION_GEOMETRY_GAP_M + PLANAR_SAMPLE_TEXT_REACH_M
    If x < region(OUTLINE_LEFT_INDEX) - reach Or y < region(OUTLINE_BOTTOM_INDEX) - reach Or x > region(OUTLINE_RIGHT_INDEX) + reach Or y > region(OUTLINE_TOP_INDEX) + reach Then
        RejectPlanarCandidate "skipped: text point outside view region"
        PlanarSampleCannotFit = True
        Exit Function
    End If
    If Not planarLayoutVisibleMaterial.Exists(view.Name) Then
        Exit Function
    End If
    polygon = planarLayoutVisibleMaterial(view.Name)
    If IsEmpty(polygon) Then
        Exit Function
    End If
    point = Array(x, y)
    If ProfilePolygonLocation(point, polygon) = PROFILE_POINT_INSIDE Then
        If PlanarPointBoundaryDistance(point, polygon) > PLANAR_SAMPLE_TEXT_REACH_M Then
            RejectPlanarCandidate "skipped: text point inside material"
            PlanarSampleCannotFit = True
        End If
    End If
End Function

' The bounded pass stops after complete rows. Its candidate list is kept so
' that an expanded pass for the same view continues from the next row instead
' of sampling the same native placements again.
Private Sub RememberPlanarCandidates(ByVal displayName As String, ByVal candidates As Collection, ByVal nextRow As Long)
    If planarCandidateResume Is Nothing Then
        Exit Sub
    End If
    If planarCandidateResume.Exists(displayName) Then
        planarCandidateResume.Remove displayName
    End If
    planarCandidateResume.Add displayName, Array(candidates, nextRow)
End Sub

' The joint planar solver checks printed text, crossings and arrowheads
' together. The bounded semantic label passes are an optional pre-pass; with
' zero passes, only labels that already satisfy every check are kept.
Private Function PlanarSemanticPasses() As Long
    If Setting("PlanarLabelPrepass", DEFAULT_PLANAR_LABEL_PREPASS) = "1" Then
        PlanarSemanticPasses = LAYOUT_LABEL_PASSES
    End If
End Function

' Dimension lines: the arrow-to-arrow segment of a linear dimension and the
' sampled arc of an angular dimension.
Private Function PlanarDimensionBases(ByVal display As SldWorks.DisplayDimension) As Collection
    Dim data As SldWorks.DisplayData, lines As Collection, result As New Collection, first As Variant, last As Variant, index As Long
    Dim segment As Variant, dx As Double, dy As Double, length As Double
    Set data = display.GetDisplayData
    Set lines = PlanarDisplayLines(data)
    If display.Type2 = SW_DIMENSION_LINEAR And data.GetArrowHeadCount = 2 Then
        first = data.GetArrowHeadAtIndex2(0)
        last = data.GetArrowHeadAtIndex2(1)
        result.Add Array(first(0), first(1), last(0), last(1))
        ' Outside arrows extend the dimension line beyond the extension lines.
        dx = last(0) - first(0)
        dy = last(1) - first(1)
        length = Sqr(dx * dx + dy * dy)
        If length > PROFILE_MATERIAL_TOLERANCE_M Then
            For index = 1 To data.GetLineCount
                segment = lines(index)
                If Abs((segment(0) - first(0)) * dy - (segment(1) - first(1)) * dx) <= PROFILE_MATERIAL_TOLERANCE_M * length And Abs((segment(2) - first(0)) * dy - (segment(3) - first(1)) * dx) <= PROFILE_MATERIAL_TOLERANCE_M * length Then
                    result.Add segment
                End If
            Next index
        End If
    End If
    For index = data.GetLineCount + 1 To lines.Count
        result.Add lines(index)
    Next index
    Set PlanarDimensionBases = result
End Function

Private Function PlanarElapsedSeconds(ByVal started As Double) As Double
    PlanarElapsedSeconds = Timer - started
    If PlanarElapsedSeconds < 0 Then
        PlanarElapsedSeconds = PlanarElapsedSeconds + SECONDS_PER_DAY
    End If
End Function

Private Function PlanarCaptureCandidate(ByVal display As SldWorks.DisplayDimension, ByVal view As SldWorks.View, ByVal region As Variant, ByVal obstacles As Collection) As Object
    Dim data As SldWorks.DisplayData, result As Object, lines As Collection, polygons As Collection, bases As New Collection
    Dim bounds As Variant, textBounds As Variant, segment As Variant, obstacle As Variant, polygon As Variant, corner As Variant
    Dim first As Variant, last As Variant, index As Long, rawPolygon As Variant, viewOutline As Variant
    Dim arrow As Variant, triangle As Variant, edge As Long, heads As New Collection, printedText As Collection
    Set data = display.GetDisplayData
    bounds = NativeDimensionBounds(display)
    Set lines = PlanarDisplayLines(data)
    Set bases = PlanarDimensionBases(display)
    If data.GetArrowHeadCount = 2 And (display.Type2 = SW_DIMENSION_LINEAR Or display.Type2 = SW_DIMENSION_ANGULAR) Then
        first = data.GetArrowHeadAtIndex2(0)
        last = data.GetArrowHeadAtIndex2(1)
        If display.ArrowSide = SW_DIMENSION_ARROWS_INSIDE Then
            If Not PlanarInsideArrowsFit(first, last) Then
                RejectPlanarCandidate "inside arrowheads do not fit"
                Exit Function
            End If
        End If
    End If
    ' Closed arrowheads are part of the collision geometry, not just their tips.
    For index = 0 To data.GetArrowHeadCount - 1
        arrow = data.GetArrowHeadAtIndex2(index)
        If arrow(DISPLAY_ARROW_STYLE) = SW_ARROW_CLOSED Then
            triangle = PlanarArrowTriangle(arrow)
            If Not IsEmpty(triangle) Then
                heads.Add triangle
                For edge = 0 To 2
                    first = triangle(edge)
                    last = triangle((edge + 1) Mod 3)
                    lines.Add Array(first(0), first(1), last(0), last(1))
                Next edge
            End If
        End If
    Next index
    rawPolygon = planarLayoutMaterial(view.Name)
    viewOutline = planarLayoutOutlines(view.Name)
    For Each segment In bases
        If NativeLeaderCrossesBox(segment(0), segment(1), segment(2), segment(3), viewOutline) Then
            If ProfileLineCutsMaterial(Array(segment(0), segment(1)), Array(segment(2), segment(3)), Array(rawPolygon)) Then
                RejectPlanarCandidate "baseline crosses material"
                Exit Function
            End If
        End If
    Next segment
    For Each segment In lines
        NativeExpandBounds bounds, segment(0), segment(1)
        NativeExpandBounds bounds, segment(2), segment(3)
    Next segment
    ' Check the text as printed after reload: imported display data can keep a
    ' smaller source font height while the layout is being sampled.
    Set printedText = PlanarTextPolygons(display, data, 0, True)
    textBounds = NativeTextBounds(data)
    For Each polygon In printedText
        For Each corner In polygon
            NativeExpandBounds bounds, corner(0), corner(1)
            NativeExpandBounds textBounds, corner(0), corner(1)
        Next corner
    Next polygon
    If Not NativeBoundsInside(bounds, region) Then
        RejectPlanarCandidate "outside view region"
        Exit Function
    End If
    polygon = planarLayoutVisibleMaterial(view.Name)
    If Not IsEmpty(polygon) Then
        Set polygons = PlanarTextPolygons(display, data, LEADER_GEOMETRY_CLEARANCE_M, True)
        For Each corner In polygons
            If PlanarPolygonsOverlap(corner, polygon) Then
                RejectPlanarCandidate "text touches material"
                Exit Function
            End If
        Next corner
    End If
    For Each obstacle In obstacles
        If NativeBoxesOverlap(textBounds, obstacle, LEADER_LABEL_CLEARANCE_M) Then
            RejectPlanarCandidate "text touches note"
            Exit Function
        End If
        For Each segment In lines
            If NativeLeaderCrossesBox(segment(0), segment(1), segment(2), segment(3), obstacle) Then
                RejectPlanarCandidate "graphics cross note"
                Exit Function
            End If
        Next segment
    Next obstacle
    Set polygons = PlanarTextPolygons(display, data, -DIMENSION_GEOMETRY_GAP_M, True)
    If PlanarLinesHitText(lines, polygons) Then
        RejectPlanarCandidate "graphics cross own text"
        Exit Function
    End If
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "position", display.GetAnnotation.GetPosition
    result.Add "text", PlanarTextPolygons(display, data, LEADER_LABEL_CLEARANCE_M / 2, True)
    result.Add "lines", lines
    result.Add "bases", bases
    result.Add "heads", heads
    result.Add "bounds", bounds
    result.Add "name", display.GetNameForSelection
    result.Add "offsetText", display.OffsetText
    result.Add "centerText", display.CenterText
    result.Add "cost", PlanarGraphicLength(lines)
    result.Add "arrowSide", display.ArrowSide
    If data.GetArrowHeadCount = 2 And display.ArrowSide = SW_DIMENSION_ARROWS_OUTSIDE Then
        result("cost") = CDbl(result("cost")) + PLANAR_OUTSIDE_ARROW_PENALTY_M
    End If
    result.Add "value", display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue
    If display.Type2 = SW_DIMENSION_ANGULAR And data.GetArrowHeadCount = 2 Then
        obstacle = data.GetArcAtIndex2(0)
        first = data.GetArrowHeadAtIndex2(0)
        last = data.GetArrowHeadAtIndex2(1)
        result.Add "sector", PlanarArrowBisector(Array(obstacle(DISPLAY_ARC_CENTER_X), obstacle(DISPLAY_ARC_CENTER_Y)), first, last)
    End If
    Set PlanarCaptureCandidate = result
End Function

Private Sub AddPlanarLayoutCandidate(ByVal d As SldWorks.ModelDoc2, ByVal display As SldWorks.DisplayDimension, ByVal view As SldWorks.View, ByVal region As Variant, ByVal obstacles As Collection, ByVal candidates As Collection, ByVal x As Double, ByVal y As Double, ByVal expected As Double, Optional ByVal basePosition As Variant, Optional ByVal oppositeAngle As Boolean = False)
    Dim side As Long, firstSide As Long, lastSide As Long, requestedCenter As Boolean, requestedOffset As Boolean
    Dim data As SldWorks.DisplayData
    requestedCenter = display.CenterText
    requestedOffset = display.OffsetText
    ' Non-centered text, including offset text, is placed at (x, y) for every arrow side.
    If Not requestedCenter Then
        If PlanarSampleCannotFit(view, region, x, y) Then
            planarSampleSkips = planarSampleSkips + 1
            Exit Sub
        End If
    End If
    firstSide = display.ArrowSide
    lastSide = firstSide
    Set data = display.GetDisplayData
    If data.GetArrowHeadCount = 2 And (display.Type2 = SW_DIMENSION_LINEAR Or display.Type2 = SW_DIMENSION_ANGULAR) Then
        firstSide = SW_DIMENSION_ARROWS_INSIDE
        lastSide = SW_DIMENSION_ARROWS_OUTSIDE
        ' Linear reference spacing is invariant under label movement. Avoid
        ' expensive native inside samples when the heads cannot physically fit.
        If display.Type2 = SW_DIMENSION_LINEAR Then
            If Not PlanarArrowSpanFits(data.GetArrowHeadAtIndex2(0), data.GetArrowHeadAtIndex2(1)) Then
                firstSide = SW_DIMENSION_ARROWS_OUTSIDE
            End If
        End If
    End If
    For side = firstSide To lastSide
        display.ArrowSide = side
        If requestedOffset And Not IsMissing(basePosition) Then
            display.OffsetText = False
            display.CenterText = True
            display.GetAnnotation.SetPosition2 basePosition(0), basePosition(1), basePosition(2)
            d.GraphicsRedraw2
        End If
        display.OffsetText = requestedOffset
        display.CenterText = requestedCenter
        If IsMissing(basePosition) Then
            CapturePlanarLayoutVariant d, display, view, region, obstacles, candidates, x, y, expected, oppositeAngle:=oppositeAngle
        Else
            CapturePlanarLayoutVariant d, display, view, region, obstacles, candidates, x, y, expected, basePosition, oppositeAngle
        End If
    Next side
    display.OffsetText = requestedOffset
    display.CenterText = requestedCenter
End Sub

Private Sub CapturePlanarLayoutVariant(ByVal d As SldWorks.ModelDoc2, ByVal display As SldWorks.DisplayDimension, ByVal view As SldWorks.View, ByVal region As Variant, ByVal obstacles As Collection, ByVal candidates As Collection, ByVal x As Double, ByVal y As Double, ByVal expected As Double, Optional ByVal basePosition As Variant, Optional ByVal oppositeAngle As Boolean = False)
    Dim candidate As Object, existing As Object, annotation As SldWorks.Annotation, first As Variant, last As Variant
    Dim firstBase As Variant, lastBase As Variant, index As Long
    Dim requestedCenter As Boolean, requestedOffset As Boolean
    Set annotation = display.GetAnnotation
    requestedCenter = display.CenterText
    requestedOffset = display.OffsetText
    planarSampleCount = planarSampleCount + 1
    AsDrawing(d).ActivateView ""
    annotation.SetPosition2 x, y, 0
    ' Moving native text can turn centering off; apply the requested text mode last.
    display.OffsetText = requestedOffset
    display.CenterText = requestedCenter
    d.GraphicsRedraw2
    DoEvents
    If Abs(display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue - expected) > LEADER_VALUE_TOLERANCE_M Then
        RejectPlanarCandidate "native value changed"
        Exit Sub
    End If
    Set candidate = PlanarCaptureCandidate(display, view, region, obstacles)
    If Not candidate Is Nothing Then
        candidate("centerText") = requestedCenter
        candidate("offsetText") = requestedOffset
        candidate.Add "oppositeAngle", oppositeAngle
        ' Centered native text can return a different point from the placement request.
        candidate.Add "controlPoint", Array(x, y, 0)
        If Not IsMissing(basePosition) Then
            candidate.Add "baseControlPoint", basePosition
        Else
            candidate.Add "baseControlPoint", Array(x, y, 0)
        End If
        first = candidate("position")
        For Each existing In candidates
            last = existing("position")
            If candidate("arrowSide") = existing("arrowSide") And candidate("offsetText") = existing("offsetText") And candidate("centerText") = existing("centerText") And candidate("oppositeAngle") = existing("oppositeAngle") Then
                If Abs(first(0) - last(0)) < NATIVE_GEOMETRY_KEY_TOLERANCE_M And Abs(first(1) - last(1)) < NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
                    firstBase = candidate("baseControlPoint")
                    lastBase = existing("baseControlPoint")
                    If Abs(firstBase(0) - lastBase(0)) < NATIVE_GEOMETRY_KEY_TOLERANCE_M And Abs(firstBase(1) - lastBase(1)) < NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
                        Exit Sub
                    End If
                End If
            End If
        Next existing
        PlanarInsertCandidate candidates, candidate
    End If
End Sub

Private Sub PlanarInsertCandidate(ByVal candidates As Collection, ByVal candidate As Object)
    Dim index As Long
    For index = 1 To candidates.Count
        If CDbl(candidate("cost")) < CDbl(candidates(index)("cost")) Then
            candidates.Add candidate, , index
            Exit Sub
        End If
    Next index
    candidates.Add candidate
End Sub

Private Function PlanarShiftPoint(ByVal point As Variant, ByVal x As Double, ByVal y As Double) As Variant
    If UBound(point) >= 2 Then
        PlanarShiftPoint = Array(point(0) + x, point(1) + y, point(2))
    Else
        PlanarShiftPoint = Array(point(0) + x, point(1) + y)
    End If
End Function

' A linear dimension moved along its normal keeps its shape: every graphic moves by the same
' offset except the extension-line feet, which stay at the measured points.
Private Function PlanarTranslateCandidate(ByVal template As Object, ByVal first As Variant, ByVal last As Variant, ByVal ux As Double, ByVal uy As Double, ByVal nx As Double, ByVal ny As Double, ByVal delta As Double) As Object
    Dim result As Object, key As Variant, item As Variant, shifted As Collection, polygon As Variant, moved() As Variant, index As Long
    Dim x As Double, y As Double, bounds As Variant, segment As Variant, startX As Double, startY As Double, endX As Double, endY As Double
    x = nx * delta
    y = ny * delta
    Set result = CreateObject("Scripting.Dictionary")
    For Each key In template.Keys
        Select Case CStr(key)
            Case "text", "heads"
                Set shifted = New Collection
                For Each polygon In template(key)
                    ReDim moved(UBound(polygon))
                    For index = 0 To UBound(polygon)
                        moved(index) = PlanarShiftPoint(polygon(index), x, y)
                    Next index
                    shifted.Add moved
                Next polygon
                result.Add key, shifted
            Case "bases"
                Set shifted = New Collection
                For Each segment In template(key)
                    shifted.Add Array(segment(0) + x, segment(1) + y, segment(2) + x, segment(3) + y)
                Next segment
                result.Add key, shifted
            Case "lines"
                Set shifted = New Collection
                For Each segment In template(key)
                    startX = segment(0)
                    startY = segment(1)
                    endX = segment(2)
                    endY = segment(3)
                    If Not PlanarExtensionFoot(startX, startY, first, last, ux, uy, nx, ny) Then
                        startX = startX + x
                        startY = startY + y
                    End If
                    If Not PlanarExtensionFoot(endX, endY, first, last, ux, uy, nx, ny) Then
                        endX = endX + x
                        endY = endY + y
                    End If
                    shifted.Add Array(startX, startY, endX, endY)
                Next segment
                result.Add key, shifted
            Case "position", "controlPoint", "baseControlPoint"
                result.Add key, PlanarShiftPoint(template(key), x, y)
            Case "bounds", "cost"
            Case Else
                result.Add key, template(key)
        End Select
    Next key
    bounds = Array(SEARCH_INITIAL_LIMIT, SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT)
    For Each segment In result("lines")
        NativeExpandBounds bounds, segment(0), segment(1)
        NativeExpandBounds bounds, segment(2), segment(3)
    Next segment
    For Each polygon In result("text")
        For Each item In polygon
            NativeExpandBounds bounds, item(0), item(1)
        Next item
    Next polygon
    result.Add "bounds", bounds
    result.Add "cost", CDbl(template("cost")) - PlanarGraphicLength(template("lines")) + PlanarGraphicLength(result("lines"))
    Set PlanarTranslateCandidate = result
End Function

' An extension line starts at a measured point, possibly after a small gap along the normal.
Private Function PlanarExtensionFoot(ByVal x As Double, ByVal y As Double, ByVal first As Variant, ByVal last As Variant, ByVal ux As Double, ByVal uy As Double, ByVal nx As Double, ByVal ny As Double) As Boolean
    Dim point As Variant
    For Each point In Array(first, last)
        If Abs((x - point(0)) * ux + (y - point(1)) * uy) <= PLANAR_MODEL_ALONG_TOLERANCE_M And Abs((x - point(0)) * nx + (y - point(1)) * ny) <= PLANAR_MODEL_FOOT_TOLERANCE_M Then
            PlanarExtensionFoot = True
            Exit Function
        End If
    Next point
End Function

' The checks of PlanarCaptureCandidate on modelled geometry. Text keeps the material clearance
' by growing the stored text boxes from their stored margin to the capture margin.
Private Function PlanarModelCandidateValid(ByVal candidate As Object, ByVal view As SldWorks.View, ByVal region As Variant, ByVal obstacles As Collection) As Boolean
    Dim segment As Variant, polygon As Variant, material As Variant, obstacle As Variant, textBounds As Variant, corner As Variant
    Dim rawPolygon As Variant, viewOutline As Variant
    rawPolygon = planarLayoutMaterial(view.Name)
    viewOutline = planarLayoutOutlines(view.Name)
    For Each segment In candidate("bases")
        If NativeLeaderCrossesBox(segment(0), segment(1), segment(2), segment(3), viewOutline) Then
            If ProfileLineCutsMaterial(Array(segment(0), segment(1)), Array(segment(2), segment(3)), Array(rawPolygon)) Then
                Exit Function
            End If
        End If
    Next segment
    If Not NativeBoundsInside(candidate("bounds"), region) Then
        Exit Function
    End If
    material = planarLayoutVisibleMaterial(view.Name)
    textBounds = Array(SEARCH_INITIAL_LIMIT, SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT)
    For Each polygon In candidate("text")
        If Not IsEmpty(material) Then
            If PlanarPolygonsOverlap(PlanarGrowRectangle(polygon, LEADER_GEOMETRY_CLEARANCE_M - LEADER_LABEL_CLEARANCE_M / 2), material) Then
                Exit Function
            End If
        End If
        For Each corner In polygon
            NativeExpandBounds textBounds, corner(0), corner(1)
        Next corner
    Next polygon
    For Each obstacle In obstacles
        If NativeBoxesOverlap(textBounds, obstacle, LEADER_LABEL_CLEARANCE_M) Then
            Exit Function
        End If
        For Each segment In candidate("lines")
            If NativeLeaderCrossesBox(segment(0), segment(1), segment(2), segment(3), obstacle) Then
                Exit Function
            End If
        Next segment
    Next obstacle
    PlanarModelCandidateValid = True
End Function

Private Function PlanarGrowRectangle(ByVal polygon As Variant, ByVal amount As Double) As Variant
    Dim ax As Double, ay As Double, bx As Double, by As Double, length As Double, result(3) As Variant, p As Variant
    Dim p0 As Variant, p1 As Variant, p3 As Variant
    p0 = polygon(0)
    p1 = polygon(1)
    p3 = polygon(3)
    ax = p1(0) - p0(0)
    ay = p1(1) - p0(1)
    length = Sqr(ax * ax + ay * ay)
    If length > 0 Then
        ax = ax / length
        ay = ay / length
    End If
    bx = p3(0) - p0(0)
    by = p3(1) - p0(1)
    length = Sqr(bx * bx + by * by)
    If length > 0 Then
        bx = bx / length
        by = by / length
    End If
    p = polygon(0)
    result(0) = Array(p(0) - (ax + bx) * amount, p(1) - (ay + by) * amount)
    p = polygon(1)
    result(1) = Array(p(0) + (ax - bx) * amount, p(1) + (ay - by) * amount)
    p = polygon(2)
    result(2) = Array(p(0) + (ax + bx) * amount, p(1) + (ay + by) * amount)
    p = polygon(3)
    result(3) = Array(p(0) + (bx - ax) * amount, p(1) + (by - ay) * amount)
    PlanarGrowRectangle = result
End Function

Private Sub RestorePlanarAngularSector(ByVal d As SldWorks.ModelDoc2, ByVal display As SldWorks.DisplayDimension, ByVal candidate As Object)
    Dim data As SldWorks.DisplayData, arc As Variant, first As Variant, last As Variant, angle As Double, changed As Boolean, attempt As Long
    SetPlanarAngleRepresentation d, display, CDbl(candidate("value"))
    For attempt = 1 To PLANAR_ANGLE_SECTOR_VARIANTS
        Set data = display.GetDisplayData
        arc = data.GetArcAtIndex2(0)
        first = data.GetArrowHeadAtIndex2(0)
        last = data.GetArrowHeadAtIndex2(1)
        angle = PlanarArrowBisector(Array(arc(DISPLAY_ARC_CENTER_X), arc(DISPLAY_ARC_CENTER_Y)), first, last)
        If Cos(angle - CDbl(candidate("sector"))) > 0 Then
            Exit Sub
        End If
        SelectPlanarAngularDimension d, display
        changed = display.VerticallyOppositeAngle
        d.GraphicsRedraw2
        If Not changed Then
            SetPlanarAngleRepresentation d, display, PI - CDbl(candidate("value"))
            SetPlanarAngleRepresentation d, display, CDbl(candidate("value"))
        End If
    Next attempt
    Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot restore the selected native angular sector."
End Sub

' Keep angular text on its measured arc. Free text can add an almost complete circular leader.
Private Function PlanarAngularCandidates(ByVal d As SldWorks.ModelDoc2, ByVal display As SldWorks.DisplayDimension, ByVal view As SldWorks.View, ByVal region As Variant, ByVal obstacles As Collection) As Collection
    Dim result As New Collection, representation As Long, opposite As Long, row As Long, changed As Boolean
    Dim original As Double, value As Double, data As SldWorks.DisplayData, arc As Variant, cx As Double, cy As Double, angle As Double, radius As Double
    Dim first As Variant, last As Variant
    original = display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue
    For representation = 0 To 1
        value = original
        If representation > 0 Then
            value = PI - original
        End If
        SetPlanarAngleRepresentation d, display, value
        For opposite = 0 To 1
            If opposite > 0 Then
                SelectPlanarAngularDimension d, display
                changed = display.VerticallyOppositeAngle
                If Not changed Then
                    Exit For
                End If
            End If
            display.OffsetText = False
            display.CenterText = True
            display.SetBrokenLeader2 False, SW_SOLID_LEADER_ALIGNED_TEXT
            d.GraphicsRedraw2
            Set data = display.GetDisplayData
            If data.GetArcCount > 0 Then
                arc = data.GetArcAtIndex2(0)
                cx = arc(DISPLAY_ARC_CENTER_X)
                cy = arc(DISPLAY_ARC_CENTER_Y)
                first = data.GetArrowHeadAtIndex2(0)
                last = data.GetArrowHeadAtIndex2(1)
                angle = PlanarArrowBisector(Array(cx, cy), first, last)
                For row = 0 To PLANAR_LAYOUT_ROWS
                    radius = GOST_OBJECT_OFFSET_M + row * PLANAR_LAYOUT_RADIAL_PITCH_M
                    display.SetBrokenLeader2 False, SW_SOLID_LEADER_ALIGNED_TEXT
                    display.OffsetText = False
                    display.CenterText = True
                    AddPlanarLayoutCandidate d, display, view, region, obstacles, result, cx + radius * Cos(angle), cy + radius * Sin(angle), expected:=value, oppositeAngle:=(opposite > 0)
                Next row
            End If
            If opposite > 0 Then
                SelectPlanarAngularDimension d, display
                Call display.VerticallyOppositeAngle
            End If
        Next opposite
        SetPlanarAngleRepresentation d, display, original
    Next representation
    Set PlanarAngularCandidates = result
End Function

Private Function PlanarDimensionCandidates(ByVal d As SldWorks.ModelDoc2, ByVal display As SldWorks.DisplayDimension, ByVal view As SldWorks.View, ByVal region As Variant, ByVal obstacles As Collection) As Collection
    Dim result As New Collection, request As Object, first As Variant, last As Variant, bounds As Variant, position As Variant
    Dim dx As Double, dy As Double, length As Double, nx As Double, ny As Double, cx As Double, cy As Double, distance As Double, shift As Double
    Dim row As Long, side As Long, along As Long, direction As Long, angle As Double, value As Double, width As Double
    Dim referenceFirst As Variant, referenceLast As Variant, ex As Double, ey As Double, denominator As Double, t As Double
    Dim basePosition As Variant, annotation As SldWorks.Annotation, offsetSide As Long
    Dim data As SldWorks.DisplayData, arc As Variant
    Dim opposite As Long, oppositeLimit As Long, changed As Boolean, oppositeChanged As Boolean, centered As Long, centeredLimit As Long, directionLimit As Long, bisector As Double
    Dim displayName As String, firstRow As Long, nextRow As Long, stored As Variant, sampleX As Double, sampleY As Double
    Dim templates As Collection, target As Collection, template As Variant, derived As Object, templateRows(1) As Long
    value = display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue
    displayName = display.GetNameForSelection
    nextRow = PLANAR_LAYOUT_ROWS + 1
    ' An expanded pass continues the rows of the bounded pass for this view.
    If Not planarCandidateResume Is Nothing Then
        If planarCandidateResume.Exists(displayName) Then
            stored = planarCandidateResume(displayName)
            Set result = stored(PLANAR_RESUME_CANDIDATES_INDEX)
            firstRow = stored(PLANAR_RESUME_NEXT_ROW_INDEX)
            If firstRow > PLANAR_LAYOUT_ROWS Then
                Set PlanarDimensionCandidates = result
                Exit Function
            End If
        End If
    End If
    If display.Type2 = SW_DIMENSION_ANGULAR Then
        Set result = PlanarAngularCandidates(d, display, view, region, obstacles)
        RememberPlanarCandidates displayName, result, nextRow
        Set PlanarDimensionCandidates = result
        Exit Function
    End If
    display.OffsetText = False
    display.CenterText = False
    display.SetBrokenLeader2 False, SW_SOLID_LEADER_ALIGNED_TEXT
    bounds = NativeTextBounds(display.GetDisplayData)
    width = Sqr((bounds(2) - bounds(0)) ^ 2 + (bounds(3) - bounds(1)) ^ 2)
    If PlanarOwnsDimension(display) Then
        Set request = planarDimensionRequests(display.GetNameForSelection)
        first = PatternSheetPoint(view, request("first"))
        last = PatternSheetPoint(view, request("last"))
        dx = last(0) - first(0)
        dy = last(1) - first(1)
        If request("kind") = "hole-location" Then
            dx = 0
            dy = 1
            If CLng(request("axis")) = SCREEN_X_AXIS_INDEX Then
                dx = 1
                dy = 0
            End If
        End If
        length = Sqr(dx * dx + dy * dy)
        dx = dx / length
        dy = dy / length
        nx = dy
        ny = -dx
        cx = (first(0) + last(0)) / 2
        cy = (first(1) + last(1)) / 2
        If request("kind") = "angle" Then
            cx = first(0)
            cy = first(1)
            Set data = display.GetDisplayData
            If data.GetArcCount > 0 Then
                arc = data.GetArcAtIndex2(0)
                cx = arc(DISPLAY_ARC_CENTER_X)
                cy = arc(DISPLAY_ARC_CENTER_Y)
            End If
        Else
            ' The model places only the first row natively. Farther rows translate those
            ' templates along the normal and stretch their extension lines.
            Set templates = New Collection
            templates.Add New Collection
            templates.Add New Collection
            templateRows(0) = -1
            templateRows(1) = -1
            For row = firstRow To PLANAR_LAYOUT_ROWS
                For side = 0 To 1
                    distance = (GOST_OBJECT_OFFSET_M + row * PLANAR_LAYOUT_RADIAL_PITCH_M) * (1 - 2 * side)
                    Set target = result
                    If planarModelCandidates Then
                        ' A side samples natively until one row yields templates; farther rows are modelled.
                        If templateRows(side) >= 0 Then
                            For Each template In templates(side + 1)
                                Set derived = PlanarTranslateCandidate(template, first, last, dx, dy, nx, ny, (row - templateRows(side)) * PLANAR_LAYOUT_RADIAL_PITCH_M * (1 - 2 * side))
                                If PlanarModelCandidateValid(derived, view, region, obstacles) Then
                                    PlanarInsertCandidate result, derived
                                    planarModelDerived = planarModelDerived + 1
                                End If
                            Next template
                            GoTo NextPlanarCandidateSide
                        End If
                        ' An edge dimension line that the capture would reject for cutting the material is not sampled.
                        If request("kind") = "edge" Then
                            If NativeLeaderCrossesBox(first(0) + nx * distance, first(1) + ny * distance, last(0) + nx * distance, last(1) + ny * distance, planarLayoutOutlines(view.Name)) Then
                                If ProfileLineCutsMaterial(Array(first(0) + nx * distance, first(1) + ny * distance), Array(last(0) + nx * distance, last(1) + ny * distance), Array(planarLayoutMaterial(view.Name))) Then
                                    planarSampleSkips = planarSampleSkips + 1
                                    GoTo NextPlanarCandidateSide
                                End If
                            End If
                        End If
                        Set target = templates(side + 1)
                    End If
                    For along = -1 To 1
                        shift = along * (length / 2 + width + GOST_ARROW_LENGTH_M)
                        If request("kind") = "hole-location" Then
                            shift = along * width
                        End If
                        AddPlanarLayoutCandidate d, display, view, region, obstacles, target, cx + nx * distance + dx * shift, cy + ny * distance + dy * shift, value
                        If request("kind") = "hole-location" Then
                            bounds = view.GetOutline
                            If CLng(request("axis")) = SCREEN_X_AXIS_INDEX Then
                                AddPlanarLayoutCandidate d, display, view, region, obstacles, result, cx + shift, bounds(OUTLINE_BOTTOM_INDEX) - Abs(distance), value
                                AddPlanarLayoutCandidate d, display, view, region, obstacles, result, cx + shift, bounds(OUTLINE_TOP_INDEX) + Abs(distance), value
                            Else
                                AddPlanarLayoutCandidate d, display, view, region, obstacles, result, bounds(OUTLINE_LEFT_INDEX) - Abs(distance), cy + shift, value
                                AddPlanarLayoutCandidate d, display, view, region, obstacles, result, bounds(OUTLINE_RIGHT_INDEX) + Abs(distance), cy + shift, value
                            End If
                        End If
                    Next along
                    If request("kind") = "edge" Then
                        ' Each offset variant below places this base itself before sampling.
                        Set annotation = display.GetAnnotation
                        basePosition = Array(cx + nx * distance, cy + ny * distance, 0)
                        For along = -1 To 1
                            For offsetSide = 0 To 1
                                shift = along * (length / 2 + width + GOST_ARROW_LENGTH_M)
                                sampleX = cx + nx * (distance + (1 - 2 * offsetSide) * COORDINATE_LANE_PITCH_M) + dx * shift
                                sampleY = cy + ny * (distance + (1 - 2 * offsetSide) * COORDINATE_LANE_PITCH_M) + dy * shift
                                ' Offset text is non-centered and placed at the sample point; skip the base placement too.
                                If PlanarSampleCannotFit(view, region, sampleX, sampleY) Then
                                    planarSampleSkips = planarSampleSkips + 1
                                Else
                                    display.OffsetText = False
                                    display.CenterText = True
                                    annotation.SetPosition2 basePosition(0), basePosition(1), basePosition(2)
                                    d.GraphicsRedraw2
                                    display.OffsetText = True
                                    display.CenterText = False
                                    display.SetBrokenLeader2 False, SW_BROKEN_LEADER_HORIZONTAL_TEXT
                                    AddPlanarLayoutCandidate d, display, view, region, obstacles, target, sampleX, sampleY, value, basePosition
                                End If
                            Next offsetSide
                        Next along
                        display.OffsetText = False
                        display.CenterText = False
                        display.SetBrokenLeader2 False, SW_SOLID_LEADER_ALIGNED_TEXT
                    End If
NextPlanarCandidateSide:
                Next side
                If planarModelCandidates Then
                    For side = 0 To 1
                        If templateRows(side) < 0 And templates(side + 1).Count > 0 Then
                            templateRows(side) = row
                            For Each template In templates(side + 1)
                                PlanarInsertCandidate result, template
                            Next template
                        End If
                    Next side
                ElseIf planarCandidateBudget > 0 And result.Count >= planarCandidateBudget Then
                    nextRow = row + 1
                    GoTo PlanarCandidatesReady
                End If
            Next row
            RememberPlanarCandidates displayName, result, nextRow
            Set PlanarDimensionCandidates = result
            Exit Function
        End If
    Else
        bounds = view.GetOutline
        cx = (bounds(0) + bounds(2)) / 2
        cy = (bounds(1) + bounds(3)) / 2
        display.SolidLeader = False
        display.SetBrokenLeader2 False, SW_BROKEN_LEADER_HORIZONTAL_TEXT
        display.SetBentLeaderLength False, LEADER_SHOULDER_LENGTH_M
    End If
    If display.Type2 = SW_DIMENSION_ANGULAR Then
        oppositeLimit = 1
        centeredLimit = 1
    End If
    For opposite = 0 To oppositeLimit
        If opposite > 0 Then
            d.ClearSelection2 True
            display.GetAnnotation.Select3 False, Nothing
            changed = display.VerticallyOppositeAngle
            d.GraphicsRedraw2
            If Not changed Then
                Exit For
            End If
            oppositeChanged = True
        End If
        For centered = 0 To centeredLimit
            display.CenterText = display.Type2 = SW_DIMENSION_ANGULAR And centered = 0
            directionLimit = PLANAR_LAYOUT_DIRECTIONS - 1
            If display.Type2 = SW_DIMENSION_ANGULAR And centered = 0 Then
                d.GraphicsRedraw2
                Set data = display.GetDisplayData
                arc = data.GetArcAtIndex2(0)
                bisector = PlanarAngle(arc(DISPLAY_LINE_START_X) + arc(DISPLAY_LINE_END_X) - 2 * cx, arc(DISPLAY_LINE_START_Y) + arc(DISPLAY_LINE_END_Y) - 2 * cy)
                directionLimit = 0
            End If
            ' Rows resume only for a single opposite/centered pass; see PlanarCandidatesReady.
            For row = firstRow To PLANAR_LAYOUT_ROWS
                distance = GOST_OBJECT_OFFSET_M + row * PLANAR_LAYOUT_RADIAL_PITCH_M
                For direction = 0 To directionLimit
                    angle = direction * 2 * PI / PLANAR_LAYOUT_DIRECTIONS
                    If directionLimit = 0 Then
                        angle = bisector
                    End If
                    AddPlanarLayoutCandidate d, display, view, region, obstacles, result, cx + distance * Cos(angle), cy + distance * Sin(angle), expected:=value, oppositeAngle:=(opposite > 0)
                Next direction
                If planarCandidateBudget > 0 And result.Count >= planarCandidateBudget Then
                    nextRow = row + 1
                    GoTo PlanarCandidatesReady
                End If
            Next row
        Next centered
    Next opposite
PlanarCandidatesReady:
    If oppositeChanged Then
        d.ClearSelection2 True
        display.GetAnnotation.Select3 False, Nothing
        changed = display.VerticallyOppositeAngle
        d.GraphicsRedraw2
    End If
    ' A row index alone cannot resume nested opposite/centered passes.
    If oppositeLimit = 0 And centeredLimit = 0 Then
        RememberPlanarCandidates displayName, result, nextRow
    End If
    Set PlanarDimensionCandidates = result
End Function

Private Function PlanarCandidatesCompatible(ByVal first As Object, ByVal last As Object) As Boolean
    Dim a As Variant, b As Variant, firstText As Collection, lastText As Collection, bases As Collection, lines As Collection
    ' Bounds contain every line, arrowhead and printed text box; separated
    ' bounds cannot produce any of the conflicts below.
    If Not NativeBoxesOverlap(first("bounds"), last("bounds"), PLANAR_COMPATIBILITY_CLEARANCE_M) Then
        PlanarCandidatesCompatible = True
        Exit Function
    End If
    Set firstText = first("text")
    Set lastText = last("text")
    For Each a In firstText
        For Each b In lastText
            If PlanarPolygonsOverlap(a, b) Then
                Exit Function
            End If
        Next b
    Next a
    If PlanarLinesHitText(first("lines"), lastText) Or PlanarLinesHitText(last("lines"), firstText) Then
        Exit Function
    End If
    Set bases = first("bases")
    Set lines = last("bases")
    For Each a In bases
        For Each b In lines
            If PlanarDimensionLineConflict(a, b) Then
                Exit Function
            End If
        Next b
    Next a
    Set bases = last("bases")
    Set lines = first("bases")
    For Each a In bases
        For Each b In lines
            If PlanarDimensionLineConflict(a, b) Then
                Exit Function
            End If
        Next b
    Next a
    ' DSTU GOST 2.307: a dimension line must not cross the extension, leader or
    ' dimension lines of another dimension.
    If planarStrictCrossings Then
        If PlanarBasesCrossLines(first("bases"), last("lines")) Or PlanarBasesCrossLines(last("bases"), first("lines")) Then
            Exit Function
        End If
    End If
    If PlanarHeadsOverlap(first("heads"), last("heads")) Then
        Exit Function
    End If
    PlanarCandidatesCompatible = True
End Function

Private Function PlanarBasesCrossLines(ByVal bases As Collection, ByVal lines As Collection) As Boolean
    Dim a As Variant, b As Variant
    For Each a In bases
        For Each b In lines
            If PlanarSegmentsCross(Array(a(0), a(1)), Array(a(2), a(3)), Array(b(0), b(1)), Array(b(2), b(3))) Then
                PlanarBasesCrossLines = True
                Exit Function
            End If
        Next b
    Next a
End Function

' Arrowheads that only touch at a shared extension line are allowed; shrink
' both outlines toward their centroids before testing a real overlap.
Private Function PlanarHeadsOverlap(ByVal firstHeads As Collection, ByVal lastHeads As Collection) As Boolean
    Dim a As Variant, b As Variant
    For Each a In firstHeads
        For Each b In lastHeads
            If PlanarPolygonsOverlap(PlanarShrinkTriangle(a), PlanarShrinkTriangle(b)) Then
                PlanarHeadsOverlap = True
                Exit Function
            End If
        Next b
    Next a
End Function

Private Function PlanarShrinkTriangle(ByVal triangle As Variant) As Variant
    Dim cx As Double, cy As Double, index As Long, result(2) As Variant, p As Variant
    For index = 0 To 2
        p = triangle(index)
        cx = cx + p(0) / 3
        cy = cy + p(1) / 3
    Next index
    For index = 0 To 2
        p = triangle(index)
        result(index) = Array(cx + (p(0) - cx) * PLANAR_HEAD_OVERLAP_SHRINK, cy + (p(1) - cy) * PLANAR_HEAD_OVERLAP_SHRINK)
    Next index
    PlanarShrinkTriangle = result
End Function

Private Function PlanarDimensionLineConflict(ByVal a As Variant, ByVal b As Variant) As Boolean
    Dim dx As Double, dy As Double, length As Double, first As Double, last As Double
    If PlanarSegmentsCross(Array(a(0), a(1)), Array(a(2), a(3)), Array(b(0), b(1)), Array(b(2), b(3))) Then
        PlanarDimensionLineConflict = True
        Exit Function
    End If
    dx = a(2) - a(0)
    dy = a(3) - a(1)
    length = Sqr(dx * dx + dy * dy)
    If length < PROFILE_MATERIAL_TOLERANCE_M Then
        Exit Function
    End If
    If Abs((b(0) - a(0)) * dy - (b(1) - a(1)) * dx) > PROFILE_MATERIAL_TOLERANCE_M * length Or Abs((b(2) - a(0)) * dy - (b(3) - a(1)) * dx) > PROFILE_MATERIAL_TOLERANCE_M * length Then
        Exit Function
    End If
    first = ((b(0) - a(0)) * dx + (b(1) - a(1)) * dy) / length
    last = ((b(2) - a(0)) * dx + (b(3) - a(1)) * dy) / length
    PlanarDimensionLineConflict = MinD(length, -MinD(-first, -last)) - -MinD(0, -MinD(first, last)) > DIMENSION_GEOMETRY_GAP_M
End Function

' Candidate IDs are scoped to one view solve; no geometry is cached across jobs.
Private Function PlanarCandidatesCompatibleCached(ByVal first As Object, ByVal last As Object) As Boolean
    Dim key As String, reverseKey As String, result As Boolean
    key = first("id") & vbTab & last("id")
    If planarLayoutCompatibility.Exists(key) Then
        PlanarCandidatesCompatibleCached = CBool(planarLayoutCompatibility(key))
        Exit Function
    End If
    result = PlanarCandidatesCompatible(first, last)
    planarLayoutCompatibility.Add key, result
    reverseKey = last("id") & vbTab & first("id")
    planarLayoutCompatibility.Add reverseKey, result
    PlanarCandidatesCompatibleCached = result
End Function

' Forward checking prevents an early compact placement from trapping the remaining labels.
Private Function SolvePlanarCandidates(ByVal remaining As Collection, ByVal chosen As Collection, ByRef steps As Long) As Boolean
    Dim record As Object, candidate As Object, existing As Object, available As Collection, options As Collection
    Dim best As Object, bestOptions As Collection, nextRemaining As Collection, compatible As Boolean
    steps = steps + 1
    If steps > planarSearchLimit Then
        Exit Function
    End If
    If remaining.Count = 0 Then
        SolvePlanarCandidates = True
        Exit Function
    End If
    For Each record In remaining
        Set available = New Collection
        Set options = record("candidates")
        For Each candidate In options
            compatible = True
            For Each existing In chosen
                If Not PlanarCandidatesCompatibleCached(candidate, existing) Then
                    compatible = False
                    Exit For
                End If
            Next existing
            If compatible Then
                available.Add candidate
            End If
        Next candidate
        If available.Count = 0 Then
            Exit Function
        End If
        If best Is Nothing Then
            Set best = record
            Set bestOptions = available
        ElseIf available.Count < bestOptions.Count Then
            Set best = record
            Set bestOptions = available
        End If
    Next record
    Set nextRemaining = New Collection
    For Each record In remaining
        If record("name") <> best("name") Then
            nextRemaining.Add record
        End If
    Next record
    For Each candidate In bestOptions
        chosen.Add candidate
        If SolvePlanarCandidates(nextRemaining, chosen, steps) Then
            SolvePlanarCandidates = True
            Exit Function
        End If
        chosen.Remove chosen.Count
    Next candidate
End Function

Private Sub ImprovePlanarAssignment(ByVal records As Collection, ByVal chosen As Collection)
    Dim pass As Long, index As Long, otherIndex As Long, record As Object, current As Object, candidate As Object, options As Collection, compatible As Boolean
    For pass = 1 To PLANAR_LAYOUT_IMPROVEMENT_PASSES
        For index = 1 To chosen.Count
            Set current = chosen(index)
            For Each record In records
                If record("name") = current("name") Then
                    Set options = record("candidates")
                    Exit For
                End If
            Next record
            For Each candidate In options
                If CDbl(candidate("cost")) >= CDbl(current("cost")) Then
                    Exit For
                End If
                compatible = True
                For otherIndex = 1 To chosen.Count
                    If otherIndex <> index Then
                        If Not PlanarCandidatesCompatibleCached(candidate, chosen(otherIndex)) Then
                            compatible = False
                            Exit For
                        End If
                    End If
                Next otherIndex
                If compatible Then
                    chosen.Remove index
                    If index <= chosen.Count Then
                        chosen.Add candidate, , index
                    Else
                        chosen.Add candidate
                    End If
                    Exit For
                End If
            Next candidate
        Next index
    Next pass
End Sub

Private Sub ReservePlanarViewNotes(ByVal d As SldWorks.ModelDoc2, ByVal view As SldWorks.View, ByVal region As Variant, ByVal obstacles As Collection)
    Dim notes As Variant, item As Variant, note As SldWorks.Note, annotation As SldWorks.Annotation
    Dim format As SldWorks.TextFormat, index As Long, extent As Variant, position As Variant
    Dim labelX As Double, labelTop As Double, outline As Variant
    notes = view.GetNotes
    If IsEmpty(notes) Then
        Exit Sub
    End If
    For Each item In notes
        Set note = item
        If Not planarDetailLabels Is Nothing Then
            If planarDetailLabels.Exists(view.Name) Then
                note.SetText CStr(planarDetailLabels(view.Name)) & " (" & ScaleText(view.ScaleDecimal) & ")"
            End If
        End If
        Set annotation = note.GetAnnotation
        For index = 0 To annotation.GetTextFormatCount - 1
            Set format = annotation.GetTextFormat(index)
            format.TypeFaceName = Setting("Font", DEFAULT_FONT)
            format.CharHeight = TEXT_HEIGHT_STANDARD_M
            format.Italic = False
            annotation.SetTextFormat index, False, format
        Next index
        d.GraphicsRedraw2
        extent = note.GetExtent
        If Not IsEmpty(extent) Then
            position = annotation.GetPosition
            ' A detail label stands over its own image, clear of the first dimension row; other notes keep the region corner.
            labelX = region(0) + LEADER_LABEL_CLEARANCE_M
            labelTop = region(3) - LEADER_LABEL_CLEARANCE_M
            If Not planarDetailLabels Is Nothing Then
                If planarDetailLabels.Exists(view.Name) Then
                    outline = view.GetOutline
                    labelX = MaxD(labelX, outline(OUTLINE_LEFT_INDEX))
                    labelTop = MinD(labelTop, outline(OUTLINE_TOP_INDEX) + PLANAR_DETAIL_TITLE_OFFSET_M + extent(BBOX_Y_MAX_INDEX) - extent(BBOX_Y_MIN_INDEX))
                End If
            End If
            annotation.SetPosition2 position(0) + labelX - extent(BBOX_X_MIN_INDEX), position(1) + labelTop - extent(BBOX_Y_MAX_INDEX), position(2)
            d.GraphicsRedraw2
            extent = note.GetExtent
            obstacles.Add Array(extent(BBOX_X_MIN_INDEX), extent(BBOX_Y_MIN_INDEX), extent(BBOX_X_MAX_INDEX), extent(BBOX_Y_MAX_INDEX))
        End If
    Next item
End Sub

Private Sub FinishPlanarDimensionLayout(ByVal d As SldWorks.ModelDoc2, ByVal views As Collection, ByVal regions As Object, ByVal work As Variant)
    Dim view As SldWorks.View, display As SldWorks.DisplayDimension, records As Collection, chosen As Collection, obstacles As Collection
    Dim record As Object, candidate As Object, candidates As Collection, region As Variant, position As Variant, steps As Long, count As Long
    Dim displays As Object, initialValues As Object, item As Variant, other As Object, rebuilt As Collection
    Dim audit As Object, number As Long, polygon As Variant, corner As Variant, line As Variant, rawPolygon() As Variant, index As Long, rejection As Variant
    Dim actualAngle As Double, initialAngle As Double, started As Double, viewStarted As Double, labelBox As Variant
    Set displays = CreateObject("Scripting.Dictionary")
    Set initialValues = CreateObject("Scripting.Dictionary")
    Set planarLayoutMaterial = CreateObject("Scripting.Dictionary")
    Set planarLayoutVisibleMaterial = CreateObject("Scripting.Dictionary")
    Set planarLayoutOutlines = CreateObject("Scripting.Dictionary")
    If Setting("DimensionLayoutAudit", "0") = "1" Then
        Set audit = fs.CreateTextFile(fs.BuildPath(runDir, PLANAR_LAYOUT_AUDIT_FILENAME), True, True)
        audit.Close
    End If
    ' Native alignment links would move other dimensions while candidates are sampled.
    d.ClearSelection2 True
    For Each item In views
        Set view = item
        Set display = view.GetFirstDisplayDimension5
        Do While Not display Is Nothing
            display.GetAnnotation.Select3 True, Nothing
            Set display = display.GetNext5
        Loop
    Next item
    d.BreakDimensionAlignment
    d.ClearSelection2 True
    For Each item In views
        Set view = item
        ReDim rawPolygon(planarContourPoints.Count - 1)
        For index = 1 To planarContourPoints.Count
            rawPolygon(index - 1) = PatternSheetPoint(view, planarContourPoints(index))
        Next index
        planarLayoutMaterial.Add view.Name, rawPolygon
        planarLayoutOutlines.Add view.Name, view.GetOutline
        planarLayoutVisibleMaterial.Add view.Name, PlanarClipPolygon(rawPolygon, view.GetOutline)
        planarCandidateBudget = PLANAR_LAYOUT_CANDIDATE_LIMIT
        Set planarCandidateResume = CreateObject("Scripting.Dictionary")
        planarModelCandidates = Setting("PlanarModelCandidates", DEFAULT_PLANAR_MODEL_CANDIDATES) = "1"
        planarModelDerived = 0
        viewStarted = Timer
BuildPlanarViewCandidates:
        Set records = New Collection
        Set chosen = New Collection
        Set obstacles = New Collection
        obstacles.Add Array(work(0), sheetHeight - LEADER_STAMP_TOP_MARGIN_M, LEADER_STAMP_WIDTH_M, sheetHeight)
        region = regions(view.Name)
        ReservePlanarViewNotes d, view, region, obstacles
        If Not planarDetailLabelBoxes Is Nothing Then
            If planarDetailLabelBoxes.Exists(view.Name) Then
                For Each labelBox In planarDetailLabelBoxes(view.Name)
                    obstacles.Add labelBox
                Next labelBox
            End If
        End If
        Set display = view.GetFirstDisplayDimension5
        Do While Not display Is Nothing
            Set planarLayoutRejections = CreateObject("Scripting.Dictionary")
            If Not initialValues.Exists(display.GetNameForSelection) Then
                initialValues.Add display.GetNameForSelection, display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue
            End If
            planarSampleCount = 0
            planarSampleSkips = 0
            started = Timer
            Set candidates = PlanarDimensionCandidates(d, display, view, region, obstacles)
            number = 0
            For Each candidate In candidates
                number = number + 1
                ' A resumed list already has IDs; renumber it for this solve.
                candidate("id") = display.GetNameForSelection & "|" & CStr(number)
            Next candidate
            Log "PLANAR CANDIDATES " & display.GetNameForSelection & "; count=" & candidates.Count & "; native samples=" & planarSampleCount & "; skipped points=" & planarSampleSkips & "; seconds=" & Format$(PlanarElapsedSeconds(started), "0.0")
            If candidates.Count = 0 Or Setting("DimensionLayoutAudit", "0") = "1" Then
                For Each rejection In planarLayoutRejections.Keys
                    Log "PLANAR REJECT " & display.GetNameForSelection & "; " & rejection & "=" & planarLayoutRejections(rejection)
                Next rejection
            End If
            If Setting("DimensionLayoutAudit", "0") = "1" Then
                Set audit = fs.OpenTextFile(fs.BuildPath(runDir, PLANAR_LAYOUT_AUDIT_FILENAME), FILE_FOR_APPENDING, True, FILE_UNICODE)
                number = 0
                For Each candidate In candidates
                    number = number + 1
                    audit.WriteLine display.GetNameForSelection & vbTab & number & vbTab & "position" & vbTab & PlanarAuditCoordinates(candidate("position"))
                    audit.WriteLine display.GetNameForSelection & vbTab & number & vbTab & "view" & vbTab & view.Name
                    audit.WriteLine display.GetNameForSelection & vbTab & number & vbTab & "arrowSide" & vbTab & candidate("arrowSide")
                    For Each polygon In candidate("text")
                        audit.WriteLine display.GetNameForSelection & vbTab & number & vbTab & "text" & vbTab & PlanarAuditCoordinates(polygon(0)) & vbTab & PlanarAuditCoordinates(polygon(1)) & vbTab & PlanarAuditCoordinates(polygon(2)) & vbTab & PlanarAuditCoordinates(polygon(3))
                    Next polygon
                    For Each line In candidate("lines")
                        audit.WriteLine display.GetNameForSelection & vbTab & number & vbTab & "line" & vbTab & PlanarAuditCoordinates(line)
                    Next line
                    For Each line In candidate("bases")
                        audit.WriteLine display.GetNameForSelection & vbTab & number & vbTab & "base" & vbTab & PlanarAuditCoordinates(line)
                    Next line
                Next candidate
                audit.Close
            End If
            If candidates.Count = 0 Then
                If planarModelCandidates Then
                    GoTo PlanarModelMismatch
                End If
                RequestPlanarSplit view.Name
                Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "No valid native planar placement: " & display.GetNameForSelection
            End If
            Set record = CreateObject("Scripting.Dictionary")
            record.Add "name", display.GetNameForSelection
            record.Add "candidates", candidates
            records.Add record
            If Not displays.Exists(display.GetNameForSelection) Then
                displays.Add display.GetNameForSelection, display
            End If
            Set display = display.GetNext5
        Loop
        steps = 0
        planarStrictCrossings = True
        planarSearchLimit = PLANAR_STRICT_SEARCH_LIMIT
        Set planarLayoutCompatibility = CreateObject("Scripting.Dictionary")
        If Not SolvePlanarCandidates(records, chosen, steps) Then
            If planarCandidateBudget > 0 Then
                Log "PLANAR bounded candidate pass needs more alternatives; expanding this view only from the next unsampled rows: " & view.Name
                planarCandidateBudget = 0
                GoTo BuildPlanarViewCandidates
            End If
            ' DSTU GOST 2.307 asks to avoid these crossings. When no assignment
            ' avoids them, allow a dimension line to cross another dimension's
            ' extension line; every text, arrowhead and baseline check remains.
            Log "PLANAR no assignment avoids extension-line crossings; search=" & steps & "; allowing them in this view only: " & view.Name
            RequestPlanarSplit view.Name
            Set chosen = New Collection
            steps = 0
            planarStrictCrossings = False
            planarSearchLimit = PLANAR_LAYOUT_SEARCH_LIMIT
            Set planarLayoutCompatibility = CreateObject("Scripting.Dictionary")
            If Not SolvePlanarCandidates(records, chosen, steps) Then
                Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "No non-overlapping planar dimension assignment: " & view.Name & "; search=" & steps
            End If
        End If
        ImprovePlanarAssignment records, chosen
        Set rebuilt = New Collection
        For Each candidate In chosen
            Set display = displays(candidate("name"))
            display.ArrowSide = CLng(candidate("arrowSide"))
            If Setting("DimensionLayoutAudit", "0") = "1" Then
                Log "PLANAR RESTORE " & candidate("name") & "; value=" & candidate("value") & "; center=" & candidate("centerText") & "; offset=" & candidate("offsetText") & "; opposite=" & candidate("oppositeAngle")
            End If
            If display.Type2 = SW_DIMENSION_ANGULAR Then
                SetPlanarAngleRepresentation d, display, CDbl(candidate("value"))
                ' Native sector switches also mirror the text. Select the sector before placing it.
                RestorePlanarAngularSector d, display, candidate
            End If
            If CBool(candidate("offsetText")) Then
                position = candidate("baseControlPoint")
                display.OffsetText = False
                display.CenterText = True
                display.SetBrokenLeader2 False, SW_SOLID_LEADER_ALIGNED_TEXT
                display.GetAnnotation.SetPosition2 position(0), position(1), position(2)
                d.GraphicsRedraw2
            End If
            display.OffsetText = candidate("offsetText")
            display.CenterText = candidate("centerText")
            If CBool(candidate("offsetText")) Or display.Type2 = SW_DIMENSION_DIAMETER Then
                display.SetBrokenLeader2 False, SW_BROKEN_LEADER_HORIZONTAL_TEXT
            End If
            position = candidate("controlPoint")
            AsDrawing(d).ActivateView ""
            display.GetAnnotation.SetPosition2 position(0), position(1), position(2)
            display.OffsetText = candidate("offsetText")
            display.CenterText = candidate("centerText")
            display.ArrowSide = CLng(candidate("arrowSide"))
            d.GraphicsRedraw2
            DoEvents
            If display.Type2 = SW_DIMENSION_ANGULAR Then
                actualAngle = display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue
                initialAngle = CDbl(initialValues(candidate("name")))
                If actualAngle < 0 Or actualAngle > PI Or Abs(MinD(actualAngle, PI - actualAngle) - MinD(initialAngle, PI - initialAngle)) > PLANAR_ANGLE_TOLERANCE_RAD Then
                    Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "The planar angle references changed during layout."
                End If
            ElseIf Abs(display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue - CDbl(initialValues(candidate("name")))) > LEADER_VALUE_TOLERANCE_M Then
                Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "The planar dimension value changed during layout."
            End If
            Set planarLayoutRejections = CreateObject("Scripting.Dictionary")
            AsDrawing(d).ActivateView ""
            Set record = PlanarCaptureCandidate(display, view, region, obstacles)
            If record Is Nothing Then
                Log "PLANAR EXPECTED BOUNDS " & Join(candidate("bounds"), ",") & "; actual=" & Join(NativeDimensionBounds(display), ",")
                Log "PLANAR CONTROL " & PlanarAuditCoordinates(candidate("controlPoint")) & "; expected position=" & PlanarAuditCoordinates(candidate("position")) & "; actual position=" & PlanarAuditCoordinates(display.GetAnnotation.GetPosition)
                If planarModelCandidates Then
                    GoTo PlanarModelMismatch
                End If
                Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "The restored native placement differs from its candidate: " & candidate("name") & "; " & Join(planarLayoutRejections.Keys, ", ")
            End If
            For Each other In rebuilt
                If Not PlanarCandidatesCompatible(record, other) Then
                    If planarModelCandidates Then
                        GoTo PlanarModelMismatch
                    End If
                    Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "The restored native dimensions overlap."
                End If
            Next other
            rebuilt.Add record
            Log "PLANAR ARROWS " & candidate("name") & "; side=" & display.ArrowSide
            ValidateNativeViewIsolation display, view.Name, views
            count = count + 1
        Next candidate
        Log "PLANAR ASSIGNMENT " & view.Name & "; dimensions=" & chosen.Count & "; search=" & steps & "; modelled candidates=" & planarModelDerived & "; native check=" & planarModelCandidates & "; seconds=" & Format$(PlanarElapsedSeconds(viewStarted), "0.0")
        GoTo NextPlanarLayoutView
PlanarModelMismatch:
        ' SolidWorks did not draw a modelled placement as predicted: sample this view natively.
        Log "PLANAR MODEL mismatch in " & view.Name & "; sampling this view natively"
        planarModelCandidates = False
        planarCandidateBudget = PLANAR_LAYOUT_CANDIDATE_LIMIT
        Set planarCandidateResume = CreateObject("Scripting.Dictionary")
        GoTo BuildPlanarViewCandidates
NextPlanarLayoutView:
    Next item
    Set planarCandidateResume = Nothing
    Log "NATIVE LEADERS dimensions=" & count & "; complete planar graphics and text checked together"
End Sub
