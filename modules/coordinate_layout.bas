Attribute VB_Name = "coordinate_layout"
Option Explicit
Private Function CompatibleRequiredCoordinateLanes(ByVal plans As Collection, ByVal side As Long, ByVal axisMinimum As Double, ByVal axisMaximum As Double) As Long
    Dim plan As Variant, count As Long, total As Double, lanes As Long, lane As Long
    Dim index As Long, used As Double, fits As Boolean
    If axisMaximum - axisMinimum < MIN_PROJECTED_EXTENT_M Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "No available axis range for graphical dimension text."
    End If
    For Each plan In plans
        If CLng(plan("side")) = side Then
            count = count + 1
            total = total + CDbl(plan("width")) + COORDINATE_LABEL_GAP_M
            If CDbl(plan("width")) > axisMaximum - axisMinimum Then
                Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "A graphical dimension label is wider than its available band."
            End If
        End If
    Next plan
    If count = 0 Then
        Exit Function
    End If
    lanes = -Int(-total / (axisMaximum - axisMinimum + COORDINATE_LABEL_GAP_M))
    Do
        fits = True
        For lane = 0 To lanes - 1
            index = 0
            used = -COORDINATE_LABEL_GAP_M
            For Each plan In plans
                If CLng(plan("side")) = side Then
                    If index Mod lanes = lane Then
                        used = used + CDbl(plan("width")) + COORDINATE_LABEL_GAP_M
                    End If
                    index = index + 1
                End If
            Next plan
            If used > axisMaximum - axisMinimum Then
                fits = False
            End If
        Next lane
        If fits Then
            CompatibleRequiredCoordinateLanes = lanes
            Exit Function
        End If
        lanes = lanes + 1
    Loop
End Function
Private Sub CompatibleBalanceCoordinateSides(ByVal plans As Collection, ByVal axisMinimum As Double, ByVal axisMaximum As Double, ByVal lowLanes As Long, ByVal highLanes As Long)
    Dim lowRequired As Long, highRequired As Long, sourceSide As Long, targetSide As Long
    Dim targetCapacity As Long, plan As Variant, best As Object, cost As Double, bestCost As Double
    Do
        lowRequired = CompatibleRequiredCoordinateLanes(plans, 0, axisMinimum, axisMaximum)
        highRequired = CompatibleRequiredCoordinateLanes(plans, 1, axisMinimum, axisMaximum)
        If lowRequired <= lowLanes And highRequired <= highLanes Then
            Exit Sub
        End If
        If lowRequired > lowLanes And highRequired > highLanes Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "The complete graphical dimension set does not fit the selected sheet."
        End If
        sourceSide = 0
        targetCapacity = highLanes
        If highRequired > highLanes Then
            sourceSide = 1
            targetCapacity = lowLanes
        End If
        targetSide = 1 - sourceSide
        Set best = Nothing
        bestCost = SEARCH_INITIAL_LIMIT
        For Each plan In plans
            If CLng(plan("side")) = sourceSide Then
                plan("side") = targetSide
                If CompatibleRequiredCoordinateLanes(plans, targetSide, axisMinimum, axisMaximum) <= targetCapacity Then
                    If targetSide = 0 Then
                        cost = CDbl(plan("lowCost")) - CDbl(plan("highCost"))
                    Else
                        cost = CDbl(plan("highCost")) - CDbl(plan("lowCost"))
                    End If
                    If cost < bestCost Then
                        Set best = plan
                        bestCost = cost
                    End If
                End If
                plan("side") = sourceSide
            End If
        Next plan
        If best Is Nothing Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Graphical dimensions cannot be moved to the opposite free band without overlaps."
        End If
        best("side") = targetSide
    Loop
End Sub
Private Sub CompatiblePackCoordinateSide(ByVal plans As Collection, ByVal side As Long, ByVal axisMinimum As Double, ByVal axisMaximum As Double, ByVal firstLane As Double, ByVal maxLanes As Long, ByVal axis As Long)
    Dim plan As Variant, count As Long, total As Double, lanes As Long, lane As Long
    Dim index As Long, used As Double, fits As Boolean, previous As Double, position As Double, i As Long
    For Each plan In plans
        If CLng(plan("side")) = side Then
            count = count + 1
            total = total + CDbl(plan("width")) + COORDINATE_LABEL_GAP_M
        End If
    Next plan
    If count = 0 Then
        Exit Sub
    End If
    If maxLanes < 1 Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "No free band for graphical dimensions."
    End If
    If maxLanes > COORDINATE_LAST_LANE + 1 Then
        maxLanes = COORDINATE_LAST_LANE + 1
    End If
    lanes = -Int(-total / (axisMaximum - axisMinimum + COORDINATE_LABEL_GAP_M))
    If lanes < 1 Then
        lanes = 1
    End If
    Do
        If lanes > maxLanes Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Graphical dimensions need more space on the selected sheet; choose a larger paper format."
        End If
        fits = True
        For lane = 0 To lanes - 1
            index = 0
            used = -COORDINATE_LABEL_GAP_M
            For Each plan In plans
                If CLng(plan("side")) = side Then
                    If index Mod lanes = lane Then
                        used = used + CDbl(plan("width")) + COORDINATE_LABEL_GAP_M
                    End If
                    index = index + 1
                End If
            Next plan
            If used > axisMaximum - axisMinimum Then
                fits = False
            End If
        Next lane
        If fits Then
            Exit Do
        End If
        lanes = lanes + 1
    Loop
    If axis = 1 Then
        If side = 0 Then
            firstLane = CoordinateRegionEdge(0) + (lanes - 1) * COORDINATE_LANE_PITCH_M
        Else
            firstLane = CoordinateRegionEdge(2) - (lanes - 1) * COORDINATE_LANE_PITCH_M
        End If
    End If
    For lane = 0 To lanes - 1
        index = 0
        previous = axisMinimum - COORDINATE_LABEL_GAP_M
        For Each plan In plans
            If CLng(plan("side")) = side Then
                If index Mod lanes = lane Then
                    position = -MinD(-CDbl(plan("desired")), -(previous + COORDINATE_LABEL_GAP_M + CDbl(plan("width")) / 2))
                    plan("center") = position
                    plan("lane") = lane
                    plan("location") = firstLane + (side * 2 - 1) * lane * COORDINATE_LANE_PITCH_M
                    previous = position + CDbl(plan("width")) / 2
                End If
                index = index + 1
            End If
        Next plan
        previous = axisMaximum + COORDINATE_LABEL_GAP_M
        For i = plans.Count To 1 Step -1
            Set plan = plans(i)
            If CLng(plan("side")) = side And CLng(plan("lane")) = lane Then
                position = MinD(CDbl(plan("center")), previous - COORDINATE_LABEL_GAP_M - CDbl(plan("width")) / 2)
                plan("center") = position
                previous = position - CDbl(plan("width")) / 2
                If previous < axisMinimum - MIN_PROJECTED_EXTENT_M Then
                    Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Graphical dimension text exceeds the available band."
                End If
            End If
        Next i
    Next lane
End Sub

Private Sub AddForbiddenCoordinate(ByVal coordinates As Object, ByVal value As Double)
    Dim record As Object, key As String
    key = CStr(value)
    If Not coordinates.Exists(key) Then
        Set record = CreateObject("Scripting.Dictionary")
        record.Add "coordinate", value
        coordinates.Add key, record
    End If
End Sub

Private Function TryCoordinatePlanRow(ByVal plans As Collection, ByVal plan As Object, ByVal side As Long, ByVal lane As Long, ByVal axisMinimum As Double, ByVal axisMaximum As Double, ByRef center As Double, ByRef start As Double, ByRef finish As Double) As Boolean
    Dim other As Variant, forbidden As Object, minimum As Double, maximum As Double
    Dim first As Double, last As Double, padding As Double, innerStart As Double, innerEnd As Double, keys As Variant
    first = CDbl(plan("firstCoordinate"))
    last = CDbl(plan("lastCoordinate"))
    padding = CoordinateMarkerPadding(CDbl(plan("span")), CDbl(plan("width")))
    start = first - padding
    finish = last + padding
    minimum = axisMinimum
    maximum = axisMaximum
    If start < minimum Or finish > maximum Then
        Exit Function
    End If
    Set forbidden = CreateObject("Scripting.Dictionary")
    AddForbiddenCoordinate forbidden, first
    AddForbiddenCoordinate forbidden, last
    For Each other In plans
        If CLng(other("lane")) >= 0 Then
            If CLng(other("side")) = side Then
                If CLng(other("lane")) = lane Then
                    If CDbl(other("lineEnd")) < start Then
                        minimum = -MinD(-minimum, -CDbl(other("lineEnd")) - DIMENSION_GEOMETRY_GAP_M)
                    ElseIf CDbl(other("lineStart")) > finish Then
                        maximum = MinD(maximum, CDbl(other("lineStart")) - DIMENSION_GEOMETRY_GAP_M)
                    Else
                        Exit Function
                    End If
                End If
                If CLng(other("lane")) >= lane Then
                    AddForbiddenCoordinate forbidden, CDbl(other("firstCoordinate"))
                    AddForbiddenCoordinate forbidden, CDbl(other("lastCoordinate"))
                Else
                    ' New outer extension lines must not cross an existing inner number.
                    innerStart = CDbl(other("desired")) - CDbl(other("width")) / 2 - DIMENSION_GEOMETRY_GAP_M
                    innerEnd = CDbl(other("desired")) + CDbl(other("width")) / 2 + DIMENSION_GEOMETRY_GAP_M
                    If (first >= innerStart And first <= innerEnd) Or (last >= innerStart And last <= innerEnd) Then
                        Exit Function
                    End If
                End If
            End If
        End If
    Next other
    If start < minimum Or finish > maximum Or maximum - minimum < CDbl(plan("width")) Then
        Exit Function
    End If
    keys = SortedCoordinateKeys(forbidden)
    If Not FindCoordinateTextCenter(forbidden, keys, first, last, CDbl(plan("width")), minimum, maximum, center) Then
        Exit Function
    End If
    start = MinD(start, center - CDbl(plan("width")) / 2 - COORDINATE_LABEL_GAP_M)
    finish = -MinD(-finish, -center - CDbl(plan("width")) / 2 - COORDINATE_LABEL_GAP_M)
    TryCoordinatePlanRow = start >= axisMinimum And finish <= axisMaximum
End Function

Private Sub BalanceCoordinateSides(ByVal plans As Collection, ByVal axisMinimum As Double, ByVal axisMaximum As Double, ByVal lowLanes As Long, ByVal highLanes As Long)
    Dim plan As Variant, row As Long, side As Long, capacity As Long, chosen As Long, chosenSide As Long
    Dim center As Double, start As Double, finish As Double, best As Double, score As Double
    Dim chosenCenter As Double, chosenStart As Double, chosenEnd As Double, minimumLane As Long
    For Each plan In plans
        plan("lane") = -1
    Next plan
    For Each plan In plans
        chosen = -1
        best = SEARCH_INITIAL_LIMIT
        For side = 0 To 1
            capacity = lowLanes
            If side = 1 Then
                capacity = highLanes
            End If
            For row = capacity - 1 To 0 Step -1
                If TryCoordinatePlanRow(plans, plan, side, row, axisMinimum, axisMaximum, center, start, finish) Then
                    score = CDbl(plan("lowCost"))
                    If side = 1 Then
                        score = CDbl(plan("highCost"))
                    End If
                    score = score + (capacity - 1 - row) * COORDINATE_ROW_PRIORITY_M + Abs(center - CDbl(plan("geometricCenter"))) * COORDINATE_LANE_PENALTY_M
                    If score < best Then
                        chosen = row
                        chosenSide = side
                        chosenCenter = center
                        chosenStart = start
                        chosenEnd = finish
                        best = score
                    End If
                End If
            Next row
        Next side
        If chosen < 0 Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Dimension arrows and numbers cannot fit without crossing extension lines at 7 mm row spacing."
        End If
        plan("side") = chosenSide
        plan("lane") = chosen
        plan("desired") = chosenCenter
        plan("lineStart") = chosenStart
        plan("lineEnd") = chosenEnd
    Next plan
    For side = 0 To 1
        minimumLane = COORDINATE_LAST_LANE + 1
        For Each plan In plans
            If CLng(plan("side")) = side Then
                If CLng(plan("lane")) < minimumLane Then
                    minimumLane = CLng(plan("lane"))
                End If
            End If
        Next plan
        For Each plan In plans
            If CLng(plan("side")) = side Then
                plan("lane") = CLng(plan("lane")) - minimumLane
            End If
        Next plan
    Next side
End Sub

Private Sub PackCoordinateSide(ByVal plans As Collection, ByVal side As Long, ByVal axisMinimum As Double, ByVal axisMaximum As Double, ByVal firstLane As Double, ByVal maxLanes As Long, ByVal axis As Long)
    Dim plan As Variant, count As Long, lanes As Long, lane As Long
    For Each plan In plans
        If CLng(plan("side")) = side Then
            count = count + 1
            If CLng(plan("lane")) + 1 > lanes Then
                lanes = CLng(plan("lane")) + 1
            End If
        End If
    Next plan
    If count = 0 Then
        Exit Sub
    End If
    If lanes > maxLanes Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Dimension row assignment exceeds its free band."
    End If
    If axis = 1 Then
        If side = 0 Then
            firstLane = CoordinateRegionEdge(0) + (lanes - 1) * COORDINATE_LANE_PITCH_M
        Else
            firstLane = CoordinateRegionEdge(2) - (lanes - 1) * COORDINATE_LANE_PITCH_M
        End If
    End If
    For Each plan In plans
        If CLng(plan("side")) = side Then
            lane = CLng(plan("lane"))
            plan("center") = CDbl(plan("desired"))
            plan("location") = firstLane + (side * 2 - 1) * lane * COORDINATE_LANE_PITCH_M
        End If
    Next plan
End Sub

Private Sub ResetCoordinatePlans(ByVal plans As Collection)
    Dim plan As Variant
    For Each plan In plans
        plan("lane") = -1
        plan("desired") = CDbl(plan("geometricCenter"))
        plan("side") = CLng(plan("preferredSide"))
    Next plan
End Sub

Private Function CoordinateAxis(ByVal d As SldWorks.ModelDoc2, ByVal v As SldWorks.View, ByVal coordinates As Object, ByVal axis As Long, ByVal bounds As Variant, ByVal detail As Boolean, Optional ByVal layoutOnly As Boolean = False) As Long
    Dim keys As Variant, i As Long, runEnd As Long, first As Object, last As Object
    Dim plans As Collection, plan As Variant, pitch As Double, nextPitch As Double, span As Double
    Dim lowEdge As Double, highEdge As Double, lowDistance As Double, highDistance As Double
    Dim lowCount As Long, highCount As Long, side As Long, lowLane As Double, highLane As Double
    Dim lowLanes As Long, highLanes As Long, axisMinimum As Double, axisMaximum As Double
    Dim text As String, prefixText As String, overallKey As String, intervalKey As String
    Dim dd As SldWorks.DisplayDimension, dimension As SldWorks.Dimension, annotation As SldWorks.Annotation
    Dim layoutError As Long, layoutMessage As String
    Dim covered As Long, expected As Long, repeated As Long, grouped As Long, padding As Double, textWidth As Double, center As Double
    If coordinates.Count < COORDINATE_MIN_KEY_COUNT And coordinateAssignedRequests Is Nothing Then
        Exit Function
    End If
    Dim requests As Collection, request As Variant
    Set plans = New Collection
    If coordinateAssignedRequests Is Nothing Then
        Set requests = AdjacentCoordinateRequests(v, coordinates, axis, detail)
    Else
        Set requests = AssignedRequestsForView(v, axis)
        Set coordinates = RequestCoordinates(requests, v.ScaleDecimal)
    End If
    If coordinates.Count < 2 Then
        Exit Function
    End If
    keys = SortedCoordinateKeys(coordinates)
    If axis = 0 Then
        lowEdge = bounds(OUTLINE_BOTTOM_INDEX)
        highEdge = bounds(OUTLINE_TOP_INDEX)
        axisMinimum = routingXMinimum
        axisMaximum = routingXMaximum
    Else
        lowEdge = bounds(OUTLINE_LEFT_INDEX)
        highEdge = bounds(OUTLINE_RIGHT_INDEX)
        axisMinimum = CoordinateRegionEdge(1)
        axisMaximum = CoordinateRegionEdge(3)
    End If
    lowLane = lowEdge - COORDINATE_BASE_OFFSET_M
    highLane = highEdge + COORDINATE_BASE_OFFSET_M
    If axis = 0 Then
        lowLanes = Int((lowLane - CoordinateRegionEdge(1)) / COORDINATE_LANE_PITCH_M) + 1
        highLanes = Int((CoordinateRegionEdge(3) - highLane) / COORDINATE_LANE_PITCH_M) + 1
    Else
        lowLanes = Int((lowLane - CoordinateRegionEdge(0)) / COORDINATE_LANE_PITCH_M) + 1
        highLanes = Int((CoordinateRegionEdge(2) - highLane) / COORDINATE_LANE_PITCH_M) + 1
    End If
    If lowLanes < 0 Then
        lowLanes = 0
    End If
    If highLanes < 0 Then
        highLanes = 0
    End If
    If lowLanes > COORDINATE_LAST_LANE + 1 Then
        lowLanes = COORDINATE_LAST_LANE + 1
    End If
    If highLanes > COORDINATE_LAST_LANE + 1 Then
        highLanes = COORDINATE_LAST_LANE + 1
    End If
    For Each request In requests
        Set first = request("firstRecord")
        Set last = request("lastRecord")
        span = CDbl(last("coordinate")) - CDbl(first("coordinate"))
        pitch = CDbl(request("pitch"))
        intervalKey = CStr(request("key"))
        prefixText = ""
        If Not coordinateSelection Is Nothing Then
            If Not coordinateSelection.Exists(intervalKey) Then
                GoTo NextRun
            End If
        End If
        expected = expected + CLng(request("count"))
        If dimensionKeys.Exists(intervalKey) Then
            covered = covered + CLng(request("count"))
            repeated = repeated + 1
            GoTo NextRun
        End If
        If CLng(request("count")) > 1 Then
            prefixText = CStr(request("count")) & "x " & DimensionLabel(pitch / v.ScaleDecimal * MILLIMETRES_PER_METRE) & " = "
            overallKey = "overall|" & CStr(detail) & "|" & PhysicalAxisKey(v, axis) & "|" & Format$(span / v.ScaleDecimal, GEOMETRY_KEY_FORMAT)
            If CBool(request("fullRange")) And dimensionKeys.Exists(overallKey) Then
                Set dd = dimensionKeys(overallKey)
                dd.SetText SW_DIMENSION_TEXT_PREFIX, prefixText & dd.GetText(SW_DIMENSION_TEXT_PREFIX)
                dimensionKeys.Add intervalKey, True
                covered = covered + CLng(request("count"))
                grouped = grouped + 1
                GoTo NextRun
            End If
        End If
        lowDistance = CDbl(first("low")) + CDbl(last("low")) - lowEdge * 2 + lowCount * COORDINATE_SIDE_BALANCE_M
        highDistance = highEdge * 2 - CDbl(first("high")) - CDbl(last("high")) + highCount * COORDINATE_SIDE_BALANCE_M
        side = 0
        If highDistance < lowDistance Then
            side = 1
        End If
        Set plan = CreateObject("Scripting.Dictionary")
        plan.Add "side", side
        plan.Add "preferredSide", side
        plan.Add "lowCost", CDbl(first("low")) + CDbl(last("low")) - lowEdge * 2
        plan.Add "highCost", highEdge * 2 - CDbl(first("high")) - CDbl(last("high"))
        plan.Add "firstRecord", first
        plan.Add "lastRecord", last
        If side = 0 Then
            plan.Add "first", first("lowEntity")
            plan.Add "last", last("lowEntity")
            lowCount = lowCount + 1
        Else
            plan.Add "first", first("highEntity")
            plan.Add "last", last("highEntity")
            highCount = highCount + 1
        End If
        text = prefixText & DimensionLabel(span / v.ScaleDecimal * MILLIMETRES_PER_METRE)
        textWidth = Len(text) * COORDINATE_GLYPH_WIDTH_M + COORDINATE_LABEL_PADDING_M
        center = (CDbl(first("coordinate")) + CDbl(last("coordinate"))) / 2
        padding = CoordinateMarkerPadding(span, textWidth)
        plan.Add "width", textWidth
        plan.Add "desired", center
        plan.Add "geometricCenter", center
        plan.Add "firstCoordinate", CDbl(first("coordinate"))
        plan.Add "lastCoordinate", CDbl(last("coordinate"))
        plan.Add "lineStart", MinD(CDbl(first("coordinate")) - padding, center - textWidth / 2 - COORDINATE_LABEL_GAP_M)
        plan.Add "lineEnd", -MinD(-CDbl(last("coordinate")) - padding, -center - textWidth / 2 - COORDINATE_LABEL_GAP_M)
        plan.Add "span", span
        plan.Add "prefix", prefixText
        plan.Add "key", intervalKey
        plan.Add "count", CLng(request("count"))
        plan.Add "lane", -1
        If Not coordinateSourceView Is Nothing Then
            SetCompactPlanReferences plan
        End If
        plans.Add plan
NextRun:
    Next request
    If Not coordinateAssignedRequests Is Nothing Then
        Set plans = SortCoordinatePlans(plans)
    End If
    On Error Resume Next
    BalanceCoordinateSides plans, axisMinimum, axisMaximum, lowLanes, highLanes
    layoutError = Err.Number
    layoutMessage = Err.Description
    Err.Clear
    On Error GoTo 0
    If layoutError = 0 Then
        PackCoordinateSide plans, 0, axisMinimum, axisMaximum, lowLane, lowLanes, axis
        PackCoordinateSide plans, 1, axisMinimum, axisMaximum, highLane, highLanes, axis
    ElseIf Setting("StrictDimensionLayout", "0") = "1" Then
        Err.Raise layoutError, "CoordinateAxis", layoutMessage
    Else
        ResetCoordinatePlans plans
        CompatibleBalanceCoordinateSides plans, axisMinimum, axisMaximum, lowLanes, highLanes
        CompatiblePackCoordinateSide plans, 0, axisMinimum, axisMaximum, lowLane, lowLanes, axis
        CompatiblePackCoordinateSide plans, 1, axisMinimum, axisMaximum, highLane, highLanes, axis
        If Not layoutOnly Then
            Warn "LAYOUT REVIEW " & v.Name & " axis=" & axis & ": arrow geometry follows GOST; complete collision-free placement was not established."
        End If
    End If
    If layoutOnly Then
        If axis = 1 Then
            For Each plan In plans
                If CLng(plan("side")) = 0 Then
                    routingXMinimum = -MinD(-routingXMinimum, -(CDbl(plan("location")) + TEXT_HEIGHT_STANDARD_M + COORDINATE_LABEL_GAP_M))
                Else
                    routingXMaximum = MinD(routingXMaximum, CDbl(plan("location")) - TEXT_HEIGHT_STANDARD_M - COORDINATE_LABEL_GAP_M)
                End If
            Next plan
        End If
        Exit Function
    End If
    For Each plan In plans
        AsDrawing(d).ActivateView ""
        d.ClearSelection2 True
        If Not SelectRoutedReference(v, plan("firstRecord"), CLng(plan("side")), False) Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot select the first graphical dimension reference."
        End If
        If Not SelectRoutedReference(v, plan("lastRecord"), CLng(plan("side")), True) Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot select the second graphical dimension reference."
        End If
        If axis = 0 Then
            Set dd = d.AddHorizontalDimension2(CDbl(plan("center")), CDbl(plan("location")), 0)
        Else
            Set dd = d.AddVerticalDimension2(CDbl(plan("location")), CDbl(plan("center")), 0)
        End If
        If dd Is Nothing Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot create a graphical dimension."
        End If
        Set dimension = dd.GetDimension2(SW_DIMENSION_INDEX_FIRST)
        If Not dimension Is Nothing Then
            dimension.SetArcEndCondition 1, SW_ARC_CENTER
            dimension.SetArcEndCondition 2, SW_ARC_CENTER
        End If
        If Not ValidateDimension(d, v, dd, CDbl(plan("span"))) Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Graphical dimension measured an unexpected span."
        End If
        dd.CenterText = False
        If UseCoordinateDots(CDbl(plan("span"))) Then
            dd.SetArrowHeadStyle2 False, SW_ARROW_DOT, SW_ARROW_DOT
        End If
        ApplyGostDimensionStyle dd
        If Not UseCoordinateDots(CDbl(plan("span"))) Then
            dd.ArrowSide = SW_DIMENSION_ARROWS_OUTSIDE
            If CDbl(plan("span")) >= CDbl(plan("width")) + GOST_ARROW_HEAD_LENGTH_M * 2 Then
                dd.ArrowSide = SW_DIMENSION_ARROWS_INSIDE
            End If
        End If
        If Len(CStr(plan("prefix"))) > 0 Then
            dd.SetText SW_DIMENSION_TEXT_PREFIX, CStr(plan("prefix")) & dd.GetText(SW_DIMENSION_TEXT_PREFIX)
        End If
        Set annotation = dd.GetAnnotation
        If axis = 0 Then
            annotation.SetPosition2 CDbl(plan("center")), CDbl(plan("location")), 0
        Else
            annotation.SetPosition2 CDbl(plan("location")), CDbl(plan("center")), 0
        End If
        dimensionKeys.Add CStr(plan("key")), True
        covered = covered + CLng(plan("count"))
        CoordinateAxis = CoordinateAxis + 1
    Next plan
    Log "ROUTING " & v.Name & " axis=" & axis & " intervals=" & expected & " covered=" & covered & " placed=" & CoordinateAxis & " duplicate=" & repeated & " grouped=" & grouped
    d.ClearSelection2 True
End Function


