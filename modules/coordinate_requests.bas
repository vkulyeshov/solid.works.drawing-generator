Attribute VB_Name = "coordinate_requests"
Option Explicit
Private Function CoordinateRegionEdge(ByVal index As Long) As Double
    If Not IsEmpty(coordinateRegion) Then
        CoordinateRegionEdge = coordinateRegion(index)
        Exit Function
    End If
    Select Case index
        Case OUTLINE_LEFT_INDEX
            CoordinateRegionEdge = COORDINATE_LEFT_LIMIT_M
        Case OUTLINE_BOTTOM_INDEX
            CoordinateRegionEdge = COORDINATE_BOTTOM_LIMIT_M
        Case OUTLINE_RIGHT_INDEX
            CoordinateRegionEdge = sheetWidth - COORDINATE_AXIS_RIGHT_MARGIN_M
        Case OUTLINE_TOP_INDEX
            CoordinateRegionEdge = sheetHeight - COORDINATE_AXIS_TOP_MARGIN_M
    End Select
End Function

Private Function ViewOrigin(ByVal v As SldWorks.View) As Variant
    Dim util As SldWorks.MathUtility, point As SldWorks.MathPoint, zero(2) As Double
    Set util = sw.GetMathUtility
    Set point = util.CreatePoint(zero)
    Set point = point.MultiplyTransform(v.ModelToViewTransform)
    ViewOrigin = point.ArrayData
End Function

Private Function DetailCoordinate(ByVal value As Double, ByVal axis As Long, ByVal v As SldWorks.View) As Double
    Dim origin As Variant
    origin = ViewOrigin(v)
    DetailCoordinate = origin(axis) + (value - coordinateSourceOrigin(axis)) * v.ScaleDecimal / coordinateSourceScale
End Function

Private Function TransformCoordinates(ByVal coordinates As Object, ByVal sourceView As SldWorks.View, ByVal v As SldWorks.View, ByVal axis As Long) As Object
    Dim key As Variant, record As Object, result As Object, candidates As Collection, positions As Collection, i As Long
    Dim origin As Variant, ratio As Double, value As Double, perpendicular As Double
    origin = ViewOrigin(v)
    ratio = v.ScaleDecimal / coordinateSourceScale
    Set result = CreateObject("Scripting.Dictionary")
    For Each key In coordinates.Keys
        Set record = coordinates(key)
        Set candidates = record("candidates")
        Set positions = record("candidatePositions")
        For i = 1 To candidates.Count
            value = origin(axis) + (CDbl(record("coordinate")) - coordinateSourceOrigin(axis)) * ratio
            perpendicular = origin(1 - axis) + (CDbl(positions(i)) - coordinateSourceOrigin(1 - axis)) * ratio
            AddCoordinate result, candidates(i), value, v.ScaleDecimal, perpendicular
        Next i
    Next key
    Set TransformCoordinates = result
End Function

Private Function CompactCoordinateRecord(ByVal original As Object, ByVal entity As Object, ByVal perpendicular As Double) As Object
    Dim record As Object, candidates As Collection, positions As Collection
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "coordinate", original("coordinate")
    record.Add "low", perpendicular
    record.Add "high", perpendicular
    record.Add "lowEntity", entity
    record.Add "highEntity", entity
    Set candidates = New Collection
    candidates.Add entity
    Set positions = New Collection
    positions.Add perpendicular
    record.Add "candidates", candidates
    record.Add "candidatePositions", positions
    Set CompactCoordinateRecord = record
End Function

Private Sub SetCompactPlanReferences(ByVal plan As Object)
    Dim first As Object, last As Object, a As Collection, b As Collection, pa As Collection, pb As Collection
    Dim i As Long, j As Long, ai As Long, bi As Long, distance As Double, best As Double
    Set first = plan("firstRecord")
    Set last = plan("lastRecord")
    Set a = first("candidates")
    Set b = last("candidates")
    Set pa = first("candidatePositions")
    Set pb = last("candidatePositions")
    best = SEARCH_INITIAL_LIMIT
    For i = 1 To a.Count
        For j = 1 To b.Count
            distance = Abs(CDbl(pa(i)) - CDbl(pb(j)))
            If distance < best Then
                best = distance
                ai = i
                bi = j
            End If
        Next j
    Next i
    Set plan("firstRecord") = CompactCoordinateRecord(first, a(ai), CDbl(pa(ai)))
    Set plan("lastRecord") = CompactCoordinateRecord(last, b(bi), CDbl(pb(bi)))
End Sub

Private Function MakeCoordinateRequest(ByVal v As SldWorks.View, ByVal first As Object, ByVal last As Object, ByVal axis As Long, ByVal detail As Boolean, ByVal count As Long, ByVal pitch As Double, ByVal fullRange As Boolean, ByVal compact As Boolean) As Object
    Dim request As Object
    Set request = CreateObject("Scripting.Dictionary")
    request.Add "firstRecord", first
    request.Add "lastRecord", last
    request.Add "axis", axis
    request.Add "key", CoordinateIntervalKey(v, axis, CDbl(first("coordinate")), CDbl(last("coordinate")), detail)
    request.Add "span", CDbl(last("coordinate")) - CDbl(first("coordinate"))
    request.Add "count", count
    request.Add "pitch", pitch
    request.Add "fullRange", fullRange
    If compact Then
        SetCompactPlanReferences request
    End If
    Set first = request("firstRecord")
    Set last = request("lastRecord")
    If axis = 0 Then
        request.Add "bounds", Array(CDbl(first("coordinate")), MinD(CDbl(first("low")), CDbl(last("low"))), CDbl(last("coordinate")), -MinD(-CDbl(first("low")), -CDbl(last("low"))))
    Else
        request.Add "bounds", Array(MinD(CDbl(first("low")), CDbl(last("low"))), CDbl(first("coordinate")), -MinD(-CDbl(first("low")), -CDbl(last("low"))), CDbl(last("coordinate")))
    End If
    Set MakeCoordinateRequest = request
End Function

Private Function ReferencePairCost(ByVal first As Object, ByVal last As Object) As Double
    Dim a As Collection, b As Collection, pa As Variant, pb As Variant, perpendicular As Double, difference As Double
    Set a = first("candidatePositions")
    Set b = last("candidatePositions")
    perpendicular = SEARCH_INITIAL_LIMIT
    For Each pa In a
        For Each pb In b
            perpendicular = MinD(perpendicular, Abs(CDbl(pa) - CDbl(pb)))
        Next pb
    Next pa
    difference = CDbl(last("coordinate")) - CDbl(first("coordinate"))
    ReferencePairCost = Sqr(difference ^ 2 + (perpendicular * LOCAL_REFERENCE_PERPENDICULAR_WEIGHT) ^ 2)
End Function

Private Function LocalNeighborBands(ByVal coordinates As Object, ByVal keys As Variant) As Object
    Dim result As Object, i As Long, band As Long, prior As Long, point As Variant
    Dim minimum As Double, maximum As Double, width As Double, present As Boolean, pointBand As Long, key As String
    Set result = CreateObject("Scripting.Dictionary")
    minimum = SEARCH_INITIAL_LIMIT
    maximum = -SEARCH_INITIAL_LIMIT
    For i = LBound(keys) To UBound(keys)
        For Each point In coordinates(keys(i))("candidatePositions")
            minimum = MinD(minimum, CDbl(point))
            maximum = -MinD(-maximum, -CDbl(point))
        Next point
    Next i
    width = (maximum - minimum) / LOCAL_DATUM_BAND_COUNT
    If width < MIN_PROJECTED_EXTENT_M Then
        width = MIN_PROJECTED_EXTENT_M
    End If
    For band = 0 To LOCAL_DATUM_BAND_COUNT - 1
        prior = -1
        For i = LBound(keys) To UBound(keys)
            present = False
            For Each point In coordinates(keys(i))("candidatePositions")
                pointBand = Int((CDbl(point) - minimum) / width)
                If pointBand >= LOCAL_DATUM_BAND_COUNT Then
                    pointBand = LOCAL_DATUM_BAND_COUNT - 1
                End If
                If pointBand = band Then
                    present = True
                    Exit For
                End If
            Next point
            If present Then
                If prior >= 0 Then
                    key = CStr(prior) & "|" & CStr(i)
                    If Not result.Exists(key) Then
                        result.Add key, Array(minimum + band * width, minimum + (band + 1) * width)
                    End If
                End If
                prior = i
            End If
        Next i
    Next band
    Set LocalNeighborBands = result
End Function

Private Function LocalDatumPairs(ByVal coordinates As Object) As Collection
    Dim result As New Collection, keys As Variant, connected() As Boolean, costs() As Double, parents() As Long
    Dim i As Long, stepIndex As Long, lastConnected As Long, chosen As Long, count As Long
    Dim cost As Double, best As Double, firstIndex As Long, lastIndex As Long, pair As Object, bands As Object, key As String
    count = coordinates.Count
    If count < 2 Then
        Set LocalDatumPairs = result
        Exit Function
    End If
    keys = SortedCoordinateKeys(coordinates)
    Set bands = LocalNeighborBands(coordinates, keys)
    ReDim connected(0 To count - 1)
    ReDim costs(0 To count - 1)
    ReDim parents(0 To count - 1)
    connected(0) = True
    lastConnected = 0
    For i = 1 To count - 1
        costs(i) = SEARCH_INITIAL_LIMIT
    Next i
    ' Each new link joins one previously unmeasured coordinate to the datum graph.
    ' Nearby references are preferred, so unrelated opposite-end features do not interleave.
    For stepIndex = 1 To count - 1
        best = SEARCH_INITIAL_LIMIT
        chosen = -1
        For i = 0 To count - 1
            If Not connected(i) Then
                cost = ReferencePairCost(coordinates(keys(lastConnected)), coordinates(keys(i)))
                key = CStr(MinD(lastConnected, i)) & "|" & CStr(-MinD(-lastConnected, -i))
                If bands.Exists(key) Then
                    cost = cost * LOCAL_DATUM_CHAIN_PRIORITY
                End If
                If cost < costs(i) Then
                    costs(i) = cost
                    parents(i) = lastConnected
                End If
                If costs(i) < best Then
                    best = costs(i)
                    chosen = i
                End If
            End If
        Next i
        If chosen < 0 Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "The local datum graph does not cover every coordinate."
        End If
        firstIndex = MinD(chosen, parents(chosen))
        lastIndex = -MinD(-chosen, -parents(chosen))
        Set pair = CreateObject("Scripting.Dictionary")
        pair.Add "firstRecord", coordinates(keys(firstIndex))
        pair.Add "lastRecord", coordinates(keys(lastIndex))
        pair.Add "fullRange", firstIndex = 0 And lastIndex = count - 1
        key = CStr(firstIndex) & "|" & CStr(lastIndex)
        If bands.Exists(key) Then
            pair.Add "band", bands(key)
        Else
            pair.Add "band", Empty
        End If
        result.Add pair
        connected(chosen) = True
        lastConnected = chosen
    Next stepIndex
    Set LocalDatumPairs = result
End Function

Private Function BandedCoordinateRecord(ByVal record As Object, ByVal band As Variant, ByVal viewScale As Double) As Object
    Dim candidates As Collection, positions As Collection, result As Object, i As Long, value As Double, key As Variant
    If IsEmpty(band) Then
        Set BandedCoordinateRecord = record
        Exit Function
    End If
    Set candidates = record("candidates")
    Set positions = record("candidatePositions")
    Set result = CreateObject("Scripting.Dictionary")
    For i = 1 To candidates.Count
        value = CDbl(positions(i))
        If value >= band(0) - MIN_PROJECTED_EXTENT_M And value <= band(1) + MIN_PROJECTED_EXTENT_M Then
            AddCoordinate result, candidates(i), CDbl(record("coordinate")), viewScale, value
        End If
    Next i
    For Each key In result.Keys
        Set BandedCoordinateRecord = result(key)
        Exit Function
    Next key
    Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "A local datum band has no selectable reference."
End Function

Private Function LocalDatumRequests(ByVal v As SldWorks.View, ByVal coordinates As Object, ByVal axis As Long) As Collection
    Dim requests As New Collection, pairs As Collection, pair As Variant, first As Object, last As Object
    Set pairs = LocalDatumPairs(coordinates)
    For Each pair In pairs
        Set first = BandedCoordinateRecord(pair("firstRecord"), pair("band"), v.ScaleDecimal)
        Set last = BandedCoordinateRecord(pair("lastRecord"), pair("band"), v.ScaleDecimal)
        requests.Add MakeCoordinateRequest(v, first, last, axis, True, 1, CDbl(last("coordinate")) - CDbl(first("coordinate")), CBool(pair("fullRange")), True)
    Next pair
    Set LocalDatumRequests = requests
End Function

Private Function ShiftCoordinateRecord(ByVal original As Object, ByVal origin As Variant, ByVal ratio As Double, ByVal axis As Long, ByVal viewScale As Double) As Object
    Dim result As Object, candidates As Collection, positions As Collection, i As Long, value As Double, perpendicular As Double, key As Variant
    Set result = CreateObject("Scripting.Dictionary")
    Set candidates = original("candidates")
    Set positions = original("candidatePositions")
    value = origin(axis) + (CDbl(original("coordinate")) - coordinateSourceOrigin(axis)) * ratio
    For i = 1 To candidates.Count
        perpendicular = origin(1 - axis) + (CDbl(positions(i)) - coordinateSourceOrigin(1 - axis)) * ratio
        AddCoordinate result, candidates(i), value, viewScale, perpendicular
    Next i
    For Each key In result.Keys
        Set ShiftCoordinateRecord = result(key)
        Exit Function
    Next key
End Function

Private Function AssignedRequestsForView(ByVal v As SldWorks.View, ByVal axis As Long) As Collection
    Dim result As New Collection, request As Variant, shifted As Object, first As Object, last As Object, origin As Variant, ratio As Double
    origin = ViewOrigin(v)
    ratio = v.ScaleDecimal / coordinateSourceScale
    For Each request In coordinateAssignedRequests
        If CLng(request("axis")) = axis Then
            If Not coordinateSelection Is Nothing Then
                If Not coordinateSelection.Exists(CStr(request("key"))) Then
                    GoTo NextRequest
                End If
            End If
            Set first = ShiftCoordinateRecord(request("firstRecord"), origin, ratio, axis, v.ScaleDecimal)
            Set last = ShiftCoordinateRecord(request("lastRecord"), origin, ratio, axis, v.ScaleDecimal)
            Set shifted = MakeCoordinateRequest(v, first, last, axis, True, CLng(request("count")), CDbl(request("pitch")) * ratio, CBool(request("fullRange")), False)
            shifted("key") = CStr(request("key"))
            result.Add shifted
        End If
NextRequest:
    Next request
    Set AssignedRequestsForView = result
End Function

Private Function RequestCoordinates(ByVal requests As Collection, ByVal viewScale As Double) As Object
    Dim result As Object, request As Variant, name As Variant, record As Object, candidates As Collection, positions As Collection, i As Long
    Set result = CreateObject("Scripting.Dictionary")
    For Each request In requests
        For Each name In Array("firstRecord", "lastRecord")
            Set record = request(name)
            Set candidates = record("candidates")
            Set positions = record("candidatePositions")
            For i = 1 To candidates.Count
                AddCoordinate result, candidates(i), CDbl(record("coordinate")), viewScale, CDbl(positions(i))
            Next i
        Next name
    Next request
    Set RequestCoordinates = result
End Function

Private Sub ValidateLocalDatumCoverage(ByVal v As SldWorks.View, ByVal requests As Collection)
    Dim axis As Long, nodes As Object, coordinates As Object, request As Variant, name As Variant, record As Object, key As Variant, counts(1) As Long
    For axis = 0 To 1
        Set nodes = CreateObject("Scripting.Dictionary")
        If axis = 0 Then
            Set coordinates = coordinateSourceX
        Else
            Set coordinates = coordinateSourceY
        End If
        For Each request In requests
            If CLng(request("axis")) = axis Then
                If Not dimensionKeys.Exists(CStr(request("key"))) Then
                    Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "A local datum link has no graphical dimension."
                End If
                For Each name In Array("firstRecord", "lastRecord")
                    Set record = request(name)
                    key = CStr(Round(CDbl(record("coordinate")) / coordinateSourceScale, COORDINATE_KEY_DECIMALS))
                    nodes(key) = True
                Next name
            End If
        Next request
        If coordinates.Count >= COORDINATE_MIN_KEY_COUNT Then
            For Each key In coordinates.Keys
                If Not nodes.Exists(CStr(key)) Then
                    Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "A source coordinate is missing from the graphical datum graph."
                End If
            Next key
        End If
        counts(axis) = nodes.Count
    Next axis
    Log "GOST LOCAL COVERAGE expectedX=" & coordinateSourceX.Count & "; coveredX=" & counts(0) & "; expectedY=" & coordinateSourceY.Count & "; coveredY=" & counts(1) & "; sheets added=0"
End Sub

Private Function AdjacentCoordinateRequests(ByVal v As SldWorks.View, ByVal coordinates As Object, ByVal axis As Long, ByVal detail As Boolean) As Collection
    Dim requests As New Collection, keys As Variant, i As Long, lastIndex As Long
    Dim first As Object, last As Object, plan As Object, pitch As Double, nextPitch As Double
    If coordinates.Count < COORDINATE_MIN_KEY_COUNT Then
        Set AdjacentCoordinateRequests = requests
        Exit Function
    End If
    keys = SortedCoordinateKeys(coordinates)
    i = LBound(keys)
    Do While i < UBound(keys)
        lastIndex = i + 1
        pitch = CDbl(coordinates(keys(lastIndex))("coordinate")) - CDbl(coordinates(keys(i))("coordinate"))
        Do While lastIndex < UBound(keys)
            nextPitch = CDbl(coordinates(keys(lastIndex + 1))("coordinate")) - CDbl(coordinates(keys(lastIndex))("coordinate"))
            If Abs(nextPitch - pitch) > v.ScaleDecimal * DIMENSION_MATCH_TOLERANCE_M Then
                Exit Do
            End If
            lastIndex = lastIndex + 1
        Loop
        Set first = coordinates(keys(i))
        Set last = coordinates(keys(lastIndex))
        Set plan = MakeCoordinateRequest(v, first, last, axis, detail, lastIndex - i, pitch, i = LBound(keys) And lastIndex = UBound(keys), False)
        requests.Add plan
        i = lastIndex
    Loop
    Set AdjacentCoordinateRequests = requests
End Function

