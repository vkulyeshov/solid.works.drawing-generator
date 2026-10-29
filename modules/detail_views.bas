Attribute VB_Name = "detail_views"
Option Explicit
Private Function UnionDetailBounds(ByVal a As Variant, ByVal b As Variant) As Variant
    UnionDetailBounds = Array(MinD(a(0), b(0)), MinD(a(1), b(1)), -MinD(-a(2), -b(2)), -MinD(-a(3), -b(3)))
End Function

Private Function DetailProfileRadius(ByVal bounds As Variant) As Double
    DetailProfileRadius = Sqr((bounds(2) - bounds(0)) ^ 2 + (bounds(3) - bounds(1)) ^ 2) / 2 + DETAIL_PROFILE_MARGIN_M
End Function

Private Function CreateDetailProfile(ByVal d As SldWorks.ModelDoc2, ByVal v As SldWorks.View, ByVal x As Double, ByVal y As Double, ByVal paperRadius As Double) As SldWorks.SketchSegment
    Dim sketch As SldWorks.Sketch, util As SldWorks.MathUtility, point As SldWorks.MathPoint, coordinates As Variant
    AsDrawing(d).ActivateView v.Name
    d.ClearSelection2 True
    Set sketch = v.GetSketch
    Set util = sw.GetMathUtility
    Set point = util.CreatePoint(Array(x, y, 0#))
    Set point = point.MultiplyTransform(sketch.ModelToSketchTransform)
    coordinates = point.ArrayData
    ' Drawing-view sketch coordinates are unscaled and relative to the view origin.
    Set CreateDetailProfile = d.SketchManager.CreateCircleByRadius(coordinates(0), coordinates(1), 0, paperRadius / v.ScaleDecimal)
End Function

Private Sub AddSameSheetDimensionDetails(ByVal d As SldWorks.ModelDoc2, ByVal overview As SldWorks.View)
    Dim requests As New Collection, groups As New Collection, axisRequests As Collection
    Dim plan As Variant, group As Variant, chosen As Object, keys As Object, remaining As Object
    Dim bounds As Variant, combined As Variant, cell As Variant, extent As Variant
    Dim width As Double, height As Double, top As Double, bottom As Double, cellWidth As Double, cellHeight As Double
    Dim maxRadius As Double, radius As Double, cost As Double, best As Double, detailViewScale As Double
    Dim i As Long, axis As Long, result As Long, row As Long, column As Long
    Dim v As SldWorks.View, profile As SldWorks.SketchSegment, x As Double, y As Double
    Dim number As Long, message As String, covered As Long, expected As Long
    On Error GoTo Failed
    Log "GOST DETAILS collecting graphical references"
    CollectViewCoordinates overview, coordinateSourceX, coordinateSourceY
    Set coordinateSourceView = overview
    coordinateSourceOrigin = ViewOrigin(overview)
    coordinateSourceScale = overview.ScaleDecimal
    For axis = 0 To 1
        Log "GOST DETAILS planning axis=" & axis
        If axis = 0 Then
            Set axisRequests = LocalDatumRequests(overview, coordinateSourceX, axis)
        Else
            Set axisRequests = LocalDatumRequests(overview, coordinateSourceY, axis)
        End If
        For Each plan In axisRequests
            requests.Add plan
            expected = expected + CLng(plan("count"))
        Next plan
    Next axis
    Set coordinateAssignedRequests = requests
    Log "GOST LOCAL coverage x=" & coordinateSourceX.Count & "; y=" & coordinateSourceY.Count & "; graph edges=" & requests.Count
    If FitFlatDimensionLayout(d, overview) Then
        result = OverallDimensions(d, overview, True)
        DimensionFeatures d, overview, True
        ValidateLocalDatumCoverage overview, requests
        Set coordinateSelection = Nothing
        Set coordinateSourceView = Nothing
        Set coordinateAssignedRequests = Nothing
        coordinateRegion = Empty
        Exit Sub
    End If
    width = CoordinateRegionEdge(2) - CoordinateRegionEdge(0)
    bottom = CoordinateRegionEdge(1)
    top = CoordinateRegionEdge(3) - DETAIL_OVERVIEW_HEIGHT_M
    cellWidth = (width - (DETAIL_GRID_COLUMNS - 1) * DETAIL_CELL_GAP_M) / DETAIL_GRID_COLUMNS
    cellHeight = (top - bottom - (DETAIL_GRID_ROWS - 1) * DETAIL_CELL_GAP_M) / DETAIL_GRID_ROWS
    maxRadius = MinD(cellWidth - DETAIL_DIMENSION_BAND_M * 2, cellHeight - DETAIL_DIMENSION_BAND_M * 2) / (DETAIL_MIN_ZOOM_FACTOR * 2)
    Set remaining = CreateObject("Scripting.Dictionary")
    For Each plan In requests
        Set chosen = Nothing
        bounds = plan("bounds")
        radius = DetailProfileRadius(bounds)
        If CDbl(plan("span")) < DETAIL_TRIGGER_SPAN_M And radius <= maxRadius Then
            best = SEARCH_INITIAL_LIMIT
            For Each group In groups
                If CLng(group("count")) < DETAIL_MAX_INTERVALS Then
                    combined = UnionDetailBounds(group("bounds"), bounds)
                    cost = DetailProfileRadius(combined)
                    If cost <= maxRadius And cost < best Then
                        Set chosen = group
                        best = cost
                    End If
                End If
            Next group
            If chosen Is Nothing And groups.Count < DETAIL_GRID_COLUMNS * DETAIL_GRID_ROWS Then
                Set chosen = CreateObject("Scripting.Dictionary")
                chosen.Add "bounds", bounds
                Set keys = CreateObject("Scripting.Dictionary")
                chosen.Add "keys", keys
                chosen.Add "count", 0
                groups.Add chosen
            End If
        End If
        If chosen Is Nothing Then
            remaining.Add CStr(plan("key")), True
        Else
            chosen("bounds") = UnionDetailBounds(chosen("bounds"), bounds)
            Set keys = chosen("keys")
            keys.Add CStr(plan("key")), True
            chosen("count") = CLng(chosen("count")) + 1
        End If
    Next plan
    Log "GOST DETAILS planned=" & groups.Count & "; overview intervals=" & remaining.Count & "; total covered intervals=" & expected
    If groups.Count = 0 Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "No compact enlarged detail can contain both references of the crowded intervals on this paper format."
    End If
    coordinateRegion = Array(CoordinateRegionEdge(0), top + DETAIL_CELL_GAP_M, CoordinateRegionEdge(2), sheetHeight - COORDINATE_AXIS_TOP_MARGIN_M)
    Set coordinateSelection = remaining
    bounds = ProjectedFlatBounds(overview)
    detailViewScale = StandardScale(overview.ScaleDecimal * MinD((width - COORDINATE_BASE_OFFSET_M * 2) / (bounds(2) - bounds(0)), (DETAIL_OVERVIEW_HEIGHT_M - COORDINATE_BASE_OFFSET_M * 2) / (bounds(3) - bounds(1))))
    overview.ScaleDecimal = MinD(coordinateSourceScale, detailViewScale)
    CenterView overview, (coordinateRegion(0) + coordinateRegion(2)) / 2, (coordinateRegion(1) + coordinateRegion(3)) / 2
    d.ForceRebuild3 False
    result = CoordinateDimensions(d, overview, True, True)
    ' Validate every view before creating any dimensions. A failed layout never drops a request.
    For i = 1 To groups.Count
        Set group = groups(i)
        column = (i - 1) Mod DETAIL_GRID_COLUMNS
        row = (i - 1) \ DETAIL_GRID_COLUMNS
        x = COORDINATE_LEFT_LIMIT_M + column * (cellWidth + DETAIL_CELL_GAP_M)
        y = top - row * (cellHeight + DETAIL_CELL_GAP_M)
        cell = Array(x, y - cellHeight, x + cellWidth, y)
        bounds = group("bounds")
        radius = DetailProfileRadius(bounds)
        detailViewScale = StandardScale(coordinateSourceScale * MinD(cellWidth - DETAIL_DIMENSION_BAND_M * 2, cellHeight - DETAIL_DIMENSION_BAND_M * 2) / (radius * 2))
        x = DetailCoordinate((bounds(0) + bounds(2)) / 2, 0, overview)
        y = DetailCoordinate((bounds(1) + bounds(3)) / 2, 1, overview)
        Set profile = CreateDetailProfile(d, overview, x, y, radius * overview.ScaleDecimal / coordinateSourceScale)
        Set v = AsDrawing(d).CreateDetailViewAt4((cell(0) + cell(2)) / 2, (cell(1) + cell(3)) / 2, 0, SW_DETAIL_STANDARD, detailViewScale, 1, Chr$(DETAIL_FIRST_LABEL_CODE + i - 1), SW_DETAIL_CIRCLE, True, False, False, SW_DETAIL_OUTLINE_INTENSITY)
        If v Is Nothing Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot create a native enlarged detail view."
        End If
        v.UseParentScale = False
        v.ScaleDecimal = detailViewScale
        d.ForceRebuild3 False
        CenterView v, (cell(0) + cell(2)) / 2, (cell(1) + cell(3)) / 2
        group.Add "view", v
        group.Add "region", cell
        coordinateRegion = cell
        Set coordinateSelection = group("keys")
        result = CoordinateDimensions(d, v, True, True)
        Log "GOST DETAIL " & v.Name & "; scale=" & detailViewScale & "; intervals=" & coordinateSelection.Count
    Next i
    coordinateRegion = Array(COORDINATE_LEFT_LIMIT_M, top + DETAIL_CELL_GAP_M, sheetWidth - COORDINATE_AXIS_RIGHT_MARGIN_M, sheetHeight - COORDINATE_AXIS_TOP_MARGIN_M)
    Set coordinateSelection = remaining
    result = OverallDimensions(d, overview)
    result = CoordinateDimensions(d, overview, True)
    DimensionFeatures d, overview, True, False
    For Each group In groups
        Set v = group("view")
        coordinateRegion = group("region")
        Set coordinateSelection = group("keys")
        result = CoordinateDimensions(d, v, True)
        DimensionFeatures d, v, True, False
    Next group
    For Each plan In requests
        If Not dimensionKeys.Exists(CStr(plan("key"))) Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "An assigned interval is missing from the completed graphical layout."
        End If
        covered = covered + CLng(plan("count"))
    Next plan
    Log "GOST DETAIL COVERAGE expected=" & expected & "; covered=" & covered & "; sheets added=0"
    ValidateLocalDatumCoverage overview, requests
    Set coordinateSelection = Nothing
    Set coordinateSourceView = Nothing
    Set coordinateAssignedRequests = Nothing
    coordinateRegion = Empty
    Exit Sub
Failed:
    number = Err.Number
    message = Err.Description
    Set coordinateSelection = Nothing
    Set coordinateSourceView = Nothing
    Set coordinateAssignedRequests = Nothing
    coordinateRegion = Empty
    Err.Raise number, "AddSameSheetDimensionDetails", "Same-sheet enlarged layout: " & message & " Select a larger paper format if both references cannot fit in a local detail."
End Sub

Private Function FitFlatDimensionLayout(ByVal d As SldWorks.ModelDoc2, ByVal v As SldWorks.View) As Boolean
    Dim bounds As Variant, position As Variant, result As Long, number As Long, message As String, attempt As Long, initialScale As Double
    initialScale = v.ScaleDecimal
    Do
        attempt = attempt + 1
        d.ForceRebuild3 False
        bounds = ProjectedFlatBounds(v)
        position = v.Position
        position(0) = position(0) + (sheetWidth + DETAIL_CENTER_X_BIAS_M) / 2 - (bounds(OUTLINE_LEFT_INDEX) + bounds(OUTLINE_RIGHT_INDEX)) / 2
        position(1) = position(1) + (COORDINATE_BOTTOM_LIMIT_M + sheetHeight - COORDINATE_AXIS_TOP_MARGIN_M) / 2 - (bounds(OUTLINE_BOTTOM_INDEX) + bounds(OUTLINE_TOP_INDEX)) / 2
        v.Position = position
        d.ForceRebuild3 False
        On Error Resume Next
        result = CoordinateDimensions(d, v, True, True)
        number = Err.Number
        message = Err.Description
        Err.Clear
        On Error GoTo 0
        If number = 0 Then
            Log "DIMENSION layout scale=" & v.ScaleDecimal & "; row pitch=" & COORDINATE_LANE_PITCH_M * MILLIMETRES_PER_METRE & " mm; minimum outside offset=" & GOST_OBJECT_OFFSET_M * MILLIMETRES_PER_METRE & " mm; strict=" & Setting("StrictDimensionLayout", "0")
            FitFlatDimensionLayout = True
            Exit Function
        End If
        If number <> vbObjectError + ERR_DIMENSION_LAYOUT Then
            Err.Raise number, "FitFlatDimensionLayout", message
        End If
        If attempt >= DIMENSION_LAYOUT_SCALE_ATTEMPTS Or v.ScaleDecimal <= SCALE_1_TO_1000 Then
            Log "GOST overview requires same-sheet enlarged details: " & message
            v.ScaleDecimal = initialScale
            d.ForceRebuild3 False
            Exit Function
        End If
        Log "GOST layout retry scale=" & v.ScaleDecimal & ": " & message
        v.ScaleDecimal = StandardScale(v.ScaleDecimal * DIMENSION_LAYOUT_RETRY_SCALE_FACTOR)
    Loop
End Function

Private Sub WriteDimensionCoverage(ByVal d As SldWorks.ModelDoc2, ByVal outName As String)
    Dim names As Variant, name As Variant, v As SldWorks.View, stream As Object, path As String, temp As String
    Dim previous As String, line As Variant
    path = runDir & COVERAGE_RELATIVE_PATH
    If fs.FileExists(path) Then
        Set stream = fs.OpenTextFile(path, FILE_FOR_READING, False, FILE_UNICODE)
        previous = stream.ReadAll
        stream.Close
    End If
    temp = path & ATOMIC_TEMP_SUFFIX
    Set stream = fs.CreateTextFile(temp, True, True)
    stream.WriteLine "Drawing" & vbTab & "Sheet" & vbTab & "View" & vbTab & "Edges" & vbTab & "Dimensions" & vbTab & "Configuration"
    For Each line In Split(previous, vbCrLf)
        If Len(CStr(line)) > 0 And Left$(CStr(line), COVERAGE_HEADER_PREFIX_LENGTH) <> "Drawing" & vbTab Then
            If Left$(CStr(line), Len(outName) + 1) <> outName & vbTab Then
                stream.WriteLine line
            End If
        End If
    Next line
    names = AsDrawing(d).GetSheetNames
    For Each name In names
        AsDrawing(d).ActivateSheet CStr(name)
        Set v = AsDrawing(d).GetFirstView: Set v = v.GetNextView
        Do While Not v Is Nothing
            stream.WriteLine outName & vbTab & name & vbTab & v.Name & vbTab & VisibleEdgeCount(v) & vbTab & v.GetDisplayDimensionCount & vbTab & v.ReferencedConfiguration
            Set v = v.GetNextView
        Loop
    Next name
    stream.Close
    If MoveFileExW(StrPtr(temp), StrPtr(path), ATOMIC_REPLACE_WRITE_THROUGH) = 0 Then
        Err.Raise vbObjectError + ERR_GENERATION_DATA, , "Cannot commit dimension-coverage.tsv"
    End If
End Sub
' Native graphical dimensions with outside routing and collision-free text bands.
' Circle centers participate as real circular edges, never as loose sketch points.
