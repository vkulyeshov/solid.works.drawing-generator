Attribute VB_Name = "pattern_dimensions"
Option Explicit
Private Function HoleAxisOrigin(ByVal parameters As Variant, ByVal normal As Variant) As Variant
    Dim station As Double
    station = parameters(0) * normal(0) + parameters(1) * normal(1) + parameters(2) * normal(2)
    HoleAxisOrigin = Array(parameters(0) - station * normal(0), parameters(1) - station * normal(1), parameters(2) - station * normal(2))
End Function

Private Function HolePatternShape(ByVal points As Variant, ByVal normal As Variant) As Variant
    Dim center As Variant, first As Variant, vector As Variant, target As Variant, point As Variant, rotated As Variant
    Dim index As Long, stepIndex As Long, axis As Long, count As Long, radius As Double, length As Double, station As Double
    Dim angle As Double, regular As Boolean, matched As Boolean, distance As Double
    center = Array(0, 0, 0)
    count = UBound(points) - LBound(points) + 1
    For Each point In points
        For axis = 0 To SPACE_LAST_AXIS_INDEX
            center(axis) = center(axis) + point(axis) / count
        Next axis
    Next point
    point = points(LBound(points))
    first = Array(point(0) - center(0), point(1) - center(1), point(2) - center(2))
    station = first(0) * normal(0) + first(1) * normal(1) + first(2) * normal(2)
    For axis = 0 To SPACE_LAST_AXIS_INDEX
        first(axis) = first(axis) - station * normal(axis)
    Next axis
    radius = Sqr(first(0) ^ 2 + first(1) ^ 2 + first(2) ^ 2)
    regular = count >= PATTERN_MIN_CIRCULAR_HOLES And radius > NATIVE_GEOMETRY_KEY_TOLERANCE_M
    For stepIndex = 0 To count - 1
        angle = FULL_REVOLUTION_RAD * stepIndex / count
        rotated = Array(first(0) * Cos(angle) + (normal(1) * first(2) - normal(2) * first(1)) * Sin(angle), first(1) * Cos(angle) + (normal(2) * first(0) - normal(0) * first(2)) * Sin(angle), first(2) * Cos(angle) + (normal(0) * first(1) - normal(1) * first(0)) * Sin(angle))
        matched = False
        For Each point In points
            vector = Array(point(0) - center(0), point(1) - center(1), point(2) - center(2))
            station = vector(0) * normal(0) + vector(1) * normal(1) + vector(2) * normal(2)
            distance = 0
            For axis = 0 To SPACE_LAST_AXIS_INDEX
                distance = distance + (vector(axis) - station * normal(axis) - rotated(axis)) ^ 2
            Next axis
            If distance <= NATIVE_GEOMETRY_KEY_TOLERANCE_M ^ 2 Then
                matched = True
                Exit For
            End If
        Next point
        If Not matched Then
            regular = False
            Exit For
        End If
    Next stepIndex
    HolePatternShape = Array(center(0), center(1), center(2), radius, count, regular)
End Function

Private Function VisiblePatternEdges(ByVal view As SldWorks.View, ByVal pattern As Object) As Object
    Dim result As Object, probe As SldWorks.Edge, components As Variant, component As Variant, edges As Variant, item As Variant
    Dim edge As SldWorks.Edge, curve As SldWorks.Curve, parameters As Variant, key As String, normal As Variant
    Set result = CreateObject("Scripting.Dictionary")
    Set probe = VisibleHoleEdge(view, pattern)
    If probe Is Nothing Then
        Set VisiblePatternEdges = result
        Exit Function
    End If
    normal = pattern("normal")
    components = view.GetVisibleComponents
    For Each component In components
        edges = view.GetVisibleEntities2(component, SW_VISIBLE_ENTITY_EDGE)
        If Not IsEmpty(edges) Then
            For Each item In edges
                Set edge = item
                Set curve = edge.GetCurve
                If Not curve Is Nothing Then
                    If curve.IsCircle Then
                        parameters = curve.CircleParams
                        If Abs(parameters(CIRCLE_RADIUS_INDEX) - CDbl(pattern("radius"))) < NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
                            key = NativePointKey(HoleAxisOrigin(parameters, normal))
                            If pattern("centers").Exists(key) And Not result.Exists(key) Then
                                result.Add key, edge
                            End If
                        End If
                    End If
                End If
            Next item
        End If
    Next component
    Set VisiblePatternEdges = result
End Function

Private Function PatternSheetPoint(ByVal view As SldWorks.View, ByVal parameters As Variant) As Variant
    Dim transform As Variant
    transform = view.ModelToViewTransform.ArrayData
    PatternSheetPoint = Array((parameters(0) * transform(0) + parameters(1) * transform(3) + parameters(2) * transform(6)) * transform(MATRIX_SCALE_INDEX) + transform(9), (parameters(0) * transform(1) + parameters(1) * transform(4) + parameters(2) * transform(7)) * transform(MATRIX_SCALE_INDEX) + transform(10), 0)
End Function

Private Sub AddPatternSpan(ByVal d As SldWorks.ModelDoc2, ByVal view As SldWorks.View, ByVal first As SldWorks.Edge, ByVal last As SldWorks.Edge, ByVal axis As Long, ByVal pattern As Object, ByVal regular As Boolean)
    Dim a As Variant, b As Variant, outline As Variant, display As SldWorks.DisplayDimension, dimension As SldWorks.Dimension, expected As Double
    Dim key As String, parameters As Variant, shape As Variant, center As Variant, angle As Double, increment As Double, dx As Double, dy As Double
    Dim sketch As SldWorks.Sketch, sketchTransform As Variant, pitchCenter As Variant, pitchCircle As SldWorks.SketchSegment
    a = PatternSheetPoint(view, first.GetCurve.CircleParams)
    b = PatternSheetPoint(view, last.GetCurve.CircleParams)
    expected = Sqr((a(0) - b(0)) ^ 2 + (a(1) - b(1)) ^ 2) / view.ScaleDecimal
    If axis >= 0 Then
        expected = Abs(a(axis) - b(axis)) / view.ScaleDecimal
    End If
    Log "PATTERN SPAN axis=" & axis & "; expected=" & expected
    If expected < NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
        Exit Sub
    End If
    If regular Then
        shape = pattern("shape")
        center = PatternSheetPoint(view, shape)
        Set sketch = view.GetSketch
        sketchTransform = sketch.ModelToSketchTransform.ArrayData
        pitchCenter = SourcePointInSketch(center, sketchTransform)
        AsDrawing(d).ActivateView CStr(view.Name)
        d.ClearSelection2 True
        Set pitchCircle = d.SketchManager.CreateCircleByRadius(pitchCenter(0), pitchCenter(1), 0, shape(PATTERN_RADIUS_INDEX) * view.ScaleDecimal * Abs(sketchTransform(MATRIX_SCALE_INDEX)))
        If pitchCircle Is Nothing Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot create the associated pitch circle."
        End If
        pitchCircle.ConstructionGeometry = True
        pitchCircle.Color = DIMENSION_COLOR_BLACK
        pitchCircle.Select4 False, Nothing
    Else
        If Not SelectFeatureEdge(d, view, first) Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot select a pattern reference."
        End If
        If Not SelectFeatureEdge(d, view, last, True) Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot select the second pattern reference."
        End If
    End If
    outline = view.GetOutline
    If regular Then
        Set display = d.AddDiameterDimension2(outline(2) + GOST_OBJECT_OFFSET_M, outline(3) + GOST_OBJECT_OFFSET_M, 0)
    ElseIf axis = SCREEN_X_AXIS_INDEX Then
        Set display = d.AddHorizontalDimension2((a(0) + b(0)) / 2, outline(3) + GOST_OBJECT_OFFSET_M, 0)
    ElseIf axis = SCREEN_Y_AXIS_INDEX Then
        Set display = d.AddVerticalDimension2(outline(0) - GOST_OBJECT_OFFSET_M, (a(1) + b(1)) / 2, 0)
    Else
        Set display = d.AddDimension2(outline(2) + GOST_OBJECT_OFFSET_M, outline(3) + GOST_OBJECT_OFFSET_M, 0)
    End If
    If display Is Nothing Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot create a pattern location dimension."
    End If
    Set dimension = display.GetDimension2(SW_DIMENSION_INDEX_FIRST)
    If Not regular Then
        dimension.SetArcEndCondition DIMENSION_REFERENCE_FIRST, SW_ARC_CENTER
        dimension.SetArcEndCondition DIMENSION_REFERENCE_SECOND, SW_ARC_CENTER
    End If
    If Abs(dimension.SystemValue - expected) > DIMENSION_MATCH_TOLERANCE_M Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "A pattern location dimension does not measure the hole centers."
    End If
    key = "pattern|" & CStr(pattern("feature")) & "|" & CStr(axis)
    If regular Then
        shape = pattern("shape")
        center = PatternSheetPoint(view, shape)
        dx = a(0) - center(0)
        dy = a(1) - center(1)
        If Abs(dx) < DIRECTION_ZERO_TOLERANCE Then
            angle = Sgn(dy) * PI / 2
        Else
            angle = Atn(dy / dx)
            If dx < 0 Then
                angle = angle + PI
            End If
        End If
        If angle < 0 Then
            angle = angle + FULL_REVOLUTION_RAD
        End If
        increment = FULL_REVOLUTION_RAD / pattern("centers").Count
        angle = angle - Int((angle + DIRECTION_ZERO_TOLERANCE) / increment) * increment
        display.SetText SW_DIMENSION_TEXT_SUFFIX, " (" & CStr(pattern("centers").Count) & " отв. через " & CStr(Round(increment * DEGREES_PER_RADIAN, ANGULAR_DECIMAL_PLACES)) & "°; " & CStr(Round(Abs(angle) * DEGREES_PER_RADIAN, ANGULAR_DECIMAL_PLACES)) & "° до осі)"
        key = "pattern-circle|" & CStr(pattern("feature"))
    End If
    featureDimensionOwners.Add display.GetNameForSelection, key
    If Not regular And axis = SCREEN_X_AXIS_INDEX Then
        featureDimensionPlacement(display.GetNameForSelection) = NATIVE_LABEL_SIDE_ABOVE
    ElseIf Not regular And axis = SCREEN_Y_AXIS_INDEX Then
        featureDimensionPlacement(display.GetNameForSelection) = NATIVE_LABEL_SIDE_LEFT
    End If
    If regular Then
        display.GetAnnotation.Visible = SW_ANNOTATION_VISIBLE
    End If
    Log "PATTERN LOCATION " & key & "; measured=" & dimension.SystemValue
End Sub

Private Function PatternBoundaryReference(ByVal view As SldWorks.View, ByVal edges As Object, ByVal axis As Long, ByVal orientation As Long) As Object
    Dim item As Variant, hole As SldWorks.Edge, lastHole As SldWorks.Edge, point As Variant, lastPoint As Variant
    Dim component As Variant, components As Variant, visible As Variant, edge As SldWorks.Edge, endpoint As SldWorks.Vertex
    Dim endIndex As Long, maximum As Double, nearest As Double, expected As Double, outline As Variant
    Dim result As Object
    Dim bestReference As Object, curve As SldWorks.Curve, parameters As Variant, matrix As Variant, aligned As Double
    matrix = view.ModelToViewTransform.ArrayData
    maximum = -SEARCH_INITIAL_LIMIT
    For Each item In edges.Items
        Set hole = item
        point = PatternSheetPoint(view, hole.GetCurve.CircleParams)
        point(axis) = point(axis) * orientation
        If point(axis) > maximum Then
            maximum = point(axis)
            lastPoint = point
            Set lastHole = hole
        End If
    Next item
    maximum = -SEARCH_INITIAL_LIMIT
    nearest = SEARCH_INITIAL_LIMIT
    components = view.GetVisibleComponents
    For Each component In components
        visible = view.GetVisibleEntities2(component, SW_VISIBLE_ENTITY_EDGE)
        If Not IsEmpty(visible) Then
            For Each item In visible
                Set edge = item
                For endIndex = EDGE_ENDPOINT_FIRST To EDGE_ENDPOINT_LAST
                    If endIndex = EDGE_ENDPOINT_FIRST Then
                        Set endpoint = edge.GetStartVertex
                    Else
                        Set endpoint = edge.GetEndVertex
                    End If
                    If Not endpoint Is Nothing Then
                        point = PatternSheetPoint(view, endpoint.GetPoint)
                        point(axis) = point(axis) * orientation
                        If point(axis) > maximum + NATIVE_GEOMETRY_KEY_TOLERANCE_M Or (Abs(point(axis) - maximum) < NATIVE_GEOMETRY_KEY_TOLERANCE_M And Abs(point(SCREEN_Y_AXIS_INDEX - axis) - lastPoint(SCREEN_Y_AXIS_INDEX - axis)) < nearest) Then
                            maximum = point(axis)
                            nearest = Abs(point(SCREEN_Y_AXIS_INDEX - axis) - lastPoint(SCREEN_Y_AXIS_INDEX - axis))
                            Set bestReference = endpoint
                        End If
                    End If
                Next endIndex
                ' An edge-on circle can define an end plane without any endpoint vertices.
                Set curve = edge.GetCurve
                If curve.IsCircle Then
                    parameters = curve.CircleParams
                    aligned = Abs(parameters(CIRCLE_AXIS_X_INDEX) * matrix(MATRIX_XX_INDEX + axis) + parameters(CIRCLE_AXIS_Y_INDEX) * matrix(MATRIX_YX_INDEX + axis) + parameters(CIRCLE_AXIS_Z_INDEX) * matrix(MATRIX_ZX_INDEX + axis))
                    If aligned > FACE_NORMAL_ALIGNMENT_COSINE Then
                        point = PatternSheetPoint(view, parameters)
                        point(axis) = point(axis) * orientation
                        If point(axis) > maximum + NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
                            maximum = point(axis)
                            nearest = Abs(point(SCREEN_Y_AXIS_INDEX - axis) - lastPoint(SCREEN_Y_AXIS_INDEX - axis))
                            Set bestReference = edge
                        End If
                    End If
                End If
            Next item
        End If
    Next component
    expected = (maximum - lastPoint(axis)) / view.ScaleDecimal
    If expected <= NATIVE_GEOMETRY_KEY_TOLERANCE_M Or bestReference Is Nothing Then
        Exit Function
    End If
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "expected", expected
    result.Add "coordinate", maximum * orientation
    result.Add "axis", axis
    result.Add "orientation", orientation
    result.Add "hole", lastHole
    result.Add "reference", bestReference
    Set PatternBoundaryReference = result
End Function

Private Sub AddPatternBoundaryOffset(ByVal d As SldWorks.ModelDoc2, ByVal view As SldWorks.View, ByVal edges As Object, ByVal pattern As Object, ByVal orientation As Long)
    Dim axis As Long, side As Variant, candidate As Object, best As Object, expected As Double, outline As Variant
    Dim lastHole As SldWorks.Edge, bestReference As Object, selectedCoordinate As Double
    Dim display As SldWorks.DisplayDimension, dimension As SldWorks.Dimension
    expected = SEARCH_INITIAL_LIMIT
    ' Choose a nearby end datum in model geometry, independent of view rotation.
    For axis = SCREEN_X_AXIS_INDEX To SCREEN_Y_AXIS_INDEX
        For Each side In Array(orientation, -orientation)
            Set candidate = PatternBoundaryReference(view, edges, axis, CLng(side))
            If Not candidate Is Nothing Then
                If CDbl(candidate("expected")) < expected - NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
                    expected = CDbl(candidate("expected"))
                    Set best = candidate
                End If
            End If
        Next side
    Next axis
    If best Is Nothing Then
        Exit Sub
    End If
    axis = CLng(best("axis"))
    orientation = CLng(best("orientation"))
    selectedCoordinate = CDbl(best("coordinate"))
    Set lastHole = best("hole")
    Set bestReference = best("reference")
    If Not SelectFeatureEdge(d, view, lastHole) Or Not SelectFeatureReference(d, view, bestReference, True) Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot select the pattern boundary references."
    End If
    outline = view.GetOutline
    If axis = SCREEN_X_AXIS_INDEX Then
        Set display = d.AddHorizontalDimension2(selectedCoordinate, outline(3) + GOST_OBJECT_OFFSET_M, 0)
    Else
        Set display = d.AddVerticalDimension2(outline(0) - GOST_OBJECT_OFFSET_M, selectedCoordinate, 0)
    End If
    If display Is Nothing Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot dimension the pattern boundary offset."
    End If
    Set dimension = display.GetDimension2(SW_DIMENSION_INDEX_FIRST)
    dimension.SetArcEndCondition DIMENSION_REFERENCE_FIRST, SW_ARC_CENTER
    If TypeOf bestReference Is SldWorks.Edge Then
        dimension.SetArcEndCondition DIMENSION_REFERENCE_SECOND, SW_ARC_CENTER
    End If
    If Abs(dimension.SystemValue - expected) > DIMENSION_MATCH_TOLERANCE_M Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "The pattern boundary offset does not measure the hole center."
    End If
    featureDimensionOwners.Add display.GetNameForSelection, "pattern-boundary|" & CStr(pattern("feature")) & "|" & CStr(axis) & "|" & CStr(orientation)
    If axis = SCREEN_X_AXIS_INDEX Then
        featureDimensionPlacement(display.GetNameForSelection) = NATIVE_LABEL_SIDE_BELOW
    Else
        featureDimensionPlacement(display.GetNameForSelection) = NATIVE_LABEL_SIDE_RIGHT
    End If
    Log "PATTERN BOUNDARY axis=" & axis & "; side=" & orientation & "; measured=" & dimension.SystemValue
End Sub

Private Sub AddHolePatternLocations(ByVal d As SldWorks.ModelDoc2, ByVal evidence As Object)
    Dim item As Variant, pattern As Object, shape As Variant, sheets As Variant, sheetIndex As Long, sheetName As String
    Dim view As SldWorks.View, edges As Object, keys As Variant, i As Long, j As Long, first As SldWorks.Edge, last As SldWorks.Edge
    Dim a As Variant, b As Variant, distance As Double, maximum As Double, firstIndex As Long, lastIndex As Long, graphic As Object, finished As Boolean
    Dim patterns As Collection, patternIndex As Long
    Set featurePatternGraphics = New Collection
    Set patterns = evidence("patterns")
    Log "PATTERN LOCATIONS groups=" & patterns.Count
    For patternIndex = 1 To patterns.Count
        Set pattern = patterns(patternIndex)
        shape = HolePatternShape(pattern("centers").Items, pattern("normal"))
        Log "PATTERN SHAPE " & pattern("feature") & "; regular=" & CStr(shape(PATTERN_REGULAR_INDEX)) & "; radius=" & CStr(shape(PATTERN_RADIUS_INDEX))
        pattern("shape") = shape
        sheets = AsDrawing(d).GetSheetNames
        finished = False
        For sheetIndex = UBound(sheets) To LBound(sheets) Step -1
            sheetName = CStr(sheets(sheetIndex))
            AsDrawing(d).ActivateSheet sheetName
            Set view = AsDrawing(d).GetFirstView.GetNextView
            Do While Not view Is Nothing
                Set edges = VisiblePatternEdges(view, pattern)
                Log "PATTERN EDGES " & view.Name & "; count=" & edges.Count
                If edges.Count = pattern("centers").Count And edges.Count >= PATTERN_MIN_POSITION_HOLES Then
                    keys = edges.Keys
                    maximum = 0
                    For i = LBound(keys) To UBound(keys)
                        Set first = edges(keys(i))
                        a = PatternSheetPoint(view, first.GetCurve.CircleParams)
                        For j = i + 1 To UBound(keys)
                            Set last = edges(keys(j))
                            b = PatternSheetPoint(view, last.GetCurve.CircleParams)
                            distance = (a(0) - b(0)) ^ 2 + (a(1) - b(1)) ^ 2
                            If distance > maximum Then
                                maximum = distance
                                firstIndex = i
                                lastIndex = j
                            End If
                        Next j
                    Next i
                    Set first = edges(keys(firstIndex))
                    Set last = edges(keys(lastIndex))
                    If CBool(shape(PATTERN_REGULAR_INDEX)) And shape(PATTERN_COUNT_INDEX) Mod 2 = 0 Then
                        AddPatternSpan d, view, first, last, PATTERN_ALIGNED_AXIS, pattern, True
                        Set graphic = CreateObject("Scripting.Dictionary")
                        graphic.Add "sheet", sheetName
                        graphic.Add "view", view
                        graphic.Add "shape", shape
                        featurePatternGraphics.Add graphic
                    Else
                        AddPatternSpan d, view, first, last, SCREEN_X_AXIS_INDEX, pattern, False
                        AddPatternSpan d, view, first, last, SCREEN_Y_AXIS_INDEX, pattern, False
                        AddPatternBoundaryOffset d, view, edges, pattern, -1
                        PreferPatternProfileDimensions d, view
                    End If
                    finished = True
                    Exit Do
                End If
                Set view = view.GetNextView
            Loop
            If finished Then
                Exit For
            End If
        Next sheetIndex
        If Not finished Then
            Err.Raise vbObjectError + ERR_FEATURE_VIEW_REQUIRED, , "No orthographic view exposes the complete hole pattern."
        End If
    Next patternIndex
End Sub

Private Sub RenderHolePatternCircles(ByVal d As SldWorks.ModelDoc2)
    Dim item As Variant, graphic As Object, view As SldWorks.View, shape As Variant, point As Variant, segment As SldWorks.SketchSegment
    Dim sheet As SldWorks.Sheet, properties As Variant, sheetScale As Double, radius As Double
    If featurePatternGraphics Is Nothing Then
        Exit Sub
    End If
    For Each item In featurePatternGraphics
        Set graphic = item
        AsDrawing(d).ActivateSheet CStr(graphic("sheet"))
        AsDrawing(d).ActivateView ""
        Set view = graphic("view")
        shape = graphic("shape")
        point = PatternSheetPoint(view, shape)
        radius = shape(PATTERN_RADIUS_INDEX) * view.ScaleDecimal
        Set sheet = AsDrawing(d).GetCurrentSheet
        properties = sheet.GetProperties
        sheetScale = properties(SHEET_SCALE_NUMERATOR_PROPERTY) / properties(SHEET_SCALE_DENOMINATOR_PROPERTY)
        Set segment = d.SketchManager.CreateCenterLine((point(0) - radius - NATIVE_CENTERLINE_EXTENSION_M) / sheetScale, point(1) / sheetScale, 0, (point(0) + radius + NATIVE_CENTERLINE_EXTENSION_M) / sheetScale, point(1) / sheetScale, 0)
        segment.Color = DIMENSION_COLOR_BLACK
        Set segment = d.SketchManager.CreateCenterLine(point(0) / sheetScale, (point(1) - radius - NATIVE_CENTERLINE_EXTENSION_M) / sheetScale, 0, point(0) / sheetScale, (point(1) + radius + NATIVE_CENTERLINE_EXTENSION_M) / sheetScale, 0)
        segment.Color = DIMENSION_COLOR_BLACK
    Next item
    d.ClearSelection2 True
End Sub
