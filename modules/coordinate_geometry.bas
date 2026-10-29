Attribute VB_Name = "coordinate_geometry"
Option Explicit
Private Sub CollectViewCoordinates(ByVal v As SldWorks.View, ByRef xs As Object, ByRef ys As Object)
    Dim comps As Variant, ci As Variant, edges As Variant, ei As Variant
    Dim edge As SldWorks.Edge, curve As SldWorks.Curve, vertex As SldWorks.Vertex, util As SldWorks.MathUtility
    Dim point As SldWorks.MathPoint, cp As Variant, p As Variant, q As Variant, a(2) As Double, endpoint As Long
    Dim normal As SldWorks.MathVector, radius As Double, fullCircle As Boolean
    Set xs = CreateObject("Scripting.Dictionary"): Set ys = CreateObject("Scripting.Dictionary")
    Set util = sw.GetMathUtility
    comps = v.GetVisibleComponents
    If IsEmpty(comps) Then
        Exit Sub
    End If
    For Each ci In comps
        edges = v.GetVisibleEntities2(ci, SW_VISIBLE_ENTITY_EDGE)
        If Not IsEmpty(edges) Then
            For Each ei In edges
                Set edge = ei: Set curve = edge.GetCurve
                fullCircle = False
                If curve.IsCircle Then
                    Dim first As SldWorks.Vertex, last As SldWorks.Vertex
                    Set first = edge.GetStartVertex: Set last = edge.GetEndVertex
                    fullCircle = (first Is Nothing Or last Is Nothing)
                    If Not first Is Nothing And Not last Is Nothing Then
                        p = first.GetPoint: q = last.GetPoint
                        fullCircle = Abs(p(0) - q(0)) + Abs(p(1) - q(1)) + Abs(p(2) - q(2)) < CIRCLE_CLOSURE_TOLERANCE_M
                    End If
                    If fullCircle Then
                        cp = curve.CircleParams
                        a(0) = cp(CIRCLE_AXIS_X_INDEX): a(1) = cp(CIRCLE_AXIS_Y_INDEX): a(2) = cp(CIRCLE_AXIS_Z_INDEX)
                        Set normal = util.CreateVector(a): Set normal = normal.MultiplyTransform(v.ModelToViewTransform)
                        p = normal.ArrayData
                        If Abs(p(2)) > FACE_NORMAL_ALIGNMENT_COSINE * Sqr(p(0) ^ 2 + p(1) ^ 2 + p(2) ^ 2) Then
                            a(0) = cp(CIRCLE_CENTER_X_INDEX): a(1) = cp(CIRCLE_CENTER_Y_INDEX): a(2) = cp(CIRCLE_CENTER_Z_INDEX)
                            Set point = util.CreatePoint(a): Set point = point.MultiplyTransform(v.ModelToViewTransform)
                            p = point.ArrayData
                            AddCoordinate xs, edge, p(0), v.ScaleDecimal, p(1)
                            AddCoordinate ys, edge, p(1), v.ScaleDecimal, p(0)
                        End If
                    End If
                End If
                If Not fullCircle Then
                    For endpoint = 0 To 1
                        Set vertex = Nothing
                        If endpoint = 0 Then
                            Set vertex = edge.GetStartVertex
                        Else
                            Set vertex = edge.GetEndVertex
                        End If
                        If Not vertex Is Nothing Then
                            Set point = util.CreatePoint(vertex.GetPoint): Set point = point.MultiplyTransform(v.ModelToViewTransform)
                            p = point.ArrayData
                            AddCoordinate xs, vertex, p(0), v.ScaleDecimal, p(1)
                            AddCoordinate ys, vertex, p(1), v.ScaleDecimal, p(0)
                        End If
                    Next endpoint
                End If
            Next ei
        End If
    Next ci
End Sub

Private Function CoordinateDimensions(ByVal d As SldWorks.ModelDoc2, ByVal v As SldWorks.View, ByVal detail As Boolean, Optional ByVal layoutOnly As Boolean = False) As Long
    Dim xs As Object, ys As Object
    On Error GoTo Failed
    If coordinateSourceView Is Nothing Then
        CollectViewCoordinates v, xs, ys
    Else
        Set xs = TransformCoordinates(coordinateSourceX, coordinateSourceView, v, 0)
        Set ys = TransformCoordinates(coordinateSourceY, coordinateSourceView, v, 1)
    End If
    Dim geometryBounds As Variant, key As Variant, record As Object, xmin As Double, xmax As Double, ymin As Double, ymax As Double
    xmin = SEARCH_INITIAL_LIMIT: xmax = -SEARCH_INITIAL_LIMIT: ymin = SEARCH_INITIAL_LIMIT: ymax = -SEARCH_INITIAL_LIMIT
    For Each key In xs.Keys
        Set record = xs(key)
        xmin = MinD(xmin, CDbl(record("coordinate")))
        xmax = -MinD(-xmax, -CDbl(record("coordinate")))
    Next key
    For Each key In ys.Keys
        Set record = ys(key)
        ymin = MinD(ymin, CDbl(record("coordinate")))
        ymax = -MinD(-ymax, -CDbl(record("coordinate")))
    Next key
    geometryBounds = Array(xmin, ymin, xmax, ymax)
    If detail Then
        If coordinateSourceView Is Nothing Then
            geometryBounds = ProjectedFlatBounds(v)
        Else
            geometryBounds = v.GetOutline
        End If
    End If
    routingXMinimum = CoordinateRegionEdge(0)
    routingXMaximum = CoordinateRegionEdge(2)
    CoordinateAxis d, v, ys, 1, geometryBounds, detail, True
    If layoutOnly Then
        CoordinateAxis d, v, xs, 0, geometryBounds, detail, True
        Exit Function
    End If
    CoordinateDimensions = CoordinateAxis(d, v, xs, 0, geometryBounds, detail) + CoordinateAxis(d, v, ys, 1, geometryBounds, detail)
    Log "COORDINATES " & v.Name & " x=" & xs.Count & " y=" & ys.Count & " dimensions=" & CoordinateDimensions
    Exit Function
Failed:
    Dim savedNumber As Long, savedMessage As String
    savedNumber = Err.Number
    savedMessage = Err.Description
    Err.Raise savedNumber, "CoordinateDimensions", savedMessage
End Function

Private Sub AddCoordinate(ByVal coordinates As Object, ByVal entity As Object, ByVal value As Double, ByVal viewScale As Double, ByVal perpendicular As Double)
    Dim key As String, record As Object, candidates As Collection, positions As Collection
    key = CStr(Round(value / viewScale, COORDINATE_KEY_DECIMALS))
    If coordinates.Exists(key) Then
        Set record = coordinates(key)
        Set candidates = record("candidates")
        candidates.Add entity
        Set positions = record("candidatePositions")
        positions.Add perpendicular
        If perpendicular < CDbl(record("low")) - MIN_PROJECTED_EXTENT_M Then
            record("low") = perpendicular
            Set record("lowEntity") = entity
        End If
        If perpendicular > CDbl(record("high")) + MIN_PROJECTED_EXTENT_M Then
            record("high") = perpendicular
            Set record("highEntity") = entity
        End If
    Else
        Set record = CreateObject("Scripting.Dictionary")
        record.Add "coordinate", value
        record.Add "low", perpendicular
        record.Add "high", perpendicular
        record.Add "lowEntity", entity
        record.Add "highEntity", entity
        Set candidates = New Collection
        candidates.Add entity
        record.Add "candidates", candidates
        Set positions = New Collection
        positions.Add perpendicular
        record.Add "candidatePositions", positions
        coordinates.Add key, record
    End If
End Sub

Private Function SelectRoutedReference(ByVal v As SldWorks.View, ByVal record As Object, ByVal side As Long, ByVal append As Boolean) As Boolean
    Dim entity As Object, candidate As Variant, candidates As Collection
    If side = 0 Then
        Set entity = record("lowEntity")
    Else
        Set entity = record("highEntity")
    End If
    If v.SelectEntity(entity, append) Then
        SelectRoutedReference = True
        Exit Function
    End If
    Set candidates = record("candidates")
    For Each candidate In candidates
        If v.SelectEntity(candidate, append) Then
            SelectRoutedReference = True
            Exit Function
        End If
    Next candidate
End Function

Private Function CoordinateIntervalKey(ByVal v As SldWorks.View, ByVal axis As Long, ByVal first As Double, ByVal last As Double, ByVal detail As Boolean) As String
    Dim zero(2) As Double, util As SldWorks.MathUtility, point As SldWorks.MathPoint
    Dim origin As Variant, startValue As Double, endValue As Double, direction(2) As Double
    Dim vector As SldWorks.MathVector, values As Variant, i As Long, sign As Double
    Set util = sw.GetMathUtility
    Set point = util.CreatePoint(zero)
    Set point = point.MultiplyTransform(v.ModelToViewTransform)
    origin = point.ArrayData
    direction(axis) = 1
    Set vector = util.CreateVector(direction)
    Set vector = vector.MultiplyTransform(v.ModelToViewTransform.Inverse)
    values = vector.ArrayData
    sign = 1
    For i = 0 To UBound(direction)
        If Abs(values(i)) > DIRECTION_ZERO_TOLERANCE Then
            If values(i) < 0 Then
                sign = -1
            End If
            Exit For
        End If
    Next i
    startValue = (first - origin(axis)) / v.ScaleDecimal * sign
    endValue = (last - origin(axis)) / v.ScaleDecimal * sign
    CoordinateIntervalKey = "coordinate|" & CStr(detail) & "|" & PhysicalAxisKey(v, axis) & "|" & Format$(MinD(startValue, endValue), GEOMETRY_KEY_FORMAT) & "|" & Format$(-MinD(-startValue, -endValue), GEOMETRY_KEY_FORMAT)
End Function


