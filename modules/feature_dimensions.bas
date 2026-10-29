Attribute VB_Name = "feature_dimensions"
Option Explicit
Private Function PhysicalAxisKey(ByVal v As SldWorks.View, ByVal axis As Long) As String
    Dim direction(2) As Double, util As SldWorks.MathUtility, vector As SldWorks.MathVector
    Dim values As Variant, length As Double, i As Long, orientation As Double
    direction(axis) = 1
    Set util = sw.GetMathUtility
    Set vector = util.CreateVector(direction)
    Set vector = vector.MultiplyTransform(v.ModelToViewTransform.Inverse)
    values = vector.ArrayData
    length = Sqr(values(0) ^ 2 + values(1) ^ 2 + values(2) ^ 2)
    orientation = 1
    For i = 0 To UBound(direction)
        If Abs(values(i) / length) > DIRECTION_ZERO_TOLERANCE Then
            If values(i) < 0 Then
                orientation = -1
            End If
            Exit For
        End If
    Next i
    For i = 0 To UBound(direction)
        PhysicalAxisKey = PhysicalAxisKey & "|" & CStr(Round(values(i) / length * orientation, DIRECTION_KEY_DECIMALS))
    Next i
End Function

Private Function SortedCoordinateKeys(ByVal coordinates As Object) As Variant
    Dim keys As Variant, i As Long, j As Long, temp As Variant
    keys = coordinates.Keys
    For i = LBound(keys) + 1 To UBound(keys)
        temp = keys(i)
        j = i - 1
        Do While j >= LBound(keys)
            If CDbl(keys(j)) <= CDbl(temp) Then
                Exit Do
            End If
            keys(j + 1) = keys(j)
            j = j - 1
        Loop
        keys(j + 1) = temp
    Next i
    SortedCoordinateKeys = keys
End Function

Private Sub DimensionFeatures(ByVal d As SldWorks.ModelDoc2, ByVal v As SldWorks.View, ByVal detail As Boolean, Optional ByVal includeCoordinates As Boolean = True)
    Dim count As Long
    If Setting("AutoDimensions", "1") <> "1" Then
        Exit Sub
    End If
    If includeCoordinates Then
        count = CoordinateDimensions(d, v, detail)
    End If
    count = DiameterDimensions(d, v)
    count = RadiusDimensions(d, v)
End Sub

Private Function RadiusDimensions(ByVal d As SldWorks.ModelDoc2, ByVal v As SldWorks.View) As Long
    Dim comps As Variant, edges As Variant, ci As Variant, ei As Variant, edge As SldWorks.Edge
    Dim curve As SldWorks.Curve, cp As Variant, first As SldWorks.Vertex, last As SldWorks.Vertex
    Dim p As Variant, q As Variant, seen As Object, key As String, outline As Variant
    Dim util As SldWorks.MathUtility, vector As SldWorks.MathVector, axis(2) As Double, dd As SldWorks.DisplayDimension
    On Error GoTo Failed
    Set seen = CreateObject("Scripting.Dictionary"): Set util = sw.GetMathUtility
    comps = v.GetVisibleComponents: outline = v.GetOutline
    If IsEmpty(comps) Then
        Exit Function
    End If
    For Each ci In comps
        edges = v.GetVisibleEntities2(ci, SW_VISIBLE_ENTITY_EDGE)
        If Not IsEmpty(edges) Then
            For Each ei In edges
                Set edge = ei: Set curve = edge.GetCurve
                If curve.IsCircle Then
                    Set first = edge.GetStartVertex: Set last = edge.GetEndVertex
                    If first Is Nothing Or last Is Nothing Then
                        GoTo NextArc
                    End If
                    p = first.GetPoint: q = last.GetPoint
                    If Abs(p(0) - q(0)) + Abs(p(1) - q(1)) + Abs(p(2) - q(2)) < CIRCLE_CLOSURE_TOLERANCE_M Then
                        GoTo NextArc
                    End If
                    cp = curve.CircleParams
                    axis(0) = cp(CIRCLE_AXIS_X_INDEX): axis(1) = cp(CIRCLE_AXIS_Y_INDEX): axis(2) = cp(CIRCLE_AXIS_Z_INDEX)
                    Set vector = util.CreateVector(axis): Set vector = vector.MultiplyTransform(v.ModelToViewTransform)
                    p = vector.ArrayData
                    If Abs(p(2)) < FACE_NORMAL_ALIGNMENT_COSINE * Sqr(p(0) ^ 2 + p(1) ^ 2 + p(2) ^ 2) Then
                        GoTo NextArc
                    End If
                    key = "radius|" & PhysicalAxisKey(v, 2) & "|" & Format$(cp(CIRCLE_RADIUS_INDEX), GEOMETRY_KEY_FORMAT)
                    If seen.Exists(key) Or dimensionKeys.Exists(key) Then
                        GoTo NextArc
                    End If
                    seen.Add key, True
                    AsDrawing(d).ActivateView "": d.ClearSelection2 True
                    If v.SelectEntity(edge, False) Then
                        Set dd = d.AddRadialDimension2(outline(OUTLINE_RIGHT_INDEX) + RADIUS_LABEL_OFFSET_M, outline(OUTLINE_BOTTOM_INDEX) + RADIUS_LABEL_PITCH_M * RadiusDimensions, 0)
                        If Not dd Is Nothing Then
                            If ValidateDimension(d, v, dd, cp(CIRCLE_RADIUS_INDEX) * v.ScaleDecimal) Then
                                dimensionKeys.Add key, True
                                RadiusDimensions = RadiusDimensions + 1
                            End If
                        End If
                    End If
                End If
NextArc:
            Next ei
        End If
    Next ci
    Exit Function
Failed:
    Warn "Radius dimensions " & v.Name & ": " & Err.Description
End Function

