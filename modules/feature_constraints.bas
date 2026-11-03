Attribute VB_Name = "feature_constraints"
Option Explicit
Private Function FeatureHalfConstraints(ByVal d As SldWorks.ModelDoc2, ByVal symmetry As Object) As Object
    Dim intervals As Object, covered As Object, sheetName As Variant, view As SldWorks.View, display As SldWorks.DisplayDimension
    Dim dimension As SldWorks.Dimension, key As String, fields As Variant, endpoints As Variant, parts As Variant, record As Variant
    Dim name As Variant, other As Variant, full As Variant, flags As Variant, axis As Long
    Set intervals = CreateObject("Scripting.Dictionary")
    Set covered = CreateObject("Scripting.Dictionary")
    flags = symmetry("axes")
    For Each sheetName In AsDrawing(d).GetSheetNames
        AsDrawing(d).ActivateSheet CStr(sheetName)
        Set view = AsDrawing(d).GetFirstView.GetNextView
        Do While Not view Is Nothing
            Set display = view.GetFirstDisplayDimension5
            Do While Not display Is Nothing
                If display.Type2 = SW_DIMENSION_LINEAR Then
                    Set dimension = display.GetDimension2(SW_DIMENSION_INDEX_FIRST)
                    key = NativeMeasurementKey(display, view, symmetry)
                    If Left$(key, Len(AXIS_INTERVAL_KEY_PREFIX)) = AXIS_INTERVAL_KEY_PREFIX Then
                        fields = Split(key, "|")
                        axis = CLng(fields(MEASUREMENT_AXIS_FIELD))
                        If CBool(flags(axis)) And Not intervals.Exists(dimension.FullName) Then
                            endpoints = Split(CStr(fields(MEASUREMENT_INTERVAL_FIELD)), ":")
                            parts = Split(dimension.FullName, "@")
                            intervals.Add dimension.FullName, Array(CStr(parts(SOURCE_DIMENSION_FEATURE_INDEX)), axis, CDbl(endpoints(0)), CDbl(endpoints(1)))
                        End If
                    End If
                End If
                Set display = display.GetNext5
            Loop
            Set view = view.GetNextView
        Loop
    Next sheetName
    For Each name In intervals.Keys
        record = intervals(name)
        If record(CONSTRAINT_LOW_INDEX) = 0 Or record(CONSTRAINT_HIGH_INDEX) = 0 Then
            For Each other In intervals.Keys
                full = intervals(other)
                If full(CONSTRAINT_FEATURE_INDEX) = record(CONSTRAINT_FEATURE_INDEX) And full(CONSTRAINT_AXIS_INDEX) = record(CONSTRAINT_AXIS_INDEX) Then
                    If full(CONSTRAINT_LOW_INDEX) = -full(CONSTRAINT_HIGH_INDEX) And full(CONSTRAINT_HIGH_INDEX) > 0 Then
                        If Abs(record(CONSTRAINT_HIGH_INDEX) - record(CONSTRAINT_LOW_INDEX)) = full(CONSTRAINT_HIGH_INDEX) Then
                            covered.Add CStr(name), CStr(other)
                            Exit For
                        End If
                    End If
                End If
            Next other
        End If
    Next name
    Set FeatureHalfConstraints = covered
End Function

Private Sub AddCoaxialCircleSizes(ByVal d As SldWorks.ModelDoc2, ByVal covered As Object, ByVal axes As Object, ByVal symmetry As Object)
    Dim sheets As Variant, sheetIndex As Long, view As SldWorks.View, components As Variant, component As Variant
    Dim edges As Variant, item As Variant, edge As SldWorks.Edge, curve As SldWorks.Curve, parameters As Variant, matrix As Variant
    Dim first As SldWorks.Vertex, last As SldWorks.Vertex, a As Variant, b As Variant, closed As Boolean, aligned As Double
    Dim key As String, axisKey As String, outline As Variant, display As SldWorks.DisplayDimension
    sheets = AsDrawing(d).GetSheetNames
    For sheetIndex = UBound(sheets) To LBound(sheets) Step -1
        AsDrawing(d).ActivateSheet CStr(sheets(sheetIndex))
        Set view = AsDrawing(d).GetFirstView.GetNextView
        Do While Not view Is Nothing
            matrix = view.ModelToViewTransform.ArrayData
            components = view.GetVisibleComponents
            If Not IsEmpty(components) Then
                For Each component In components
                    edges = view.GetVisibleEntities2(component, SW_VISIBLE_ENTITY_EDGE)
                    If Not IsEmpty(edges) Then
                        For Each item In edges
                            Set edge = item
                            Set curve = edge.GetCurve
                            If Not curve Is Nothing Then
                                If curve.IsCircle Then
                                    parameters = curve.CircleParams
                                    aligned = Abs(parameters(CIRCLE_AXIS_X_INDEX) * matrix(MATRIX_XZ_INDEX) + parameters(CIRCLE_AXIS_Y_INDEX) * matrix(MATRIX_YZ_INDEX) + parameters(CIRCLE_AXIS_Z_INDEX) * matrix(MATRIX_ZZ_INDEX))
                                    If aligned >= FACE_NORMAL_ALIGNMENT_COSINE Then
                                        key = CircularAxisKey(parameters, parameters(CIRCLE_RADIUS_INDEX), symmetry)
                                        axisKey = Left$(key, InStrRev(key, "|"))
                                        If axes.Exists(axisKey) And Not covered.Exists(key) Then
                                            Set first = edge.GetStartVertex
                                            Set last = edge.GetEndVertex
                                            closed = first Is Nothing Or last Is Nothing
                                            If Not closed Then
                                                a = first.GetPoint
                                                b = last.GetPoint
                                                closed = Sqr((a(0) - b(0)) ^ 2 + (a(1) - b(1)) ^ 2 + (a(2) - b(2)) ^ 2) <= CIRCLE_CLOSURE_TOLERANCE_M
                                            End If
                                            If closed And Not DerivedChamferRim(edge) Then
                                                If Not SelectFeatureEdge(d, view, edge) Then
                                                    Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot select a coaxial circle."
                                                End If
                                                outline = view.GetOutline
                                                Set display = d.AddDiameterDimension2(outline(2) + DIAMETER_LABEL_OFFSET_M, outline(3), 0)
                                                If display Is Nothing Then
                                                    Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot dimension a coaxial circle."
                                                End If
                                                featureDimensionOwners.Add display.GetNameForSelection, "feature|" & key
                                                covered(key) = True
                                                Log "FEATURE COAXIAL diameter=" & CStr(display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Next item
                    End If
                Next component
            End If
            Set view = view.GetNextView
        Loop
    Next sheetIndex
End Sub

Private Function FeatureCircularQuality(ByVal key As String, ByVal view As SldWorks.View, ByVal symmetry As Object, ByVal preferReverse As Boolean) As Long
    Dim expected As String, components As Variant, component As Variant, edges As Variant, item As Variant
    Dim edge As SldWorks.Edge, curve As SldWorks.Curve, parameters As Variant, matrix As Variant, aligned As Double
    Dim first As SldWorks.Vertex, last As SldWorks.Vertex, a As Variant, b As Variant, closed As Boolean, quality As Long
    If Left$(key, Len(FEATURE_GEOMETRY_PREFIX & CIRCULAR_SIZE_KEY_PREFIX)) <> FEATURE_GEOMETRY_PREFIX & CIRCULAR_SIZE_KEY_PREFIX Then
        Exit Function
    End If
    expected = Mid$(key, Len(FEATURE_GEOMETRY_PREFIX) + 1)
    matrix = view.ModelToViewTransform.ArrayData
    components = view.GetVisibleComponents
    If IsEmpty(components) Then
        Exit Function
    End If
    For Each component In components
        edges = view.GetVisibleEntities2(component, SW_VISIBLE_ENTITY_EDGE)
        If Not IsEmpty(edges) Then
            For Each item In edges
                Set edge = item
                Set curve = edge.GetCurve
                If Not curve Is Nothing Then
                    If curve.IsCircle Then
                        parameters = curve.CircleParams
                        aligned = Abs(parameters(CIRCLE_AXIS_X_INDEX) * matrix(MATRIX_XZ_INDEX) + parameters(CIRCLE_AXIS_Y_INDEX) * matrix(MATRIX_YZ_INDEX) + parameters(CIRCLE_AXIS_Z_INDEX) * matrix(MATRIX_ZZ_INDEX))
                        If aligned >= FACE_NORMAL_ALIGNMENT_COSINE Then
                            If CircularAxisKey(parameters, parameters(CIRCLE_RADIUS_INDEX), symmetry) = expected Then
                                Set first = edge.GetStartVertex
                                Set last = edge.GetEndVertex
                                closed = first Is Nothing Or last Is Nothing
                                If Not closed Then
                                    a = first.GetPoint
                                    b = last.GetPoint
                                    closed = Sqr((a(0) - b(0)) ^ 2 + (a(1) - b(1)) ^ 2 + (a(2) - b(2)) ^ 2) <= CIRCLE_CLOSURE_TOLERANCE_M
                                End If
                                quality = FEATURE_VISIBLE_ARC_PRIORITY
                                If closed Then
                                    quality = FEATURE_COMPLETE_CIRCLE_PRIORITY
                                End If
                                If CircleIsBore(edge) And Not preferReverse Then
                                    quality = quality + FEATURE_PROFILE_FRONT_PRIORITY
                                ElseIf preferReverse Then
                                    quality = quality + FEATURE_REVERSE_CIRCLE_BONUS
                                End If
                                If quality > FeatureCircularQuality Then
                                    FeatureCircularQuality = quality
                                End If
                            End If
                        End If
                    End If
                End If
            Next item
        End If
    Next component
End Function
