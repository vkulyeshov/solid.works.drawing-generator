Attribute VB_Name = "symmetry"
Option Explicit
Private Function NativePointKey(ByVal point As Variant) As String
    Dim axis As Long, coordinate As Double
    For axis = 0 To 2
        coordinate = Round(point(axis) / NATIVE_GEOMETRY_KEY_TOLERANCE_M, 0)
        If coordinate = 0 Then
            coordinate = 0
        End If
        NativePointKey = NativePointKey & "|" & CStr(coordinate)
    Next axis
End Function

Private Function NativeReflectedPoint(ByVal point As Variant, ByVal axis As Long, ByVal center As Variant) As Variant
    Dim result As Variant
    result = Array(point(0), point(1), point(2))
    If axis >= 0 Then
        result(axis) = 2 * center(axis) - result(axis)
    End If
    NativeReflectedPoint = result
End Function

Private Function NativePairKey(ByVal first As Variant, ByVal last As Variant) As String
    Dim a As String, b As String
    a = NativePointKey(first)
    b = NativePointKey(last)
    If a < b Then
        NativePairKey = a & ";" & b
    Else
        NativePairKey = b & ";" & a
    End If
End Function

Private Function NativeDirectionKey(ByVal direction As Variant) As String
    Dim axis As Long, sign As Double, result As Variant
    result = Array(direction(0), direction(1), direction(2))
    sign = 1
    For axis = 0 To 2
        If Abs(result(axis)) > DIRECTION_ZERO_TOLERANCE Then
            If result(axis) < 0 Then
                sign = -1
            End If
            Exit For
        End If
    Next axis
    For axis = 0 To 2
        result(axis) = result(axis) * sign
    Next axis
    NativeDirectionKey = NativePointKey(result)
End Function

Private Function NativeEdgeSymmetryKey(ByVal edge As SldWorks.Edge, ByVal axis As Long, ByVal center As Variant) As String
    Dim curve As SldWorks.Curve, first As SldWorks.Vertex, last As SldWorks.Vertex
    Dim a As Variant, b As Variant, circleData As Variant, origin As Variant, direction As Variant
    Dim parameters As SldWorks.CurveParamData, evaluated As Variant, samples() As String, index As Long, other As Long, temporary As String
    Set curve = edge.GetCurve
    Set first = edge.GetStartVertex
    Set last = edge.GetEndVertex
    If Not first Is Nothing And Not last Is Nothing Then
        a = NativeReflectedPoint(first.GetPoint, axis, center)
        b = NativeReflectedPoint(last.GetPoint, axis, center)
        NativeEdgeSymmetryKey = NativePairKey(a, b)
    End If
    If curve.IsLine Then
        NativeEdgeSymmetryKey = "L" & NativeEdgeSymmetryKey
    ElseIf curve.IsCircle Then
        circleData = curve.CircleParams
        origin = NativeReflectedPoint(circleData, axis, center)
        direction = Array(circleData(3), circleData(4), circleData(5))
        If axis >= 0 Then
            direction(axis) = -direction(axis)
        End If
        NativeEdgeSymmetryKey = "C" & NativePointKey(origin) & NativeDirectionKey(direction) & "|" & Round(circleData(CIRCLE_RADIUS_INDEX) / NATIVE_GEOMETRY_KEY_TOLERANCE_M, 0) & ";" & NativeEdgeSymmetryKey
    Else
        Set parameters = edge.GetCurveParams3
        ReDim samples(1 To NATIVE_SYMMETRY_EDGE_SAMPLES)
        For index = 1 To NATIVE_SYMMETRY_EDGE_SAMPLES
            evaluated = edge.Evaluate2(parameters.UMinValue + (parameters.UMaxValue - parameters.UMinValue) * index / (NATIVE_SYMMETRY_EDGE_SAMPLES + 1), 0)
            If IsEmpty(evaluated) Then
                NativeEdgeSymmetryKey = ""
                Exit Function
            End If
            If evaluated(UBound(evaluated)) = 0 Then
                NativeEdgeSymmetryKey = ""
                Exit Function
            End If
            samples(index) = NativePointKey(NativeReflectedPoint(evaluated, axis, center))
        Next index
        For index = 1 To NATIVE_SYMMETRY_EDGE_SAMPLES - 1
            For other = index + 1 To NATIVE_SYMMETRY_EDGE_SAMPLES
                If samples(other) < samples(index) Then
                    temporary = samples(index)
                    samples(index) = samples(other)
                    samples(other) = temporary
                End If
            Next other
        Next index
        NativeEdgeSymmetryKey = "S" & curve.Identity & ";" & NativeEdgeSymmetryKey & ";" & Join(samples, ";")
    End If
End Function

Private Function NativeMirroredEdgeMatches(ByVal edge As SldWorks.Edge, ByVal axis As Long, ByVal center As Variant, ByVal edges As Collection) As Boolean
    Dim first As SldWorks.Vertex, last As SldWorks.Vertex, candidate As SldWorks.Edge, item As Variant
    Dim sourceKey As String, candidateKey As String, candidates As New Collection, curve As SldWorks.Curve
    Dim parameters As SldWorks.CurveParamData, point As Variant, nearest As Variant, index As Long, matched As Boolean, distance As Double
    Set first = edge.GetStartVertex
    Set last = edge.GetEndVertex
    If Not first Is Nothing And Not last Is Nothing Then
        sourceKey = NativePairKey(NativeReflectedPoint(first.GetPoint, axis, center), NativeReflectedPoint(last.GetPoint, axis, center))
    End If
    For Each item In edges
        Set candidate = item
        candidateKey = ""
        Set first = candidate.GetStartVertex
        Set last = candidate.GetEndVertex
        If Not first Is Nothing And Not last Is Nothing Then
            candidateKey = NativePairKey(first.GetPoint, last.GetPoint)
        End If
        If Len(sourceKey) = 0 Or sourceKey = candidateKey Then
            candidates.Add candidate
        End If
    Next item
    If candidates.Count = 0 Then
        Exit Function
    End If
    Set curve = edge.GetCurve
    Set parameters = edge.GetCurveParams3
    For index = 1 To NATIVE_SYMMETRY_EDGE_SAMPLES
        point = edge.Evaluate2(parameters.UMinValue + (parameters.UMaxValue - parameters.UMinValue) * index / (NATIVE_SYMMETRY_EDGE_SAMPLES + 1), 0)
        If IsEmpty(point) Then
            Exit Function
        End If
        If point(UBound(point)) = 0 Then
            Exit Function
        End If
        point = NativeReflectedPoint(point, axis, center)
        matched = False
        For Each item In candidates
            Set candidate = item
            nearest = candidate.GetClosestPointOn(point(0), point(1), point(2))
            If Not IsEmpty(nearest) Then
                distance = (nearest(0) - point(0)) ^ 2 + (nearest(1) - point(1)) ^ 2 + (nearest(2) - point(2)) ^ 2
                If distance <= NATIVE_GEOMETRY_KEY_TOLERANCE_M ^ 2 Then
                    matched = True
                    Exit For
                End If
            End If
        Next item
        If Not matched Then
            Exit Function
        End If
    Next index
    NativeMirroredEdgeMatches = True
End Function

Private Function NativeSymmetryEvidence(ByVal v As SldWorks.View) As Object
    Dim result As Object, vertices As New Collection, edges As New Collection, vertexKeys As Object, edgeKeys As Object
    Dim bodies As Variant, body As Variant, items As Variant, item As Variant, point As Variant, center As Variant, flags As Variant
    Dim model As SldWorks.ModelDoc2, vertex As SldWorks.Vertex, edge As SldWorks.Edge, key As String, bounds As Variant
    Dim axis As Long, valid As Boolean
    Dim cylinders As New Collection, circles As New Collection, face As SldWorks.Face2, surface As SldWorks.Surface, curve As SldWorks.Curve
    Set result = CreateObject("Scripting.Dictionary")
    Set vertexKeys = CreateObject("Scripting.Dictionary")
    Set edgeKeys = CreateObject("Scripting.Dictionary")
    bounds = Array(SEARCH_INITIAL_LIMIT, SEARCH_INITIAL_LIMIT, SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT)
    bodies = v.Bodies
    If IsEmpty(bodies) Then
        Set model = v.ReferencedDocument
        bodies = AsPart(model).GetBodies2(SW_SOLID_BODY, True)
    End If
    For Each body In bodies
        items = body.GetVertices
        If Not IsEmpty(items) Then
            For Each item In items
                Set vertex = item
                point = vertex.GetPoint
                vertices.Add point
                vertexKeys(NativePointKey(point)) = True
                For axis = 0 To 2
                    bounds(axis) = MinD(bounds(axis), point(axis))
                    bounds(axis + 3) = -MinD(-bounds(axis + 3), -point(axis))
                Next axis
            Next item
        End If
        items = body.GetEdges
        For Each item In items
            edges.Add item
        Next item
        items = body.GetFaces
        For Each item In items
            Set face = item
            Set surface = face.GetSurface
            If surface.IsCylinder Then
                cylinders.Add surface.CylinderParams
            End If
        Next item
    Next body
    center = Array((bounds(0) + bounds(3)) / 2, (bounds(1) + bounds(4)) / 2, (bounds(2) + bounds(5)) / 2)
    flags = Array(False, False, False)
    For Each item In edges
        Set edge = item
        Set curve = edge.GetCurve
        If curve.IsCircle Then
            circles.Add curve.CircleParams
        End If
        key = NativeEdgeSymmetryKey(edge, -1, center)
        If Len(key) > 0 Then
            edgeKeys(key) = True
        End If
    Next item
    For axis = 0 To 2
        valid = vertices.Count >= NATIVE_SYMMETRY_MIN_VERTICES And bounds(axis + 3) - bounds(axis) > NATIVE_GEOMETRY_KEY_TOLERANCE_M
        For Each item In vertices
            If Not vertexKeys.Exists(NativePointKey(NativeReflectedPoint(item, axis, center))) Then
                valid = False
                Exit For
            End If
        Next item
        Log "NATIVE SYMMETRY axis=" & axis & "; vertex match=" & CStr(valid)
        If valid Then
            For Each item In edges
                Set edge = item
                key = NativeEdgeSymmetryKey(edge, axis, center)
                If Len(key) = 0 Or Not edgeKeys.Exists(key) Then
                    If Not NativeMirroredEdgeMatches(edge, axis, center, edges) Then
                        Log "NATIVE SYMMETRY edge mismatch axis=" & axis & "; key=" & key
                        valid = False
                        Exit For
                    End If
                End If
            Next item
        End If
        flags(axis) = valid
    Next axis
    result.Add "center", center
    result.Add "axes", flags
    result.Add "cylinders", cylinders
    result.Add "circles", circles
    Set NativeSymmetryEvidence = result
    Log "NATIVE SYMMETRY verified axes=" & CStr(flags(0)) & "," & CStr(flags(1)) & "," & CStr(flags(2)) & "; vertices=" & vertices.Count & "; edges=" & edges.Count
End Function

Private Function NativeIntervalKey(ByVal first As Double, ByVal last As Double, ByVal center As Double, ByVal symmetric As Boolean) As String
    Dim a As Double, b As Double, reflected As Double
    a = Round((MinD(first, last) - center) / NATIVE_GEOMETRY_KEY_TOLERANCE_M, 0)
    b = Round((-MinD(-first, -last) - center) / NATIVE_GEOMETRY_KEY_TOLERANCE_M, 0)
    If symmetric Then
        If -b < a Then
            reflected = -b
            b = -a
            a = reflected
        End If
    End If
    If a = 0 Then
        a = 0
    End If
    If b = 0 Then
        b = 0
    End If
    NativeIntervalKey = CStr(a) & ":" & CStr(b)
End Function

Private Function NativeCylinderDiameterKey(ByVal center As Variant, ByVal diameter As Double, ByVal symmetry As Object) As String
    Dim cylinder As Variant, cylinders As Collection, candidate As String, axisKey As String, circles As Collection
    Dim dx As Double, dy As Double, dz As Double, station As Double, radial As Double, referenceStation As Double
    Dim axis As Long, sign As Double, origin As Variant, originStation As Double
    Set circles = symmetry("circles")
    For Each cylinder In circles
        If Abs(cylinder(6) * 2 - diameter) <= NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
            radial = (center(0) - cylinder(0)) ^ 2 + (center(1) - cylinder(1)) ^ 2 + (center(2) - cylinder(2)) ^ 2
            If radial <= NATIVE_GEOMETRY_KEY_TOLERANCE_M ^ 2 Then
                candidate = "circular-diameter" & NativePointKey(center) & NativeDirectionKey(Array(cylinder(3), cylinder(4), cylinder(5))) & "|" & Round(diameter / NATIVE_GEOMETRY_KEY_TOLERANCE_M, 0)
                If Len(NativeCylinderDiameterKey) > 0 And candidate <> NativeCylinderDiameterKey Then
                    NativeCylinderDiameterKey = ""
                    Exit Function
                End If
                NativeCylinderDiameterKey = candidate
            End If
        End If
    Next cylinder
    If Len(NativeCylinderDiameterKey) > 0 Then
        Exit Function
    End If
    Set cylinders = symmetry("cylinders")
    For Each cylinder In cylinders
        If Abs(cylinder(6) * 2 - diameter) <= NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
            dx = center(0) - cylinder(0)
            dy = center(1) - cylinder(1)
            dz = center(2) - cylinder(2)
            station = dx * cylinder(3) + dy * cylinder(4) + dz * cylinder(5)
            radial = dx * dx + dy * dy + dz * dz - station * station
            If radial <= NATIVE_GEOMETRY_KEY_TOLERANCE_M ^ 2 Then
                axisKey = NativeDirectionKey(Array(cylinder(3), cylinder(4), cylinder(5)))
                sign = 1
                For axis = 0 To 2
                    If Abs(cylinder(axis + 3)) > DIRECTION_ZERO_TOLERANCE Then
                        If cylinder(axis + 3) < 0 Then
                            sign = -1
                        End If
                        Exit For
                    End If
                Next axis
                referenceStation = (center(0) * cylinder(3) + center(1) * cylinder(4) + center(2) * cylinder(5)) * sign
                originStation = cylinder(0) * cylinder(3) + cylinder(1) * cylinder(4) + cylinder(2) * cylinder(5)
                origin = Array(cylinder(0) - originStation * cylinder(3), cylinder(1) - originStation * cylinder(4), cylinder(2) - originStation * cylinder(5))
                candidate = "cylindrical-diameter" & axisKey & NativePointKey(origin) & "|" & Round(referenceStation / NATIVE_GEOMETRY_KEY_TOLERANCE_M, 0) & "|" & Round(diameter / NATIVE_GEOMETRY_KEY_TOLERANCE_M, 0)
                If Len(NativeCylinderDiameterKey) > 0 And candidate <> NativeCylinderDiameterKey Then
                    NativeCylinderDiameterKey = ""
                    Exit Function
                End If
                NativeCylinderDiameterKey = candidate
            End If
        End If
    Next cylinder
End Function

Private Function NativeMeasurementKey(ByVal display As SldWorks.DisplayDimension, ByVal v As SldWorks.View, ByVal symmetry As Object) As String
    Dim dimension As SldWorks.Dimension, refs As Variant, first As SldWorks.MathPoint, last As SldWorks.MathPoint
    Dim a As Variant, b As Variant, center As Variant, flags As Variant, data As SldWorks.DisplayData, arrow As Variant, arrow2 As Variant
    Dim transform As Variant, values As Variant, axis As Long, screenAxis As Long
    Dim prefix As String, suffix As String, key As String, reflected As String, point As Variant, i As Long
    Set dimension = display.GetDimension2(SW_DIMENSION_INDEX_FIRST)
    refs = dimension.ReferencePoints
    If IsEmpty(refs) Then
        Exit Function
    End If
    If UBound(refs) < 1 Then
        Exit Function
    End If
    Set first = refs(0)
    Set last = refs(1)
    a = first.ArrayData
    b = last.ArrayData
    center = symmetry("center")
    flags = symmetry("axes")
    prefix = display.GetText(SW_DIMENSION_TEXT_PREFIX)
    suffix = display.GetText(SW_DIMENSION_TEXT_SUFFIX)
    Set data = display.GetDisplayData
    For i = 0 To data.GetTextCount - 1
        If InStr(data.GetTextAtIndex(i), SW_DIMENSION_TOKEN_DIAMETER) > 0 Then
            If display.Type2 = SW_DIMENSION_DIAMETER Then
                Exit Function
            End If
            point = Array((a(0) + b(0)) / 2, (a(1) + b(1)) / 2, (a(2) + b(2)) / 2)
            key = NativeCylinderDiameterKey(point, dimension.SystemValue, symmetry)
            If Len(key) > 0 Then
                NativeMeasurementKey = key & "|" & prefix & "|" & suffix
            End If
            Exit Function
        End If
    Next i
    If display.Type2 <> SW_DIMENSION_LINEAR Then
        Exit Function
    End If
    If data.GetArrowHeadCount = 2 Then
        arrow = data.GetArrowHeadAtIndex2(0)
        arrow2 = data.GetArrowHeadAtIndex2(1)
        screenAxis = -1
        If Abs(arrow(1) - arrow2(1)) < DIRECTION_ZERO_TOLERANCE Then
            screenAxis = 0
        ElseIf Abs(arrow(0) - arrow2(0)) < DIRECTION_ZERO_TOLERANCE Then
            screenAxis = 1
        End If
        If screenAxis >= 0 Then
            transform = v.ModelToViewTransform.ArrayData
            values = Array(transform(screenAxis), transform(screenAxis + 3), transform(screenAxis + 6))
            axis = 0
            If Abs(values(1)) > Abs(values(axis)) Then
                axis = 1
            End If
            If Abs(values(2)) > Abs(values(axis)) Then
                axis = 2
            End If
            If Abs(values(axis)) < MIN_PROJECTED_EXTENT_M Then
                Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Invalid native measurement direction."
            End If
            If Abs(a(axis) - b(axis)) < NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
                Exit Function
            End If
            NativeMeasurementKey = "axis-interval|" & axis & "|" & NativeIntervalKey(a(axis), b(axis), center(axis), CBool(flags(axis))) & "|" & Round(dimension.SystemValue / NATIVE_GEOMETRY_KEY_TOLERANCE_M, 0) & "|" & prefix & "|" & suffix
            Exit Function
        End If
    End If
    key = NativePairKey(a, b)
    For axis = 0 To 2
        If CBool(flags(axis)) Then
            reflected = NativePairKey(NativeReflectedPoint(a, axis, center), NativeReflectedPoint(b, axis, center))
            If reflected < key Then
                key = reflected
            End If
        End If
    Next axis
    transform = v.ModelToViewTransform.ArrayData
    NativeMeasurementKey = "symmetric-references" & NativeDirectionKey(Array(transform(MATRIX_XZ_INDEX), transform(MATRIX_YZ_INDEX), transform(MATRIX_ZZ_INDEX))) & "|" & key & "|" & Round(dimension.SystemValue / NATIVE_GEOMETRY_KEY_TOLERANCE_M, 0) & "|" & prefix & "|" & suffix
End Function

Private Sub DeduplicateNativeDimensions(ByVal d As SldWorks.ModelDoc2, ByVal symmetry As Object)
    Dim seen As Object, v As SldWorks.View, display As SldWorks.DisplayDimension, following As SldWorks.DisplayDimension
    Dim key As String, kept As Variant, value As Double, audit As Object, ann As SldWorks.Annotation, removed As Long
    Set seen = CreateObject("Scripting.Dictionary")
    Set v = AsDrawing(d).GetFirstView.GetNextView
    Do While Not v Is Nothing
        Set display = v.GetFirstDisplayDimension5
        Do While Not display Is Nothing
            Set following = display.GetNext5
            key = NativeMeasurementKey(display, v, symmetry)
            value = display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue
            If Len(key) > 0 Then
                If seen.Exists(key) Then
                    kept = seen(key)
                    If Abs(value - CDbl(kept(1))) <= LEADER_VALUE_TOLERANCE_M Then
                        Set audit = fs.OpenTextFile(runDir & NATIVE_AUDIT_FILENAME, FILE_FOR_APPENDING, True, FILE_UNICODE)
                        audit.WriteLine CStr(jobNumber) & vbTab & v.Name & vbTab & display.GetNameForSelection & vbTab & CStr(kept(0)) & vbTab & key
                        audit.Close
                        Set ann = display.GetAnnotation
                        d.ClearSelection2 True
                        If Not ann.Select3(False, Nothing) Then
                            Err.Raise vbObjectError + ERR_REMOVE_DIMENSION, , "Cannot select a duplicate native dimension."
                        End If
                        d.EditDelete
                        removed = removed + 1
                    End If
                Else
                    seen.Add key, Array(display.GetNameForSelection, value)
                End If
            End If
            Set display = following
        Loop
        Set v = v.GetNextView
    Loop
    d.ClearSelection2 True
    Log "NATIVE DEDUP removed=" & removed & "; matching model intervals and verified reflection geometry"
End Sub

