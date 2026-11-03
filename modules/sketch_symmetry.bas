Attribute VB_Name = "sketch_symmetry"
Option Explicit
Private Function QuarterTurnPoint(ByVal point As Variant, ByVal center As Variant) As Variant
    QuarterTurnPoint = Array(center(0) - point(1) + center(1), center(1) + point(0) - center(0), point(2))
End Function

Private Function SourceSketchPoint(ByVal point As SldWorks.SketchPoint) As Variant
    SourceSketchPoint = Array(point.X, point.Y, point.Z)
End Function

Private Function QuarterSketchSegmentKey(ByVal segment As SldWorks.SketchSegment, ByVal center As Variant, ByVal turn As Boolean) As String
    Dim line As SldWorks.SketchLine, arc As SldWorks.SketchArc, a As Variant, b As Variant, c As Variant, key As String
    If segment.GetType = SW_SKETCH_LINE Then
        Set line = segment
        a = SourceSketchPoint(line.GetStartPoint2)
        b = SourceSketchPoint(line.GetEndPoint2)
        If turn Then
            a = QuarterTurnPoint(a, center)
            b = QuarterTurnPoint(b, center)
        End If
        key = "line|" & NativePairKey(a, b)
    ElseIf segment.GetType = SW_SKETCH_ARC Then
        Set arc = segment
        c = SourceSketchPoint(arc.GetCenterPoint2)
        a = SourceSketchPoint(arc.GetStartPoint2)
        b = SourceSketchPoint(arc.GetEndPoint2)
        If turn Then
            a = QuarterTurnPoint(a, center)
            b = QuarterTurnPoint(b, center)
            c = QuarterTurnPoint(c, center)
        End If
        key = "arc|" & NativePointKey(c) & "|" & CStr(Round(arc.GetRadius / NATIVE_GEOMETRY_KEY_TOLERANCE_M, 0)) & "|" & CStr(Round(segment.GetLength / NATIVE_GEOMETRY_KEY_TOLERANCE_M, 0))
        If arc.IsCircle = 0 Then
            key = key & "|" & NativePairKey(a, b)
        End If
    Else
        Exit Function
    End If
    QuarterSketchSegmentKey = key & "|" & CStr(segment.ConstructionGeometry)
End Function

Private Function SketchQuarterEvidence(ByVal feature As SldWorks.Feature) As Object
    Dim result As Object, sketch As SldWorks.Sketch, segments As Variant, item As Variant, segment As SldWorks.SketchSegment
    Dim arc As SldWorks.SketchArc, line As SldWorks.SketchLine, a As Variant, b As Variant, c As Variant, radius As Double
    Dim bounds As Variant, center As Variant, keys As Object, key As String, axis As Long, valid As Boolean
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "valid", False
    If feature Is Nothing Then
        Set SketchQuarterEvidence = result
        Exit Function
    End If
    If feature.GetTypeName2 <> FEATURE_TYPE_SKETCH Then
        Set SketchQuarterEvidence = result
        Exit Function
    End If
    Set sketch = feature.GetSpecificFeature2
    segments = sketch.GetSketchSegments
    If IsEmpty(segments) Then
        Set SketchQuarterEvidence = result
        Exit Function
    End If
    bounds = Array(SEARCH_INITIAL_LIMIT, SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT)
    For Each item In segments
        Set segment = item
        If segment.ConstructionGeometry Then
            GoTo NextMaterialBounds
        End If
        If segment.GetType = SW_SKETCH_LINE Then
            Set line = segment
            a = SourceSketchPoint(line.GetStartPoint2)
            b = SourceSketchPoint(line.GetEndPoint2)
        ElseIf segment.GetType = SW_SKETCH_ARC Then
            Set arc = segment
            c = SourceSketchPoint(arc.GetCenterPoint2)
            radius = arc.GetRadius
            a = Array(c(0) - radius, c(1) - radius, c(2))
            b = Array(c(0) + radius, c(1) + radius, c(2))
        Else
            Set SketchQuarterEvidence = result
            Exit Function
        End If
        NativeExpandBounds bounds, a(0), a(1)
        NativeExpandBounds bounds, b(0), b(1)
NextMaterialBounds:
    Next item
    center = Array((bounds(0) + bounds(2)) / 2, (bounds(1) + bounds(3)) / 2, 0)
    Set keys = CreateObject("Scripting.Dictionary")
    For Each item In segments
        Set segment = item
        If segment.ConstructionGeometry Then
            GoTo NextMaterialKey
        End If
        key = QuarterSketchSegmentKey(segment, center, False)
        keys(key) = True
NextMaterialKey:
    Next item
    valid = keys.Count > 0
    For Each item In segments
        Set segment = item
        If segment.ConstructionGeometry Then
            GoTo NextMaterialRotation
        End If
        key = QuarterSketchSegmentKey(segment, center, True)
        If Not keys.Exists(key) Then
            valid = False
            Exit For
        End If
NextMaterialRotation:
    Next item
    result("valid") = valid
    result.Add "center", center
    result.Add "transform", sketch.ModelToSketchTransform.ArrayData
    Set SketchQuarterEvidence = result
    Log "FEATURE SKETCH rotation90=" & CStr(valid) & "; source=" & feature.Name
End Function

Private Function SourcePointInSketch(ByVal point As Variant, ByVal transform As Variant) As Variant
    SourcePointInSketch = Array((point(0) * transform(0) + point(1) * transform(3) + point(2) * transform(6)) * transform(MATRIX_SCALE_INDEX) + transform(9), (point(0) * transform(1) + point(1) * transform(4) + point(2) * transform(7)) * transform(MATRIX_SCALE_INDEX) + transform(10), 0)
End Function

Private Function SketchRotationKey(ByVal display As SldWorks.DisplayDimension, ByVal model As SldWorks.ModelDoc2, ByVal source As String, ByVal cache As Object) As String
    Dim evidence As Object, feature As SldWorks.Feature, dimension As SldWorks.Dimension, refs As Variant
    Dim first As SldWorks.MathPoint, last As SldWorks.MathPoint, a As Variant, b As Variant, center As Variant
    Dim key As String, candidate As String, turn As Long
    If display.Type2 <> SW_DIMENSION_LINEAR Then
        Exit Function
    End If
    If Not cache.Exists(source) Then
        Set feature = model.FeatureByName(source)
        Set evidence = SketchQuarterEvidence(feature)
        cache.Add source, evidence
    End If
    Set evidence = cache(source)
    ' Individual dimension references can repeat even when the complete sketch is asymmetric.
    If Not evidence.Exists("center") Then
        Exit Function
    End If
    Set dimension = display.GetDimension2(SW_DIMENSION_INDEX_FIRST)
    refs = dimension.ReferencePoints
    If IsEmpty(refs) Then
        Exit Function
    End If
    If UBound(refs) < DIMENSION_REFERENCE_SECOND - 1 Then
        Exit Function
    End If
    Set first = refs(DIMENSION_REFERENCE_FIRST - 1)
    Set last = refs(DIMENSION_REFERENCE_SECOND - 1)
    a = SourcePointInSketch(first.ArrayData, evidence("transform"))
    b = SourcePointInSketch(last.ArrayData, evidence("transform"))
    center = evidence("center")
    key = NativePairKey(a, b)
    For turn = 1 To QUARTER_TURN_EQUIVALENT_COUNT - 1
        a = QuarterTurnPoint(a, center)
        b = QuarterTurnPoint(b, center)
        candidate = NativePairKey(a, b)
        If candidate < key Then
            key = candidate
        End If
    Next turn
    SketchRotationKey = source & "|" & key & "|" & CStr(Round(dimension.SystemValue / NATIVE_GEOMETRY_KEY_TOLERANCE_M, 0))
End Function
