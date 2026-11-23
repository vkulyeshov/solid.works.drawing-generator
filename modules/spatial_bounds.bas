Attribute VB_Name = "spatial_bounds"
Option Explicit

' Closed bounds deliberately retain touching and tolerance-sized gaps. Only
' definitely separated graphics bypass the original exact predicates.
Private Function SpatialBoundsOverlap(ByVal a As Variant, ByVal b As Variant) As Boolean
    SpatialBoundsOverlap = a(0) <= b(2) + PROFILE_MATERIAL_TOLERANCE_M And a(2) >= b(0) - PROFILE_MATERIAL_TOLERANCE_M And a(1) <= b(3) + PROFILE_MATERIAL_TOLERANCE_M And a(3) >= b(1) - PROFILE_MATERIAL_TOLERANCE_M
End Function

Private Function SpatialLineBounds(ByVal line As Variant) As Variant
    SpatialLineBounds = Array(MinD(line(0), line(2)), MinD(line(1), line(3)), MaxD(line(0), line(2)), MaxD(line(1), line(3)))
End Function

Private Function SpatialLineBoxes(ByVal lines As Collection) As Collection
    Dim result As New Collection, line As Variant
    For Each line In lines
        result.Add SpatialLineBounds(line)
    Next line
    Set SpatialLineBoxes = result
End Function

Private Function SpatialPolygonBoxes(ByVal polygons As Collection) As Collection
    Dim result As New Collection, polygon As Variant
    For Each polygon In polygons
        result.Add PanelMaterialBounds(polygon)
    Next polygon
    Set SpatialPolygonBoxes = result
End Function

' Copies inherit dictionary entries, so every geometry-changing operation
' must discard inherited bounds. Cache contents are immutable afterwards.
Private Sub SpatialClear(ByVal candidate As Object)
    If candidate.Exists("spatial") Then
        candidate.Remove "spatial"
    End If
End Sub

Private Function SpatialCandidate(ByVal candidate As Object) As Object
    Dim result As Object, innerHeads As New Collection, triangle As Variant
    If candidate.Exists("spatial") Then
        Set SpatialCandidate = candidate("spatial")
        Exit Function
    End If
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "text", SpatialPolygonBoxes(candidate("text"))
    result.Add "lines", SpatialLineBoxes(candidate("lines"))
    result.Add "bases", SpatialLineBoxes(candidate("bases"))
    For Each triangle In candidate("heads")
        innerHeads.Add PlanarShrinkTriangle(triangle)
    Next triangle
    result.Add "heads", innerHeads
    result.Add "headBoxes", SpatialPolygonBoxes(innerHeads)
    candidate.Add "spatial", result
    Set SpatialCandidate = result
End Function

Private Function SpatialPolygonsOverlap(ByVal first As Collection, ByVal last As Collection, ByVal firstBoxes As Collection, ByVal lastBoxes As Collection) As Boolean
    Dim i As Long, j As Long, box As Variant
    For i = 1 To first.Count
        box = firstBoxes(i)
        For j = 1 To last.Count
            If SpatialBoundsOverlap(box, lastBoxes(j)) Then
                If PlanarPolygonsOverlap(first(i), last(j)) Then
                    SpatialPolygonsOverlap = True
                    Exit Function
                End If
            End If
        Next j
    Next i
End Function

Private Function SpatialLinesHitText(ByVal lines As Collection, ByVal polygons As Collection, ByVal lineBoxes As Collection, ByVal polygonBoxes As Collection) As Boolean
    Dim i As Long, j As Long, box As Variant, line As Variant
    For i = 1 To lines.Count
        box = lineBoxes(i)
        line = lines(i)
        For j = 1 To polygons.Count
            If SpatialBoundsOverlap(box, polygonBoxes(j)) Then
                If ProfileLineCutsMaterial(Array(line(0), line(1)), Array(line(2), line(3)), Array(polygons(j))) Then
                    SpatialLinesHitText = True
                    Exit Function
                End If
            End If
        Next j
    Next i
End Function

Private Function SpatialBasesConflict(ByVal first As Collection, ByVal last As Collection, ByVal firstBoxes As Collection, ByVal lastBoxes As Collection) As Boolean
    Dim i As Long, j As Long, box As Variant
    For i = 1 To first.Count
        box = firstBoxes(i)
        For j = 1 To last.Count
            If SpatialBoundsOverlap(box, lastBoxes(j)) Then
                If PlanarDimensionLineConflict(first(i), last(j)) Then
                    SpatialBasesConflict = True
                    Exit Function
                End If
            End If
        Next j
    Next i
End Function

Private Function SpatialBasesCrossLines(ByVal bases As Collection, ByVal lines As Collection, ByVal baseBoxes As Collection, ByVal lineBoxes As Collection) As Boolean
    Dim i As Long, j As Long, box As Variant, a As Variant, b As Variant
    For i = 1 To bases.Count
        box = baseBoxes(i)
        a = bases(i)
        For j = 1 To lines.Count
            If SpatialBoundsOverlap(box, lineBoxes(j)) Then
                b = lines(j)
                If PlanarSegmentsCross(Array(a(0), a(1)), Array(a(2), a(3)), Array(b(0), b(1)), Array(b(2), b(3))) Then
                    SpatialBasesCrossLines = True
                    Exit Function
                End If
            End If
        Next j
    Next i
End Function

Private Function SpatialCandidatesCompatible(ByVal first As Object, ByVal last As Object) As Boolean
    Dim a As Object, b As Object
    Set a = SpatialCandidate(first)
    Set b = SpatialCandidate(last)
    If SpatialPolygonsOverlap(first("text"), last("text"), a("text"), b("text")) Then
        Exit Function
    End If
    If SpatialLinesHitText(first("lines"), last("text"), a("lines"), b("text")) Or SpatialLinesHitText(last("lines"), first("text"), b("lines"), a("text")) Then
        Exit Function
    End If
    If SpatialBasesConflict(first("bases"), last("bases"), a("bases"), b("bases")) Or SpatialBasesConflict(last("bases"), first("bases"), b("bases"), a("bases")) Then
        Exit Function
    End If
    If planarStrictCrossings Then
        If SpatialBasesCrossLines(first("bases"), last("lines"), a("bases"), b("lines")) Or SpatialBasesCrossLines(last("bases"), first("lines"), b("bases"), a("lines")) Then
            Exit Function
        End If
    End If
    If SpatialPolygonsOverlap(a("heads"), b("heads"), a("headBoxes"), b("headBoxes")) Then
        Exit Function
    End If
    SpatialCandidatesCompatible = True
End Function
