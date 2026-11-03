Attribute VB_Name = "planar_features"
Option Explicit

' Constant-thickness prismatic bodies may have circular through holes. Bent,
' tapered and freeform bodies keep the normal solid/sheet-metal workflow.
Private Function PlanarFeatureFace(ByVal body As SldWorks.Body2) As SldWorks.Face2
    Dim item As Variant, faces As Variant, face As SldWorks.Face2, chosen As SldWorks.Face2, surface As SldWorks.Surface
    Dim normal As Variant, values As Variant, point As Variant, vertices As Variant, vertex As SldWorks.Vertex
    Dim area As Double, x As Double, y As Double, z As Double, low As Double, high As Double, coordinate As Double
    faces = body.GetFaces
    For Each item In faces
        Set face = item
        If face.GetSurface.IsPlane And face.GetArea > area Then
            area = face.GetArea
            Set chosen = face
        End If
    Next item
    If chosen Is Nothing Then
        Log "PLANAR unsupported: no planar face"
        Exit Function
    End If
    normal = chosen.Normal
    Log "PLANAR FACE area=" & area & "; normal=" & normal(0) & "," & normal(1) & "," & normal(2)
    If Not body.GetExtremePoint(CDbl(normal(0)), CDbl(normal(1)), CDbl(normal(2)), x, y, z) Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot project the planar cap thickness."
    End If
    high = x * normal(0) + y * normal(1) + z * normal(2)
    If Not body.GetExtremePoint(-CDbl(normal(0)), -CDbl(normal(1)), -CDbl(normal(2)), x, y, z) Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot project the opposite planar cap."
    End If
    low = x * normal(0) + y * normal(1) + z * normal(2)
    Log "PLANAR INSPECT cap area=" & area & "; depth=" & CStr(high - low)
    If high - low <= PROFILE_COORDINATE_TOLERANCE_M Or high - low > Sqr(area) * PROFILE_MAX_THICKNESS_RATIO Then
        Log "PLANAR unsupported: thickness ratio"
        Exit Function
    End If
    vertices = body.GetVertices
    For Each item In vertices
        Set vertex = item
        point = vertex.GetPoint
        coordinate = point(0) * normal(0) + point(1) * normal(1) + point(2) * normal(2)
        If Abs(coordinate - low) > PROFILE_COORDINATE_TOLERANCE_M And Abs(coordinate - high) > PROFILE_COORDINATE_TOLERANCE_M Then
            Log "PLANAR unsupported: vertex between cap planes; coordinate=" & coordinate
            Exit Function
        End If
    Next item
    For Each item In faces
        Set face = item
        Set surface = face.GetSurface
        If surface.IsCylinder Then
            values = surface.CylinderParams
            If Abs(values(3) * normal(0) + values(4) * normal(1) + values(5) * normal(2)) < FACE_NORMAL_ALIGNMENT_COSINE Then
                Log "PLANAR unsupported: cylinder axis=" & values(3) & "," & values(4) & "," & values(5)
                Exit Function
            End If
        ElseIf surface.IsPlane = False Then
            Log "PLANAR unsupported: non-prismatic surface; identity=" & surface.Identity & "; isPlane=" & CStr(surface.IsPlane)
            Exit Function
        End If
    Next item
    Set PlanarFeatureFace = chosen
End Function

Private Function PlanarFeatureNode(ByVal view As SldWorks.View, ByVal point As Variant) As Object
    Dim node As Object, matrix As Variant
    Set node = CreateObject("Scripting.Dictionary")
    matrix = view.ModelToViewTransform.ArrayData
    node.Add "point", point
    node.Add "x", point(0) * matrix(0) + point(1) * matrix(3) + point(2) * matrix(6)
    node.Add "y", point(0) * matrix(1) + point(1) * matrix(4) + point(2) * matrix(7)
    Set PlanarFeatureNode = node
End Function

Private Function PlanarHoleLoops(ByVal face As SldWorks.Face2) As Collection
    Dim holes As New Collection, boundary As SldWorks.Loop2, item As Variant, loops As Variant, edges As Variant
    Dim edge As SldWorks.Edge, data As Variant, hole As Object, a As SldWorks.Vertex, b As SldWorks.Vertex
    loops = face.GetLoops
    For Each item In loops
        Set boundary = item
        If boundary.IsOuter = False Then
            edges = boundary.GetEdges
            If UBound(edges) <> LBound(edges) Then
                Log "PLANAR unsupported: multi-edge inner loop"
                Exit Function
            End If
            Set edge = edges(LBound(edges))
            If edge.GetCurve.IsCircle = False Then
                Log "PLANAR unsupported: non-circular inner loop"
                Exit Function
            End If
            Set a = edge.GetStartVertex
            Set b = edge.GetEndVertex
            If Not a Is Nothing And Not b Is Nothing Then
                If Not ProfilePointClose(a.GetPoint, b.GetPoint) Then
                    Exit Function
                End If
            End If
            data = edge.GetCurve.CircleParams
            Set hole = CreateObject("Scripting.Dictionary")
            hole.Add "point", Array(data(0), data(1), data(2))
            hole.Add "radius", data(CIRCLE_RADIUS_INDEX)
            holes.Add hole
        End If
    Next item
    Set PlanarHoleLoops = holes
End Function

Private Function PlanarAngle(ByVal dx As Double, ByVal dy As Double) As Double
    If Abs(dx) < DIRECTION_ZERO_TOLERANCE Then
        PlanarAngle = Sgn(dy) * PI / 2
    Else
        PlanarAngle = Atn(dy / dx)
        If dx < 0 Then
            PlanarAngle = PlanarAngle + PI
        End If
    End If
End Function

Private Function PlanarVisibleLine(ByVal view As SldWorks.View, ByVal first As Variant, ByVal last As Variant) As SldWorks.Edge
    Dim components As Variant, component As Variant, edges As Variant, item As Variant, edge As SldWorks.Edge
    Dim a As SldWorks.Vertex, b As SldWorks.Vertex
    components = view.GetVisibleComponents
    If IsEmpty(components) Then
        Exit Function
    End If
    For Each component In components
        edges = view.GetVisibleEntities2(component, SW_VISIBLE_ENTITY_EDGE)
        If IsEmpty(edges) Then
            GoTo NextPlanarLineComponent
        End If
        For Each item In edges
            Set edge = item
            If edge.GetCurve.IsLine Then
                Set a = edge.GetStartVertex
                Set b = edge.GetEndVertex
                If Not a Is Nothing And Not b Is Nothing Then
                    If (PlanarProjectedClose(view, a.GetPoint, first) And PlanarProjectedClose(view, b.GetPoint, last)) Or (PlanarProjectedClose(view, a.GetPoint, last) And PlanarProjectedClose(view, b.GetPoint, first)) Then
                        Set PlanarVisibleLine = edge
                        Exit Function
                    End If
                End If
            End If
        Next item
NextPlanarLineComponent:
    Next component
End Function

Private Function PlanarProjectedClose(ByVal view As SldWorks.View, ByVal first As Variant, ByVal last As Variant) As Boolean
    Dim a As Variant, b As Variant
    a = PatternSheetPoint(view, first)
    b = PatternSheetPoint(view, last)
    PlanarProjectedClose = Abs(a(0) - b(0)) <= PROFILE_COORDINATE_TOLERANCE_M * view.ScaleDecimal And Abs(a(1) - b(1)) <= PROFILE_COORDINATE_TOLERANCE_M * view.ScaleDecimal
End Function

Private Function PlanarVisibleCircle(ByVal view As SldWorks.View, ByVal hole As Object) As SldWorks.Edge
    Dim components As Variant, component As Variant, edges As Variant, item As Variant, edge As SldWorks.Edge, data As Variant, target As Variant
    target = hole("point")
    components = view.GetVisibleComponents
    For Each component In components
        edges = view.GetVisibleEntities2(component, SW_VISIBLE_ENTITY_EDGE)
        For Each item In edges
            Set edge = item
            If edge.GetCurve.IsCircle Then
                data = edge.GetCurve.CircleParams
                If Abs(data(CIRCLE_RADIUS_INDEX) - hole("radius")) <= PROFILE_COORDINATE_TOLERANCE_M Then
                    If PlanarProjectedClose(view, Array(data(0), data(1), data(2)), target) Then
                        Set PlanarVisibleCircle = edge
                        Exit Function
                    End If
                End If
            End If
        Next item
    Next component
End Function
