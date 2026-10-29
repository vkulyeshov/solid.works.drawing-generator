Attribute VB_Name = "views"
Option Explicit
Private Function BodyOrdinal(ByVal doc As SldWorks.ModelDoc2, ByVal bodyName As String) As Long
    Dim bodies As Variant, b As Variant
    BodyOrdinal = 1
    bodies = AsPart(doc).GetBodies2(SW_SOLID_BODY, True)
    For Each b In bodies
        If StrComp(b.Name, bodyName, vbTextCompare) < 0 Then
            BodyOrdinal = BodyOrdinal + 1
        End If
    Next b
End Function

Private Function MakeView(ByVal d As SldWorks.ModelDoc2, ByVal doc As SldWorks.ModelDoc2, ByVal cfg As String, ByVal body As Object, ByVal orientation As String, ByVal x As Double, ByVal y As Double) As Object
    Dim v As Object, active As SldWorks.ModelDoc2, nativeName As String, standardId As Long, er As Long
    Select Case orientation
        Case "*Front": standardId = SW_VIEW_FRONT
        Case "*Top": standardId = SW_VIEW_TOP
        Case "*Right": standardId = SW_VIEW_RIGHT
        Case "*Isometric": standardId = SW_VIEW_ISOMETRIC
    End Select
    nativeName = orientation
    If standardId > 0 Then
        Set active = sw.ActivateDoc3(fs.GetFileName(doc.GetPathName), False, SW_REBUILD_USER_DECISION, er)
        If active Is Nothing Then
            Err.Raise vbObjectError + ERR_VIEW_ACTIVATION, , "Cannot activate model for named view: " & er
        End If
        doc.ShowConfiguration2 cfg
        doc.ShowNamedView2 "", standardId
        viewCounter = viewCounter + 1
        nativeName = NAMED_VIEW_PREFIX & CStr(viewCounter)
        doc.NameView nativeName
        Set active = sw.ActivateDoc3(d.GetTitle, False, SW_REBUILD_USER_DECISION, er)
        If active Is Nothing Then
            Err.Raise vbObjectError + ERR_VIEW_ACTIVATION, , "Cannot reactivate drawing: " & er & " " & d.GetTitle
        End If
    End If
    AsDrawing(d).ActivateView "": d.ClearSelection2 True
    Set v = AsDrawing(d).CreateDrawViewFromModelView3(doc.GetPathName, nativeName, x, y, 0)
    If v Is Nothing Then
        Warn "View API returned no view: " & doc.GetPathName & " | " & nativeName
        Exit Function
    End If
    v.ReferencedConfiguration = cfg
    ApplyBody v, body
    v.SetDisplayMode3 False, SW_DISPLAY_HIDDEN_LINES_REMOVED, False, False ' hidden lines removed
    If orientation = "*Isometric" Then
        UprightView v
    End If
    Set MakeView = v
End Function

Private Sub ApplyBody(ByVal v As Object, ByVal b As Object)
    Dim a(0) As Object
    If b Is Nothing Then
        Exit Sub
    End If
    Set a(0) = b
    v.Bodies = a
End Sub

Private Function MakeFlatView(ByVal d As SldWorks.ModelDoc2, ByVal doc As SldWorks.ModelDoc2, ByVal cfg As String, ByVal b As Object) As Object
    Dim er As Long, active As Object, v As Object
    On Error GoTo Failed
    Set active = sw.ActivateDoc3(fs.GetFileName(doc.GetPathName), False, SW_REBUILD_USER_DECISION, er)
    doc.ShowConfiguration2 cfg
    doc.ClearSelection2 True
    If b.Select2(False, Nothing) = False Then
        GoTo Failed
    End If
    Set active = sw.ActivateDoc3(d.GetTitle, False, SW_REBUILD_USER_DECISION, er)
    Set v = AsDrawing(d).CreateFlatPatternViewFromModelView3(doc.GetPathName, cfg, LX(AUXILIARY_COLUMN_X_M), LY(AUXILIARY_ROW_Y_M), 0, False, False)
    If Not v Is Nothing Then
        If v.IsFlatPatternView = False Then
            GoTo Failed
        End If
        v.SetDisplayMode3 False, SW_DISPLAY_HIDDEN_LINES_REMOVED, False, False
        v.ShowSheetMetalBendNotes = (Setting("ShowBendNotes", "0") = "1")
        d.ForceRebuild3 False
        Set MakeFlatView = v
    End If
    Exit Function
Failed:
    Warn "Flat API: " & Err.Description
End Function

Private Function LayoutFlatView(ByVal d As SldWorks.ModelDoc2, ByVal doc As SldWorks.ModelDoc2, ByRef v As SldWorks.View) As Double
    Dim bounds As Variant, position As Variant
    Set v = CreatePlanarFlatView(d, doc, v)
    v.UseSheetScale = False: v.ScaleDecimal = 1: v.Angle = 0
    d.ForceRebuild3 False
    bounds = ProjectedFlatBounds(v)
    Log "FLAT bounds=" & bounds(OUTLINE_LEFT_INDEX) & "," & bounds(OUTLINE_BOTTOM_INDEX) & "," & bounds(OUTLINE_RIGHT_INDEX) & "," & bounds(OUTLINE_TOP_INDEX)
    If bounds(OUTLINE_RIGHT_INDEX) - bounds(OUTLINE_LEFT_INDEX) < MIN_PROJECTED_EXTENT_M Or bounds(OUTLINE_TOP_INDEX) - bounds(OUTLINE_BOTTOM_INDEX) < MIN_PROJECTED_EXTENT_M Then
        Err.Raise vbObjectError + ERR_GENERATION_DATA, , "Empty projected flat geometry."
    End If
    If bounds(OUTLINE_TOP_INDEX) - bounds(OUTLINE_BOTTOM_INDEX) > bounds(OUTLINE_RIGHT_INDEX) - bounds(OUTLINE_LEFT_INDEX) Then
        v.Angle = PI / 2
        d.ForceRebuild3 False
        bounds = ProjectedFlatBounds(v)
    End If
    LayoutFlatView = StandardScale(VIEW_FIT_SAFETY_FACTOR * MinD(FitWidth(AUXILIARY_BOX_WIDTH_M) / (bounds(OUTLINE_RIGHT_INDEX) - bounds(OUTLINE_LEFT_INDEX)), FitHeight(BOTTOM_BOX_HEIGHT_M) / (bounds(OUTLINE_TOP_INDEX) - bounds(OUTLINE_BOTTOM_INDEX))))
    v.ScaleDecimal = LayoutFlatView
    d.ForceRebuild3 False
    bounds = ProjectedFlatBounds(v): position = v.Position
    position(0) = position(0) + LX(AUXILIARY_COLUMN_X_M) - (bounds(OUTLINE_LEFT_INDEX) + bounds(OUTLINE_RIGHT_INDEX)) / 2
    position(1) = position(1) + LY(AUXILIARY_ROW_Y_M) - (bounds(OUTLINE_BOTTOM_INDEX) + bounds(OUTLINE_TOP_INDEX)) / 2
    v.Position = position
    Log "FLAT layout scale=" & LayoutFlatView
End Function

Private Function ProjectedFlatBounds(ByVal v As SldWorks.View) As Variant
    Dim util As SldWorks.MathUtility, inverse As SldWorks.MathTransform, vector As SldWorks.MathVector
    Dim point As SldWorks.MathPoint, directions As Variant, direction(2) As Double, coords(2) As Double, axis As Variant, p As Variant
    Dim x As Double, y As Double, z As Double, bounds(OUTLINE_TOP_INDEX) As Double, i As Long
    Set util = sw.GetMathUtility: Set inverse = v.ModelToViewTransform.Inverse
    directions = Array(Array(-1#, 0#, 0#), Array(0#, -1#, 0#), Array(1#, 0#, 0#), Array(0#, 1#, 0#))
    For i = OUTLINE_LEFT_INDEX To OUTLINE_TOP_INDEX
        direction(0) = directions(i)(0): direction(1) = directions(i)(1): direction(2) = directions(i)(2)
        Set vector = util.CreateVector(direction): Set vector = vector.MultiplyTransform(inverse)
        axis = vector.ArrayData
        If flatBody.GetExtremePoint(axis(0), axis(1), axis(2), x, y, z) = False Then
            Err.Raise vbObjectError + ERR_GENERATION_DATA, , "Cannot calculate exact flat bounds."
        End If
        coords(0) = x: coords(1) = y: coords(2) = z
        Set point = util.CreatePoint(coords): Set point = point.MultiplyTransform(v.ModelToViewTransform)
        p = point.ArrayData: bounds(i) = p(i Mod 2)
    Next i
    ProjectedFlatBounds = bounds
End Function

Private Function CreatePlanarFlatView(ByVal d As SldWorks.ModelDoc2, ByVal doc As SldWorks.ModelDoc2, ByVal v As SldWorks.View) As SldWorks.View
    Dim cfg As String, oldCfg As String, er As Long, active As Object, bodies As Variant, body As SldWorks.Body2
    Dim faces As Variant, item As Variant, face As SldWorks.Face2, bestFace As SldWorks.Face2, area As Double
    Dim viewName As String, oldName As String, wr As Long
    Dim edges As Variant, edgeItem As Variant, edge As SldWorks.Edge, first As SldWorks.Vertex, last As SldWorks.Vertex
    Dim n As Variant, p As Variant, q As Variant, rotation(MATRIX_LAST_INDEX) As Double, length As Double, longest As Double
    Dim ax As Double, ay As Double, az As Double, dx As Double, dy As Double, dz As Double
    Dim util As SldWorks.MathUtility, camera As SldWorks.ModelView, transform As SldWorks.MathTransform, normal As SldWorks.MathVector
    Dim curve As SldWorks.Curve, parameters As SldWorks.CurveParamData, sample As Long, xyz As Variant
    Set doc = v.ReferencedDocument
    PrepareDocumentWindow doc
    oldCfg = doc.ConfigurationManager.ActiveConfiguration.Name: cfg = v.ReferencedConfiguration
    Set active = sw.ActivateDoc3(fs.GetFileName(doc.GetPathName), False, SW_REBUILD_USER_DECISION, er)
    If active Is Nothing Then
        Err.Raise vbObjectError + ERR_FLAT_MODEL_ACTIVATION, , "Cannot activate flat source: " & er
    End If
    If LCase$(active.GetPathName) <> LCase$(doc.GetPathName) Then
        Err.Raise vbObjectError + ERR_WRONG_ACTIVE_MODEL, , "Wrong active model."
    End If
    doc.ShowConfiguration2 cfg: doc.ForceRebuild3 False
    If doc.ConfigurationManager.ActiveConfiguration.Name <> cfg Then
        Err.Raise vbObjectError + ERR_FLAT_CONFIGURATION, , "Missing flat configuration."
    End If
    bodies = AsPart(doc).GetBodies2(SW_SOLID_BODY, True)
    Log "FLAT source config=" & cfg & " bodies=" & (UBound(bodies) + 1)
    For Each item In bodies
        Set body = item
        If InStr(1, body.Name, "Flat-Pattern", vbTextCompare) > 0 Or UBound(bodies) = 0 Then
            Exit For
        End If
    Next item
    Set flatBody = body.Copy
    xyz = body.GetBodyBox
    Log "FLAT exact body=" & body.Name & " box=" & xyz(BBOX_X_MIN_INDEX) & "," & xyz(BBOX_Y_MIN_INDEX) & "," & xyz(BBOX_Z_MIN_INDEX) & "," & xyz(BBOX_X_MAX_INDEX) & "," & xyz(BBOX_Y_MAX_INDEX) & "," & xyz(BBOX_Z_MAX_INDEX)
    faces = body.GetFaces
    For Each item In faces
        Set face = item
        If face.GetSurface.IsPlane Then
            If face.GetArea > area Then
                area = face.GetArea
                Set bestFace = face
            End If
        End If
    Next item
    If bestFace Is Nothing Then
        Err.Raise vbObjectError + ERR_FLAT_FACE, , "No flat plane."
    End If
    n = bestFace.Normal: edges = bestFace.GetEdges
    For Each edgeItem In edges
        Set edge = edgeItem
        If edge.GetCurve.IsLine Then
            Set first = edge.GetStartVertex: Set last = edge.GetEndVertex
            If Not first Is Nothing And Not last Is Nothing Then
                p = first.GetPoint: q = last.GetPoint
                dx = q(0) - p(0): dy = q(1) - p(1): dz = q(2) - p(2)
                length = Sqr(dx * dx + dy * dy + dz * dz)
                If length > longest Then
                    longest = length
                    ax = dx / length
                    ay = dy / length
                    az = dz / length
                End If
            End If
        End If
    Next edgeItem
    If longest = 0 Then
        Err.Raise vbObjectError + ERR_FLAT_EDGE, , "No straight flat edge."
    End If
    rotation(MATRIX_XX_INDEX) = ax: rotation(MATRIX_YX_INDEX) = ay: rotation(MATRIX_ZX_INDEX) = az
    rotation(MATRIX_XY_INDEX) = n(1) * az - n(2) * ay
    rotation(MATRIX_YY_INDEX) = n(2) * ax - n(0) * az
    rotation(MATRIX_ZY_INDEX) = n(0) * ay - n(1) * ax
    rotation(MATRIX_XZ_INDEX) = n(0): rotation(MATRIX_YZ_INDEX) = n(1): rotation(MATRIX_ZZ_INDEX) = n(2): rotation(MATRIX_SCALE_INDEX) = 1
    Set util = sw.GetMathUtility: Set camera = doc.ActiveView
    Set transform = util.CreateTransform(rotation)
    camera.Orientation3 = transform
    doc.ViewZoomtofit2
    doc.ClearSelection2 True
    viewCounter = viewCounter + 1
    viewName = NAMED_FLAT_VIEW_PREFIX & CStr(viewCounter) & "_" & cfg
    doc.NameView viewName
    doc.Save3 SW_SAVE_SILENT, er, wr
    oldName = v.GetName2
    Set active = sw.ActivateDoc3(d.GetTitle, False, SW_REBUILD_USER_DECISION, er)
    d.ClearSelection2 True
    If d.Extension.SelectByID2(oldName, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0) = False Then
        Err.Raise vbObjectError + ERR_FLAT_ORIENTATION, , "Cannot replace flat orientation."
    End If
    d.EditDelete
    Set v = AsDrawing(d).CreateDrawViewFromModelView3(doc.GetPathName, viewName, LX(AUXILIARY_COLUMN_X_M), LY(AUXILIARY_ROW_Y_M), 0)
    If v Is Nothing Then
        Err.Raise vbObjectError + ERR_CREATE_FLAT_VIEW, , "Cannot create planar flat view."
    End If
    v.ReferencedConfiguration = cfg
    v.SetDisplayMode3 False, SW_DISPLAY_HIDDEN_LINES_REMOVED, False, False
    d.ForceRebuild3 False
    Set normal = util.CreateVector(n): Set normal = normal.MultiplyTransform(v.ModelToViewTransform)
    p = normal.ArrayData
    If Abs(p(2)) / Sqr(p(0) * p(0) + p(1) * p(1) + p(2) * p(2)) < FACE_NORMAL_ALIGNMENT_COSINE Then
        Err.Raise vbObjectError + ERR_GENERATION_DATA, , "Flat view is not normal to developed face."
    End If
    Set CreatePlanarFlatView = v
End Function

Private Function BestOrientation(ByVal doc As SldWorks.ModelDoc2, ByVal b As Object) As String
    Dim box As Variant, x As Double, y As Double, z As Double
    If doc.GetType = SW_DOC_ASSEMBLY Then
        box = AsAssembly(doc).GetBox(0)
    Else
        box = b.GetBodyBox
    End If
    x = box(BBOX_X_MAX_INDEX) - box(BBOX_X_MIN_INDEX): y = box(BBOX_Y_MAX_INDEX) - box(BBOX_Y_MIN_INDEX): z = box(BBOX_Z_MAX_INDEX) - box(BBOX_Z_MIN_INDEX)
    If doc.GetType = SW_DOC_ASSEMBLY Then
        Select Case Right$(UCase$(Setting("UpAxis", DEFAULT_UP_AXIS)), 1)
            Case "Y": BestOrientation = IIf(x >= z, "*Front", "*Right")
            Case "Z": BestOrientation = IIf(x >= y, "*Top", "*Right")
            Case "X": BestOrientation = IIf(z >= y, "*Top", "*Front")
            Case Else: Err.Raise vbObjectError + ERR_UP_AXIS, , "UpAxis must be +X, -X, +Y, -Y, +Z or -Z."
        End Select
        Exit Function
    End If
    BestOrientation = "*Front"
    If x * z > x * y And x * z >= y * z Then
        BestOrientation = "*Top"
    End If
    If y * z > x * y And y * z > x * z Then
        BestOrientation = "*Right"
    End If
End Function

Private Function FitFactor(ByVal v As Object, ByVal w As Double, ByVal h As Double) As Double
    Dim a As Variant
    a = v.GetOutline
    If a(OUTLINE_RIGHT_INDEX) - a(OUTLINE_LEFT_INDEX) < MIN_PROJECTED_EXTENT_M Or a(OUTLINE_TOP_INDEX) - a(OUTLINE_BOTTOM_INDEX) < MIN_PROJECTED_EXTENT_M Then
        FitFactor = 1
        Exit Function
    End If
    FitFactor = MinD(w / (a(OUTLINE_RIGHT_INDEX) - a(OUTLINE_LEFT_INDEX)), h / (a(OUTLINE_TOP_INDEX) - a(OUTLINE_BOTTOM_INDEX)))
End Function

Private Function FitScale(ByVal v As Object, ByVal w As Double, ByVal h As Double) As Double
    FitScale = StandardScale(v.ScaleDecimal * FitFactor(v, w, h))
End Function

Private Function StandardScale(ByVal limit As Double) As Double
    Dim a As Variant, i As Long
    a = Array(SCALE_10_TO_1, SCALE_5_TO_1, SCALE_4_TO_1, SCALE_2_5_TO_1, SCALE_2_TO_1, SCALE_1_TO_1, SCALE_1_TO_2, SCALE_1_TO_2_5, SCALE_1_TO_4, SCALE_1_TO_5, SCALE_1_TO_10, SCALE_1_TO_20, SCALE_1_TO_25, SCALE_1_TO_40, SCALE_1_TO_50, SCALE_1_TO_100, SCALE_1_TO_200, SCALE_1_TO_400, SCALE_1_TO_500, SCALE_1_TO_1000)
    For i = 0 To UBound(a)
        If CDbl(a(i)) <= limit + SCALE_COMPARISON_TOLERANCE Then
            StandardScale = a(i)
            Exit Function
        End If
    Next i
    Err.Raise vbObjectError + ERR_UNSUPPORTED_SCALE, , "Model does not fit the supported drawing scale range."
End Function

Private Sub CenterView(ByVal v As Object, ByVal x As Double, ByVal y As Double)
    Dim a As Variant, p As Variant
    a = v.GetOutline: p = v.Position
    p(0) = p(0) + x - (a(OUTLINE_LEFT_INDEX) + a(OUTLINE_RIGHT_INDEX)) / 2
    p(1) = p(1) + y - (a(OUTLINE_BOTTOM_INDEX) + a(OUTLINE_TOP_INDEX)) / 2
    v.Position = p
End Sub

Private Function VisibleEdgeCount(ByVal v As Object) As Long
    Dim c As Variant, entities As Variant, i As Long
    c = v.GetVisibleComponents
    If IsEmpty(c) Then
        Exit Function
    End If
    For i = LBound(c) To UBound(c)
        entities = v.GetVisibleEntities2(c(i), 1)
        If Not IsEmpty(entities) Then
            VisibleEdgeCount = VisibleEdgeCount + UBound(entities) - LBound(entities) + 1
        End If
    Next i
End Function

