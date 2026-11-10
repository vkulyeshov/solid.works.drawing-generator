Attribute VB_Name = "planar_drawing"
Option Explicit

Private Sub RememberPlanarDimension(ByVal display As SldWorks.DisplayDimension, ByVal request As Object, ByVal view As SldWorks.View)
    Dim name As String, audit As Object
    If display Is Nothing Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot create a planned native planar dimension."
    End If
    name = display.GetNameForSelection
    request.Add "view", view.Name
    request.Add "value", display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue
    planarDimensionRequests.Add name, request
    featureDimensionOwners(name) = "planar|" & request("kind") & "|" & request("source")
    ApplyGostDimensionStyle display
    Set audit = fs.OpenTextFile(fs.BuildPath(runDir, PLANAR_DIMENSIONS_FILENAME), FILE_FOR_APPENDING, True, FILE_UNICODE)
    audit.WriteLine jobNumber & vbTab & view.Name & vbTab & name & vbTab & request("kind") & vbTab & request("value") & vbTab & request("source")
    audit.Close
End Sub

Private Sub AddPlanarEdgeDimension(ByVal d As SldWorks.ModelDoc2, ByVal view As SldWorks.View, ByVal group As Object, ByVal winding As Double)
    Dim edge As SldWorks.Edge, display As SldWorks.DisplayDimension, request As Object, a As Object, b As Object, p As Variant, q As Variant
    Dim dx As Double, dy As Double, length As Double
    Set a = group("first")
    Set b = group("last")
    Set edge = PlanarVisibleLine(view, a("point"), b("point"))
    If edge Is Nothing Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "A selected planar contour edge is not visible."
    End If
    p = PatternSheetPoint(view, a("point"))
    q = PatternSheetPoint(view, b("point"))
    dx = q(0) - p(0)
    dy = q(1) - p(1)
    length = Sqr(dx * dx + dy * dy)
    AsDrawing(d).ActivateView ""
    d.ClearSelection2 True
    If Not SelectFeatureReference(d, view, edge) Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot select a planar length reference."
    End If
    Set display = d.AddDimension2((p(0) + q(0)) / 2 + dy / length * Sgn(winding) * GOST_OBJECT_OFFSET_M, (p(1) + q(1)) / 2 - dx / length * Sgn(winding) * GOST_OBJECT_OFFSET_M, 0)
    If display Is Nothing Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot dimension a planar edge."
    End If
    If Abs(display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue - group("value")) > PROFILE_COORDINATE_TOLERANCE_M Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "An aligned dimension differs from the true edge length."
    End If
    If group("count") > 1 Then
        display.SetText SW_DIMENSION_TEXT_PREFIX, group("count") & "x "
    End If
    Set request = CreateObject("Scripting.Dictionary")
    request.Add "kind", "edge"
    request.Add "source", group("id")
    request.Add "first", a("point")
    request.Add "last", b("point")
    request.Add "winding", Sgn(winding)
    RememberPlanarDimension display, request, view
End Sub

Private Function MakePlanarAngleDatum(ByVal d As SldWorks.ModelDoc2, ByVal view As SldWorks.View, ByVal anchor As Variant, ByVal direction As Variant) As SldWorks.SketchSegment
    Dim sketch As SldWorks.Sketch, utility As SldWorks.MathUtility, point As SldWorks.MathPoint
    Dim first As Variant, last As Variant, target(2) As Double, manager As SldWorks.SketchManager, oldAddToDB As Boolean
    Dim errorNumber As Long, errorText As String
    On Error GoTo Failed
    Set sketch = view.GetSketch
    Set utility = sw.GetMathUtility
    target(0) = anchor(0)
    target(1) = anchor(1)
    Set point = utility.CreatePoint(target)
    first = point.MultiplyTransform(sketch.ModelToSketchTransform).ArrayData
    target(0) = target(0) + direction(0) * PLANAR_ANGLE_DATUM_LENGTH_M
    target(1) = target(1) + direction(1) * PLANAR_ANGLE_DATUM_LENGTH_M
    Set point = utility.CreatePoint(target)
    last = point.MultiplyTransform(sketch.ModelToSketchTransform).ArrayData
    AsDrawing(d).ActivateView view.Name
    d.ClearSelection2 True
    Set manager = d.SketchManager
    oldAddToDB = manager.AddToDB
    manager.AddToDB = True
    Set MakePlanarAngleDatum = manager.CreateLine(first(0), first(1), 0, last(0), last(1), 0)
    MakePlanarAngleDatum.ConstructionGeometry = True
    Log "PLANAR DATUM anchor=" & anchor(0) & "," & anchor(1) & "; local=" & first(0) & "," & first(1) & "; end=" & last(0) & "," & last(1) & "; length=" & MakePlanarAngleDatum.GetLength
    manager.AddToDB = oldAddToDB
    Exit Function
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    On Error Resume Next
    If Not manager Is Nothing Then
        manager.AddToDB = oldAddToDB
    End If
    On Error GoTo 0
    Err.Raise errorNumber, "MakePlanarAngleDatum", errorText
End Function

Private Sub SelectPlanarAngularDimension(ByVal d As SldWorks.ModelDoc2, ByVal display As SldWorks.DisplayDimension)
    d.ClearSelection2 True
    If d.Extension.SelectByID2(display.GetNameForSelection, SW_SELECTION_DIMENSION_TYPE, 0, 0, 0, False, SW_DEFAULT_SELECTION_MARK, Nothing, SW_DEFAULT_SELECTION_OPTIONS) = False Then
        If display.GetAnnotation.Select3(False, Nothing) = False Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot select the native angular dimension."
        End If
    End If
End Sub

Private Sub SetPlanarAngleRepresentation(ByVal d As SldWorks.ModelDoc2, ByVal display As SldWorks.DisplayDimension, ByVal expected As Double)
    Dim attempt As Long, annotation As SldWorks.Annotation, changed As Boolean
    For attempt = 0 To PLANAR_ANGLE_SECTOR_VARIANTS - 1
        If Abs(display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue - expected) < PLANAR_ANGLE_TOLERANCE_RAD Then
            Exit Sub
        End If
        SelectPlanarAngularDimension d, display
        If attempt Mod 2 = 0 Then
            changed = display.SupplementaryAngle
        Else
            changed = display.ExplementaryAngle
        End If
        d.GraphicsRedraw2
    Next attempt
    If Abs(display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue - expected) >= PLANAR_ANGLE_TOLERANCE_RAD Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "The local angle differs from the measured model direction."
    End If
End Sub

Private Sub AddPlanarAngleDimension(ByVal d As SldWorks.ModelDoc2, ByVal view As SldWorks.View, ByVal group As Object, ByVal winding As Double)
    Dim first As SldWorks.Edge, datum As SldWorks.SketchSegment, display As SldWorks.DisplayDimension, request As Object
    Dim a As Object, b As Object, p As Variant, q As Variant, dx As Double, dy As Double, length As Double
    Dim direction As Variant, axis As Long, bx As Double, by As Double, bisectorLength As Double
    Dim selection As SldWorks.SelectData, manager As SldWorks.SelectionMgr, result As Long
    Dim expected As Double
    Dim datumPlan As Variant, anchor As Variant, polygon() As Variant, index As Long
    Set a = group("first")
    Set b = group("last")
    Set first = PlanarVisibleLine(view, a("point"), b("point"))
    p = PatternSheetPoint(view, a("point"))
    q = PatternSheetPoint(view, b("point"))
    dx = q(0) - p(0)
    dy = q(1) - p(1)
    length = Sqr(dx * dx + dy * dy)
    ReDim polygon(planarContourPoints.Count - 1)
    For index = 1 To planarContourPoints.Count
        polygon(index - 1) = PatternSheetPoint(view, planarContourPoints(index))
    Next index
    datumPlan = PlanarDatumPlan(p, q, PlanarClipPolygon(polygon, view.GetOutline))
    anchor = p
    bx = -dx / length
    by = -dy / length
    If CLng(datumPlan(0)) = EDGE_ENDPOINT_LAST Then
        anchor = q
        bx = dx / length
        by = dy / length
    End If
    axis = CLng(datumPlan(1))
    direction = datumPlan(2)
    expected = CDbl(datumPlan(3))
    Set datum = MakePlanarAngleDatum(d, view, anchor, direction)
    If datum Is Nothing Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot create a local principal-direction angle datum."
    End If
    datum.Color = DrawingDimensionColor()
    datum.Width = SW_LINE_WEIGHT_DIMENSION
    datum.LayerOverride = datum.LayerOverride Or SW_LAYER_OVERRIDE_COLOR Or SW_LAYER_OVERRIDE_WIDTH
    d.ClearSelection2 True
    AsDrawing(d).ActivateView view.Name
    If Not SelectFeatureReference(d, view, first) Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot select inclined planar angle references."
    End If
    Set manager = d.SelectionManager
    Set selection = manager.CreateSelectData
    If datum.Select4(True, selection) = False Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot select the local angle datum."
    End If
    bx = bx + direction(0)
    by = by + direction(1)
    bisectorLength = Sqr(bx * bx + by * by)
    Log "PLANAR ANGLE selections=" & manager.GetSelectedObjectCount2(SW_SELECTION_ALL_MARKS) & "; datum axis=" & axis
    Set display = d.Extension.AddSpecificDimension(anchor(0) + bx / bisectorLength * GOST_OBJECT_OFFSET_M, anchor(1) + by / bisectorLength * GOST_OBJECT_OFFSET_M, 0, SW_DIMENSION_ANGULAR, result)
    If display Is Nothing Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot create the angle of an inclined contour: " & result
    End If
    If display.Type2 <> SW_DIMENSION_ANGULAR Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Inclined edge references did not create an angular dimension."
    End If
    display.GetDimension2(SW_DIMENSION_INDEX_FIRST).DrivenState = SW_DIMENSION_DRIVEN
    SetPlanarAngleRepresentation d, display, expected
    Set request = CreateObject("Scripting.Dictionary")
    request.Add "kind", "angle"
    request.Add "datumAxis", axis
    request.Add "source", group("id")
    Set a = group("first")
    Set b = group("last")
    request.Add "first", a("point")
    request.Add "last", b("point")
    request.Add "winding", Sgn(winding)
    RememberPlanarDimension display, request, view
End Sub

Private Function MakePlanarEndDetail(ByVal d As SldWorks.ModelDoc2, ByVal overview As SldWorks.View, ByVal plan As Collection, ByVal centerX As Double, ByVal longest As Double, ByVal owner As Long) As SldWorks.View
    Dim group As Object, a As Object, b As Object, anchor As Object, bounds As Variant, index As Long, belongs As Boolean
    Dim p As Variant, x As Double, y As Double, radius As Double, profile As SldWorks.SketchSegment, view As SldWorks.View
    Dim selection As SldWorks.SelectData, parent As SldWorks.View
    Dim manager As SldWorks.SelectionMgr
    Dim stage As String, errorNumber As Long, errorText As String
    Dim detailScale As Double
    Dim label As String, labels As Variant
    On Error GoTo Failed
    bounds = Array(SEARCH_INITIAL_LIMIT, SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT)
    For Each group In plan
        If group("value") < longest * PLANAR_LOCAL_EDGE_RATIO Then
            Set a = group("first")
            Set b = group("last")
            belongs = (a("x") + b("x")) / 2 <= centerX
            If (owner = PLANAR_LEFT_VIEW_INDEX And belongs) Or (owner = PLANAR_RIGHT_VIEW_INDEX And Not belongs) Then
                Set anchor = a
                NativeExpandBounds bounds, a("x"), a("y")
                NativeExpandBounds bounds, b("x"), b("y")
            End If
        End If
    Next group
    If anchor Is Nothing Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "A planar end detail has no local features."
    End If
    p = PatternSheetPoint(overview, anchor("point"))
    x = p(0) + ((bounds(0) + bounds(2)) / 2 - anchor("x")) * overview.ScaleDecimal
    y = p(1) + ((bounds(1) + bounds(3)) / 2 - anchor("y")) * overview.ScaleDecimal
    radius = DetailProfileRadius(bounds)
    detailScale = MinD(PLANAR_DETAIL_SCALE, StandardScale((sheetHeight * PLANAR_ROW_PITCH_RATIO - 2 * GOST_OBJECT_OFFSET_M) / (2 * radius + PLANAR_DETAIL_BOUNDARY_MARGIN_M)))
    Log "PLANAR DETAIL PROFILE owner=" & owner & "; center=" & x & "," & y & "; radius=" & radius * overview.ScaleDecimal
    stage = "profile creation"
    Set profile = CreateDetailProfile(d, overview, x, y, radius * overview.ScaleDecimal)
    If profile Is Nothing Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot create the planar detail boundary."
    End If
    stage = "selection manager"
    Set manager = d.SelectionManager
    Set selection = manager.CreateSelectData
    stage = "source view selection"
    AsDrawing(d).ActivateView overview.Name
    profile.Select4 False, selection
    stage = "detail view creation"
    labels = Split(PLANAR_DETAIL_LABELS, "|")
    label = labels(owner - PLANAR_LEFT_VIEW_INDEX)
    Set view = AsDrawing(d).CreateDetailViewAt4(sheetWidth * PLANAR_CENTER_X_RATIO, sheetHeight * (PLANAR_FIRST_Y_RATIO - (owner - 1) * PLANAR_ROW_PITCH_RATIO), 0, SW_DETAIL_STANDARD, 1, 1 / detailScale, label, SW_DETAIL_CIRCLE, False, False, False, SW_DETAIL_OUTLINE_INTENSITY)
    If view Is Nothing Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot create an enlarged planar end detail."
    End If
    stage = "detail parent"
    Set parent = view.GetBaseView
    If parent Is Nothing Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "The planar detail has no source model view."
    End If
    Log "PLANAR DETAIL PARENT " & parent.Name
    planarDetailLabels.Add view.Name, label
    view.DisableAutoUpdate = False
    view.UpdateViewDisplayGeometry
    d.GraphicsRedraw2
    d.ForceRebuild3 False
    d.ForceRebuild3 False
    Log "PLANAR DETAIL " & view.Name & "; scale=" & view.ScaleDecimal
    Set MakePlanarEndDetail = view
    Exit Function
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    Err.Raise errorNumber, "MakePlanarEndDetail", stage & ": " & errorText
End Function

Private Sub AddPlanarHoleLocations(ByVal d As SldWorks.ModelDoc2, ByVal view As SldWorks.View, ByVal holes As Collection, ByVal reference As Object)
    Dim hole As Object, previous As Object, datum As Object, a As Object, b As Object, edge As SldWorks.Edge, priorEdge As SldWorks.Edge
    Dim vertex As SldWorks.Vertex, display As SldWorks.DisplayDimension, request As Object, axis As Long, p As Variant, origin As Variant, datumPoint As Variant
    Dim index As Long, other As Long, ordered As New Collection, best As Object, used As Object, chosen As Long
    Set used = CreateObject("Scripting.Dictionary")
    For index = 1 To holes.Count
        Set best = Nothing
        For other = 1 To holes.Count
            If Not used.Exists(other) Then
                Set hole = holes(other)
                If best Is Nothing Then
                    Set best = hole
                    chosen = other
                ElseIf hole("x") < best("x") Then
                    Set best = hole
                    chosen = other
                End If
            End If
        Next other
        ordered.Add best
        used(chosen) = True
    Next index
    Set a = reference("first")
    Set b = reference("last")
    Set datum = a
    If b("x") < a("x") Then
        Set datum = b
    End If
    For Each hole In ordered
        For axis = SCREEN_X_AXIS_INDEX To SCREEN_Y_AXIS_INDEX
            origin = datum("point")
            If Not previous Is Nothing Then
                origin = previous("point")
                If axis = SCREEN_Y_AXIS_INDEX And Abs(hole("y") - previous("y")) <= PROFILE_COORDINATE_TOLERANCE_M Then
                    AddPlanarHoleAxis d, view, previous, hole
                    GoTo NextHoleAxis
                End If
            End If
            Set edge = PlanarVisibleCircle(view, hole)
            d.ClearSelection2 True
            AsDrawing(d).ActivateView ""
            If previous Is Nothing Then
                Set vertex = ProfileReference(view, origin)
                If Not SelectFeatureReference(d, view, vertex) Then
                    Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot select the hole location datum."
                End If
            Else
                Set priorEdge = PlanarVisibleCircle(view, previous)
                If Not SelectFeatureReference(d, view, priorEdge) Then
                    Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot select the previous hole center."
                End If
            End If
            If Not SelectFeatureReference(d, view, edge, True) Then
                Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot select a hole center reference."
            End If
            p = PatternSheetPoint(view, hole("point"))
            If axis = SCREEN_X_AXIS_INDEX Then
                Set display = d.AddHorizontalDimension2(p(0), p(1) + GOST_OBJECT_OFFSET_M, 0)
            Else
                Set display = d.AddVerticalDimension2(p(0) - GOST_OBJECT_OFFSET_M, p(1), 0)
            End If
            If display Is Nothing Then
                Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot create a hole-center location dimension."
            End If
            display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SetArcEndCondition DIMENSION_REFERENCE_FIRST, SW_ARC_CENTER
            display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SetArcEndCondition DIMENSION_REFERENCE_SECOND, SW_ARC_CENTER
            datumPoint = PatternSheetPoint(view, origin)
            If Abs(display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue - Abs(p(axis) - datumPoint(axis)) / view.ScaleDecimal) > PROFILE_COORDINATE_TOLERANCE_M Then
                Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "A hole location does not measure the planned center coordinate."
            End If
            Set request = CreateObject("Scripting.Dictionary")
            request.Add "kind", "hole-location"
            request.Add "source", NativePairKey(origin, hole("point")) & "|" & axis
            request.Add "first", origin
            request.Add "last", hole("point")
            request.Add "axis", axis
            RememberPlanarDimension display, request, view
NextHoleAxis:
        Next axis
        Set previous = hole
    Next hole
End Sub

Private Sub AddPlanarHoleAxis(ByVal d As SldWorks.ModelDoc2, ByVal view As SldWorks.View, ByVal first As Object, ByVal last As Object)
    Dim a As SldWorks.Edge, b As SldWorks.Edge, marks As SldWorks.CenterMark
    Set a = PlanarVisibleCircle(view, first)
    Set b = PlanarVisibleCircle(view, last)
    AsDrawing(d).ActivateView ""
    d.ClearSelection2 True
    If Not SelectFeatureReference(d, view, a) Or Not SelectFeatureReference(d, view, b, True) Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot select the aligned hole pattern."
    End If
    Set marks = AsDrawing(d).InsertCenterMark3(SW_CENTER_MARK_LINEAR_GROUP, True, False)
    If marks Is Nothing Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot show the verified alignment of hole centers."
    End If
    d.ClearSelection2 True
End Sub

Private Function MakePlanarFeatureDrawing(ByRef d As SldWorks.ModelDoc2, ByVal doc As SldWorks.ModelDoc2, ByVal cfg As String, ByVal body As SldWorks.Body2, ByVal outName As String) As Boolean
    Dim face As SldWorks.Face2, raw As Collection, holes As Collection, projected As New Collection, projectedHoles As New Collection
    Dim boundary As SldWorks.Loop2, item As Variant, loops As Variant, node As Object, hole As Object, points As Collection
    Dim first As SldWorks.View, view As SldWorks.View, views As New Collection, plan As Collection, group As Object, reference As Object
    Dim a As Object, b As Object, longest As Double, area As Double, angle As Double, outline As Variant, viewScaleValue As Double
    Dim index As Long, owner As Long, centerX As Double, directions As Object, key As String, rc As Long, display As SldWorks.DisplayDimension
    Dim request As Object, audit As Object, normal As Variant, box As Variant, thickness As Double, angleReference As Object
    Dim error As Long, warning As Long, nativePath As String
    Dim ownerViews As Object, endSource As SldWorks.View, referenceView As SldWorks.View
    Dim firstSheet As String, detailSheet As String, spread As Boolean, localCount As Long, hasFeatures As Boolean
    Set face = PlanarFeatureFace(body)
    If face Is Nothing Then
        Exit Function
    End If
    Set holes = PlanarHoleLoops(face)
    If holes Is Nothing Then
        Exit Function
    End If
    loops = face.GetLoops
    For Each item In loops
        Set boundary = item
        If boundary.IsOuter Then
            Set raw = ProfileLoopPoints(boundary)
        End If
    Next item
    If raw Is Nothing Then
        Log "PLANAR unsupported: curved or disconnected outer loop"
        Exit Function
    End If
    Set first = MakeView(d, doc, cfg, body, BestOrientation(doc, body), sheetWidth * PLANAR_CENTER_X_RATIO, sheetHeight * PLANAR_FIRST_Y_RATIO)
    first.UseSheetScale = False
    first.ScaleDecimal = 1
    d.ForceRebuild3 False
    Set points = ProfileProjection(raw, first)
    For index = 1 To points.Count
        Set a = points(index)
        Set b = points(index Mod points.Count + 1)
        If PlanarDistance(a, b) > longest Then
            longest = PlanarDistance(a, b)
            angle = PlanarAngle(b("x") - a("x"), b("y") - a("y"))
        End If
    Next index
    first.Angle = first.Angle - angle
    d.ForceRebuild3 False
    Set points = ProfileProjection(raw, first)
    For Each item In holes
        Set hole = item
        Set node = PlanarFeatureNode(first, hole("point"))
        node.Add "radius", hole("radius")
        projectedHoles.Add node
    Next item
    Set plan = PlanarEdgePlan(points, projectedHoles)
    Set planarContourPoints = raw
    For index = 1 To points.Count
        Set a = points(index)
        Set b = points(index Mod points.Count + 1)
        centerX = centerX + a("x") / points.Count
        area = area + a("x") * b("y") - b("x") * a("y")
    Next index
    Set planarDimensionRequests = CreateObject("Scripting.Dictionary")
    Set planarDetailLabels = CreateObject("Scripting.Dictionary")
    Set featureDimensionOwners = CreateObject("Scripting.Dictionary")
    Set featureDimensionPlacement = CreateObject("Scripting.Dictionary")
    Set featureDimensionLanes = CreateObject("Scripting.Dictionary")
    Set directions = CreateObject("Scripting.Dictionary")
    longest = 0
    For Each group In plan
        If group("value") > longest Then
            Set reference = group
            longest = group("value")
        End If
    Next group
    For Each group In plan
        If CBool(group("selected")) And group("value") < longest * PLANAR_LOCAL_EDGE_RATIO Then
            localCount = localCount + 1
        End If
    Next group
    spread = localCount > PLANAR_PAGE_LOCAL_DIMENSION_LIMIT
    viewScaleValue = FitScale(first, sheetWidth * PLANAR_WIDTH_RATIO, sheetHeight * PLANAR_HEIGHT_RATIO)
    SetPaperScale AsDrawing(d).GetCurrentSheet, viewScaleValue
    first.UseSheetScale = True
    CenterView first, sheetWidth * PLANAR_CENTER_X_RATIO, sheetHeight * PLANAR_FIRST_Y_RATIO
    For Each group In plan
        If group("value") >= longest * PLANAR_LOCAL_EDGE_RATIO And CBool(group("selected")) Then
            AddPlanarEdgeDimension d, first, group, area
        End If
    Next group
    AddPlanarHoleLocations d, first, projectedHoles, reference
    rc = DiameterDimensions(d, first)
    Set display = first.GetFirstDisplayDimension5
    Do While Not display Is Nothing
        If display.Type2 = SW_DIMENSION_DIAMETER Then
            display.SetSecondArrow False, False
        End If
        Set display = display.GetNext5
    Loop
    AsDrawing(d).ActivateView first.Name
    doc.Save3 SW_SAVE_SILENT, error, warning
    nativePath = runDir & DRAWINGS_RELATIVE_DIRECTORY & outName & DRAWING_EXTENSION
    If d.Extension.SaveAs(nativePath, SW_SAVE_CURRENT_VERSION, SW_SAVE_SILENT, Nothing, error, warning) = False Then
        Err.Raise vbObjectError + ERR_SAVE_DRAWING, , "Cannot persist the planar overview before creating local details."
    End If
    sw.CloseDoc d.GetTitle
    Set d = sw.OpenDoc6(nativePath, SW_DOC_DRAWING, SW_OPEN_SILENT, "", error, warning)
    If d Is Nothing Then
        Err.Raise vbObjectError + ERR_REOPEN_DRAWING, , "Cannot refresh the planar overview sketch transform."
    End If
    Set currentDrawing = d
    PrepareDocumentWindow d
    d.ForceRebuild3 False
    Set first = AsDrawing(d).GetFirstView.GetNextView
    firstSheet = AsDrawing(d).GetCurrentSheet.GetName
    detailSheet = firstSheet
    Set ownerViews = CreateObject("Scripting.Dictionary")
    ownerViews.Add CStr(PLANAR_OVERVIEW_INDEX), first
    views.Add first
    Set endSource = first
    If spread Then
        NewFormattedSheet d, "Місцеві елементи"
        detailSheet = AsDrawing(d).GetCurrentSheet.GetName
        Set referenceView = MakeView(d, doc, cfg, body, first.GetOrientationName, sheetWidth * PLANAR_CENTER_X_RATIO, sheetHeight - FAMILY_REFERENCE_CENTER_MARGIN_M)
        referenceView.Angle = first.Angle
        referenceView.UseSheetScale = False
        referenceView.ScaleDecimal = 1
        d.ForceRebuild3 False
        referenceView.ScaleDecimal = FitScale(referenceView, sheetWidth - LEADER_WORK_LEFT_M - LEADER_WORK_EDGE_MARGIN_M, FAMILY_REFERENCE_HEIGHT_M)
        d.ForceRebuild3 False
        CenterView referenceView, sheetWidth * PLANAR_CENTER_X_RATIO, sheetHeight - FAMILY_REFERENCE_CENTER_MARGIN_M
        Set endSource = referenceView
    End If
    For index = 1 To PLANAR_LOCAL_VIEW_COUNT
        owner = index + 1
        hasFeatures = False
        For Each group In plan
            If group("value") < longest * PLANAR_LOCAL_EDGE_RATIO Then
                Set a = group("first")
                Set b = group("last")
                If ((a("x") + b("x")) / 2 <= centerX And owner = PLANAR_LEFT_VIEW_INDEX) Or ((a("x") + b("x")) / 2 > centerX And owner = PLANAR_RIGHT_VIEW_INDEX) Then
                    hasFeatures = True
                End If
            End If
        Next group
        If hasFeatures Then
            Set view = MakePlanarEndDetail(d, endSource, plan, centerX, longest, owner)
            views.Add view
            ownerViews.Add CStr(owner), view
        End If
    Next index
    d.ForceRebuild3 False
    AsDrawing(d).ActivateSheet firstSheet
    CenterView first, sheetWidth * PLANAR_CENTER_X_RATIO, sheetHeight * PLANAR_FIRST_Y_RATIO
    For index = PLANAR_LEFT_VIEW_INDEX To PLANAR_RIGHT_VIEW_INDEX
        If Not ownerViews.Exists(CStr(index)) Then
            GoTo NextPlanarFamilyView
        End If
        AsDrawing(d).ActivateSheet detailSheet
        Set view = ownerViews(CStr(index))
        If spread Then
            angle = PLANAR_PAGE_UPPER_DETAIL_Y_RATIO
            If index = PLANAR_RIGHT_VIEW_INDEX Then
                angle = PLANAR_PAGE_LOWER_DETAIL_Y_RATIO
            End If
            CenterView view, sheetWidth * PLANAR_CENTER_X_RATIO, sheetHeight * angle
        Else
            CenterView view, sheetWidth * PLANAR_CENTER_X_RATIO, sheetHeight * (PLANAR_FIRST_Y_RATIO - (index - 1) * PLANAR_ROW_PITCH_RATIO)
        End If
        view.UpdateViewDisplayGeometry
NextPlanarFamilyView:
    Next index
    Set audit = fs.OpenTextFile(fs.BuildPath(runDir, PLANAR_CONSTRAINTS_FILENAME), FILE_FOR_APPENDING, True, FILE_UNICODE)
    For Each group In plan
        Set a = group("first")
        Set b = group("last")
        audit.WriteLine jobNumber & vbTab & group("id") & vbTab & group("value") & vbTab & group("dx") & vbTab & group("dy") & vbTab & group("count") & vbTab & group("selected")
        owner = PLANAR_OVERVIEW_INDEX
        If group("value") < longest * PLANAR_LOCAL_EDGE_RATIO Then
            owner = PLANAR_LEFT_VIEW_INDEX
            If (a("x") + b("x")) / 2 > centerX Then
                owner = PLANAR_RIGHT_VIEW_INDEX
            End If
        End If
        Set view = ownerViews(CStr(owner))
        If owner = PLANAR_OVERVIEW_INDEX Then
            AsDrawing(d).ActivateSheet firstSheet
        Else
            AsDrawing(d).ActivateSheet detailSheet
        End If
        If CBool(group("selected")) And owner <> PLANAR_OVERVIEW_INDEX Then
            AddPlanarEdgeDimension d, view, group, area
        End If
        angle = PlanarAngle(b("x") - a("x"), b("y") - a("y"))
        Do While angle < 0
            angle = angle + PI / 2
        Loop
        angle = angle - Int(angle / (PI / 2)) * (PI / 2)
        If angle > DIRECTION_ZERO_TOLERANCE And PI / 2 - angle > DIRECTION_ZERO_TOLERANCE Then
            key = CStr(Round(angle / PLANAR_ANGLE_TOLERANCE_RAD, 0))
            If Not directions.Exists(key) Then
                AddPlanarAngleDimension d, view, group, area
                directions.Add key, True
            End If
        End If
    Next group
    audit.Close
    AsDrawing(d).ActivateSheet firstSheet
    profileRoleDrawing = True
    normal = face.Normal
    box = body.GetBodyBox
    For index = SCREEN_X_AXIS_INDEX To SPACE_LAST_AXIS_INDEX
        thickness = thickness + Abs(normal(index)) * (box(index + BBOX_MAX_OFFSET) - box(index))
    Next index
    AddText d, "Товщина " & CStr(Round(thickness * MILLIMETRES_PER_METRE, LINEAR_DECIMAL_PLACES)) & " мм", CONTENT_LEFT_M, PROFILE_THICKNESS_NOTE_Y_M, TEXT_HEIGHT_STANDARD_M, 0
    Log "PLANAR FEATURE DRAWING views=" & views.Count & "; sheets=" & AsDrawing(d).GetSheetCount & "; native planned dimensions=" & planarDimensionRequests.Count & "; grouped hole diameters=" & rc & "; viewScaleValue=" & viewScaleValue
    MakePlanarFeatureDrawing = True
End Function
