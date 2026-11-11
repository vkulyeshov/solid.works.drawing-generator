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
    Dim datumPlan As Variant, anchor As Variant, polygon() As Variant, index As Long, visible As Variant
    Dim stage As String, errorNumber As Long, errorText As String
    On Error GoTo Failed
    stage = "edge lookup"
    Set a = group("first")
    Set b = group("last")
    Set first = PlanarVisibleLine(view, a("point"), b("point"))
    If first Is Nothing Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "An inclined planar contour edge is not visible."
    End If
    stage = "datum plan"
    p = PatternSheetPoint(view, a("point"))
    q = PatternSheetPoint(view, b("point"))
    dx = q(0) - p(0)
    dy = q(1) - p(1)
    length = Sqr(dx * dx + dy * dy)
    ReDim polygon(planarContourPoints.Count - 1)
    For index = 1 To planarContourPoints.Count
        polygon(index - 1) = PatternSheetPoint(view, planarContourPoints(index))
    Next index
    ' A contour clipped away by the view outline keeps the whole contour as the side test.
    visible = PlanarClipPolygon(polygon, view.GetOutline)
    If IsEmpty(visible) Then
        visible = polygon
    End If
    datumPlan = PlanarDatumPlan(p, q, visible)
    stage = "angle datum"
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
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    Err.Raise errorNumber, "AddPlanarAngleDimension", "Angle of group " & group("id") & " in " & view.Name & "; " & stage & ": " & errorText
End Sub

Private Function MakePlanarEndDetail(ByVal d As SldWorks.ModelDoc2, ByVal overview As SldWorks.View, ByVal plan As Collection, ByVal groupOwners As Object, ByVal owner As Long, ByVal cellWidth As Double, ByVal cellHeight As Double, ByVal band As Double, ByVal createX As Double, ByVal createY As Double, ByVal circles As Collection) As SldWorks.View
    Dim group As Object, a As Object, b As Object, anchor As Object, bounds As Variant, index As Long
    Dim p As Variant, x As Double, y As Double, radius As Double, profile As SldWorks.SketchSegment, view As SldWorks.View
    Dim selection As SldWorks.SelectData, parent As SldWorks.View
    Dim manager As SldWorks.SelectionMgr
    Dim stage As String, errorNumber As Long, errorText As String
    Dim detailScale As Double, outline As Variant, fitRatio As Double
    Dim label As String, labels As Variant, origin As Variant, minimumScale As Double
    On Error GoTo Failed
    bounds = PlanarOwnerBounds(plan, groupOwners, owner)
    For Each group In plan
        If groupOwners.Exists(CStr(group("id"))) Then
            If CLng(groupOwners(CStr(group("id")))) = owner Then
                Set anchor = group("first")
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
    detailScale = PlanarDetailScale(radius, MinD(cellWidth, cellHeight) - 2 * band)
    ' A detail stays enlarged; one that then overflows its cell asks the next attempt for smaller fragments.
    minimumScale = StandardScale(overview.ScaleDecimal * PLANAR_DETAIL_MIN_ENLARGEMENT)
    detailScale = MaxD(detailScale, minimumScale)
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
    ' The circle is stored relative to its view origin because the view may still be recentered.
    origin = overview.Position
    circles.Add Array(overview, AsDrawing(d).GetCurrentSheet.GetName, label, x - origin(0), y - origin(1), radius * overview.ScaleDecimal)
    Set view = AsDrawing(d).CreateDetailViewAt4(createX, createY, 0, SW_DETAIL_STANDARD, 1, 1 / detailScale, label, SW_DETAIL_CIRCLE, False, False, False, SW_DETAIL_OUTLINE_INTENSITY)
    If view Is Nothing Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot create an enlarged planar end detail."
    End If
    stage = "detail parent"
    Set parent = view.GetBaseView
    If parent Is Nothing Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "The planar detail has no source model view."
    End If
    Log "PLANAR DETAIL PARENT " & parent.Name
    view.DisableAutoUpdate = False
    view.UpdateViewDisplayGeometry
    d.GraphicsRedraw2
    d.ForceRebuild3 False
    d.ForceRebuild3 False
    ' Changing a detail's scale after creation shifts its crop away from the measured features,
    ' so a detail that still overflows its cell asks the next attempt for smaller fragments.
    outline = view.GetOutline
    fitRatio = MinD((cellHeight - 2 * band) / (outline(OUTLINE_TOP_INDEX) - outline(OUTLINE_BOTTOM_INDEX)), (cellWidth - 2 * band) / (outline(OUTLINE_RIGHT_INDEX) - outline(OUTLINE_LEFT_INDEX)))
    If fitRatio < 1 - PLANAR_DETAIL_OVERFLOW_TOLERANCE Then
        Log "PLANAR DETAIL OVERFLOW " & view.Name & "; ratio=" & fitRatio
        RequestPlanarSplit "detail " & label, PlanarOwnerGroupIds(groupOwners, owner)
    End If
    planarDetailLabels.Add view.Name, label
    Log "PLANAR DETAIL " & view.Name & "; scale=" & view.ScaleDecimal
    Set MakePlanarEndDetail = view
    Exit Function
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    Err.Raise errorNumber, "MakePlanarEndDetail", stage & ": " & errorText
End Function

Private Function PlanarOwnerHasFeatures(ByVal groupOwners As Object, ByVal owner As Long) As Boolean
    Dim item As Variant
    For Each item In groupOwners.Items
        If CLng(item) = owner Then
            PlanarOwnerHasFeatures = True
            Exit Function
        End If
    Next item
End Function

Private Function PlanarOwnerGroupIds(ByVal groupOwners As Object, ByVal owner As Long) As String
    Dim key As Variant, result As String
    For Each key In groupOwners.Keys
        If CLng(groupOwners(key)) = owner Then
            If Len(result) > 0 Then
                result = result & ","
            End If
            result = result & CStr(key)
        End If
    Next key
    PlanarOwnerGroupIds = result
End Function

Private Function PlanarOwnerBounds(ByVal plan As Collection, ByVal groupOwners As Object, ByVal owner As Long) As Variant
    Dim group As Object, a As Object, b As Object, bounds As Variant
    bounds = Array(SEARCH_INITIAL_LIMIT, SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT)
    For Each group In plan
        If groupOwners.Exists(CStr(group("id"))) Then
            If CLng(groupOwners(CStr(group("id")))) = owner Then
                Set a = group("first")
                Set b = group("last")
                NativeExpandBounds bounds, a("x"), a("y")
                NativeExpandBounds bounds, b("x"), b("y")
            End If
        End If
    Next group
    PlanarOwnerBounds = bounds
End Function

' Largest standard scale, at most the detail scale limit, whose detail outline fits the given paper size.
Private Function PlanarDetailScale(ByVal radius As Double, ByVal paper As Double) As Double
    PlanarDetailScale = MinD(PLANAR_DETAIL_SCALE, StandardScale(paper / ((2 * radius + PLANAR_DETAIL_BOUNDARY_MARGIN_M) * PLANAR_DETAIL_OUTLINE_ALLOWANCE)))
End Function

' Largest detail first, each at the highest and then leftmost free spot of a sheet: candidate corners are
' the area corner and the corners of placed cells. Page 0 is the area under the overview on the first
' sheet, where the thickness notes keep the lower-left corner.
Private Function PackPlanarDetails(ByVal sizes As Object, ByVal firstTop As Double, ByVal placements As Object) As Long
    Dim keys As Variant, order() As String, count As Long, index As Long, other As Long, swapKey As String
    Dim cells As Collection, cell As Variant, xs As Collection, ys As Collection, x As Variant, y As Variant
    Dim left As Double, right As Double, top As Double, bottom As Double, size As Double, page As Long, lastPage As Long
    Dim bestX As Double, bestY As Double, found As Boolean, box As Variant, free As Boolean
    left = LEADER_WORK_LEFT_M
    right = sheetWidth - LEADER_WORK_EDGE_MARGIN_M
    keys = sizes.Keys
    count = sizes.Count
    If count = 0 Then
        Exit Function
    End If
    ReDim order(count - 1)
    For index = 0 To count - 1
        order(index) = keys(index)
    Next index
    For index = 1 To count - 1
        For other = index To 1 Step -1
            If sizes(order(other)) > sizes(order(other - 1)) Then
                swapKey = order(other)
                order(other) = order(other - 1)
                order(other - 1) = swapKey
            End If
        Next other
    Next index
    Set cells = New Collection
    cells.Add Array(0, left, LEADER_WORK_BOTTOM_M, left + PLANAR_NOTE_BLOCK_WIDTH_M, LEADER_WORK_BOTTOM_M + PLANAR_OVERVIEW_NOTE_RESERVE_M)
    For index = 0 To count - 1
        size = sizes(order(index))
        found = False
        For page = 0 To lastPage + 1
            top = firstTop
            bottom = LEADER_WORK_BOTTOM_M
            If page > 0 Then
                top = PlanarDetailCellTop()
                bottom = LAYOUT_CONTINUATION_BOTTOM_M
            End If
            Set xs = New Collection
            Set ys = New Collection
            xs.Add left
            ys.Add top
            For Each cell In cells
                If cell(0) = page Then
                    xs.Add cell(3)
                    ys.Add cell(2)
                End If
            Next cell
            For Each y In ys
                For Each x In xs
                    box = Array(x, y - size, x + size, y)
                    free = box(0) >= left - PROFILE_MATERIAL_TOLERANCE_M And box(2) <= right And box(1) >= bottom And box(3) <= top + PROFILE_MATERIAL_TOLERANCE_M
                    If free Then
                        For Each cell In cells
                            If cell(0) = page Then
                                If NativeBoxesOverlap(box, Array(cell(1), cell(2), cell(3), cell(4)), -PROFILE_MATERIAL_TOLERANCE_M) Then
                                    free = False
                                    Exit For
                                End If
                            End If
                        Next cell
                    End If
                    If free Then
                        If Not found Or y > bestY + PROFILE_MATERIAL_TOLERANCE_M Or (Abs(y - bestY) <= PROFILE_MATERIAL_TOLERANCE_M And x < bestX) Then
                            bestX = x
                            bestY = y
                            found = True
                        End If
                    End If
                Next x
            Next y
            If found Then
                Exit For
            End If
        Next page
        If page > lastPage Then
            lastPage = page
        End If
        cells.Add Array(page, bestX, bestY - size, bestX + size, bestY)
        placements.Add order(index), Array(page, bestX + size / 2, bestY - size / 2)
    Next index
    PackPlanarDetails = lastPage
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

' Every local measurement belongs to one enlarged detail. Ordinary profiles keep
' one detail per end. Dense profiles give one detail to the local measurements within
' PLANAR_DETAIL_MAX_SPAN_M along the overview, ordered from the part end; a detail that an
' earlier attempt found crowded is cut into equal runs.
' Detail letters go to the nearest free spot outside the part contour, the detail circles,
' the other views with their dimension offsets and the letters already placed on the sheet.
Private Sub PlacePlanarDetailLabels(ByVal d As SldWorks.ModelDoc2, ByVal circles As Collection)
    Dim record As Variant, other As Variant, view As SldWorks.View, item As Variant, detailCircle As SldWorks.DetailCircle, found As SldWorks.DetailCircle
    Dim origin As Variant, cx As Double, cy As Double, radius As Double, lx As Double, ly As Double, height As Double, width As Double
    Dim polygon As Variant, index As Long, obstacles As Collection, placed As Object, sheetName As String, candidate As SldWorks.View
    Dim direction As Long, angle As Double, distance As Double, px As Double, py As Double, box As Variant, free As Boolean
    Dim best As Double, bestX As Double, bestY As Double, bestBox As Variant, outline As Variant, otherOrigin As Variant
    Set placed = CreateObject("Scripting.Dictionary")
    For Each record In circles
        Set view = record(0)
        sheetName = record(1)
        AsDrawing(d).ActivateSheet sheetName
        Set found = Nothing
        item = view.GetDetailCircles
        If Not IsEmpty(item) Then
            For Each other In item
                Set detailCircle = other
                If detailCircle.GetLabel = record(2) Then
                    Set found = detailCircle
                End If
            Next other
        End If
        If found Is Nothing Then
            Log "PLANAR DETAIL LABEL missing circle " & record(2)
            GoTo NextPlanarDetailLabel
        End If
        origin = view.Position
        cx = origin(0) + record(3)
        cy = origin(1) + record(4)
        radius = record(5)
        found.GetLabelPosition lx, ly
        ' The label point must be in sheet space; otherwise keep the native placement.
        If Sqr((lx - cx) ^ 2 + (ly - cy) ^ 2) > radius + PLANAR_DETAIL_LABEL_REACH_M Then
            Log "PLANAR DETAIL LABEL kept " & record(2) & "; native=" & lx & "," & ly & "; circle=" & cx & "," & cy & "," & radius
            GoTo NextPlanarDetailLabel
        End If
        height = TEXT_HEIGHT_STANDARD_M
        If Not found.GetTextFormat Is Nothing Then
            If found.GetTextFormat.CharHeight > 0 Then
                height = found.GetTextFormat.CharHeight
            End If
        End If
        ' The anchor of the label point is not documented, so the box covers it on every side.
        width = MaxD(height, height * LEADER_GLYPH_WIDTH_RATIO * Len(record(2)))
        ReDim polygon(planarContourPoints.Count - 1)
        For index = 1 To planarContourPoints.Count
            polygon(index - 1) = PatternSheetPoint(view, planarContourPoints(index))
        Next index
        Set obstacles = New Collection
        obstacles.Add Array(LEADER_WORK_LEFT_M, sheetHeight - LEADER_STAMP_TOP_MARGIN_M, LEADER_STAMP_WIDTH_M, sheetHeight)
        Set candidate = AsDrawing(d).GetFirstView.GetNextView
        Do While Not candidate Is Nothing
            If candidate.Name <> view.Name Then
                outline = candidate.GetOutline
                obstacles.Add Array(outline(OUTLINE_LEFT_INDEX) - GOST_OBJECT_OFFSET_M, outline(OUTLINE_BOTTOM_INDEX) - GOST_OBJECT_OFFSET_M, outline(OUTLINE_RIGHT_INDEX) + GOST_OBJECT_OFFSET_M, outline(OUTLINE_TOP_INDEX) + GOST_OBJECT_OFFSET_M)
            End If
            Set candidate = candidate.GetNextView
        Loop
        If placed.Exists(sheetName) Then
            For Each item In placed(sheetName)
                obstacles.Add item
            Next item
        End If
        best = SEARCH_INITIAL_LIMIT
        For direction = 0 To PLANAR_DETAIL_LABEL_DIRECTIONS - 1
            angle = 2 * PI * direction / PLANAR_DETAIL_LABEL_DIRECTIONS
            For distance = radius + height To radius + PLANAR_DETAIL_LABEL_REACH_M Step height / 2
                px = cx + distance * Cos(angle)
                py = cy + distance * Sin(angle)
                box = Array(px - width, py - height, px + width, py + height)
                free = box(0) > LEADER_WORK_LEFT_M And box(2) < sheetWidth - LEADER_WORK_EDGE_MARGIN_M And box(1) > LEADER_WORK_BOTTOM_M And box(3) < sheetHeight - LEADER_WORK_EDGE_MARGIN_M
                If free Then
                    free = Not PlanarPolygonsOverlap(Array(Array(box(0), box(1)), Array(box(2), box(1)), Array(box(2), box(3)), Array(box(0), box(3))), polygon)
                End If
                If free Then
                    For Each item In obstacles
                        If NativeBoxesOverlap(box, item, LEADER_LABEL_CLEARANCE_M) Then
                            free = False
                            Exit For
                        End If
                    Next item
                End If
                If free Then
                    For Each other In circles
                        If other(1) = sheetName Then
                            otherOrigin = other(0).Position
                            If PlanarBoxCircleDistance(box, otherOrigin(0) + other(3), otherOrigin(1) + other(4)) < other(5) + LEADER_LABEL_CLEARANCE_M Then
                                free = False
                                Exit For
                            End If
                        End If
                    Next other
                End If
                If free Then
                    If distance < best Then
                        best = distance
                        bestX = px
                        bestY = py
                        bestBox = box
                    End If
                    Exit For
                End If
            Next distance
        Next direction
        If best = SEARCH_INITIAL_LIMIT Then
            Log "PLANAR DETAIL LABEL no free place for " & record(2)
            GoTo NextPlanarDetailLabel
        End If
        found.SetLabelPosition bestX, bestY
        If Not placed.Exists(sheetName) Then
            placed.Add sheetName, New Collection
        End If
        placed(sheetName).Add bestBox
        If Not planarDetailLabelBoxes.Exists(view.Name) Then
            planarDetailLabelBoxes.Add view.Name, New Collection
        End If
        planarDetailLabelBoxes(view.Name).Add bestBox
        Log "PLANAR DETAIL LABEL " & record(2) & "; native=" & lx & "," & ly & "; placed=" & bestX & "," & bestY
NextPlanarDetailLabel:
    Next record
    d.GraphicsRedraw2
End Sub

Private Function PlanarBoxCircleDistance(ByVal box As Variant, ByVal x As Double, ByVal y As Double) As Double
    Dim nearestX As Double, nearestY As Double
    nearestX = MinD(MaxD(x, box(0)), box(2))
    nearestY = MinD(MaxD(y, box(1)), box(3))
    PlanarBoxCircleDistance = Sqr((nearestX - x) ^ 2 + (nearestY - y) ^ 2)
End Function

Private Function AddPlanarHoleDimensions(ByVal d As SldWorks.ModelDoc2, ByVal view As SldWorks.View, ByVal holes As Collection, ByVal reference As Object) As Long
    Dim display As SldWorks.DisplayDimension
    AddPlanarHoleLocations d, view, holes, reference
    AddPlanarHoleDimensions = DiameterDimensions(d, view)
    Set display = view.GetFirstDisplayDimension5
    Do While Not display Is Nothing
        If display.Type2 = SW_DIMENSION_DIAMETER Then
            display.SetSecondArrow False, False
        End If
        Set display = display.GetNext5
    Loop
End Function

Private Function PlanarSpreadHeight() As Double
    PlanarSpreadHeight = sheetHeight - LEADER_STAMP_TOP_MARGIN_M - LEADER_WORK_BOTTOM_M - PLANAR_OVERVIEW_NOTE_RESERVE_M
End Function

Private Function PlanarSpreadCenterY() As Double
    PlanarSpreadCenterY = LEADER_WORK_BOTTOM_M + PLANAR_OVERVIEW_NOTE_RESERVE_M + PlanarSpreadHeight() / 2
End Function

' Each split level requested by a failed attempt halves the measurements per detail.
Private Sub LoadPlanarSplit()
    Dim path As String, stream As Object, fields As Variant, id As Variant
    planarSplitOverview = False
    planarSplitPending = False
    Set planarSplitGroups = CreateObject("Scripting.Dictionary")
    Set planarSplitRequested = CreateObject("Scripting.Dictionary")
    path = fs.BuildPath(runDir, PLANAR_SPLIT_FILENAME)
    If Not fs.FileExists(path) Then
        Exit Sub
    End If
    Set stream = fs.OpenTextFile(path, FILE_FOR_READING, False, FILE_UNICODE)
    Do While Not stream.AtEndOfStream
        fields = Split(stream.ReadLine, vbTab)
        If UBound(fields) >= 2 Then
            If CLng(fields(0)) = jobNumber Then
                If fields(1) = "1" Then
                    planarSplitOverview = True
                End If
                ' Each request names the measurements of one crowded detail; repeated requests split it finer.
                For Each id In Split(fields(2), ",")
                    If Len(id) > 0 Then
                        planarSplitGroups(CStr(id)) = CLng(planarSplitGroups(CStr(id))) + 1
                    End If
                Next id
            End If
        End If
    Loop
    stream.Close
    If planarSplitOverview Or planarSplitGroups.Count > 0 Then
        Log "PLANAR SPLIT loaded; overview=" & planarSplitOverview & "; flagged measurements=" & planarSplitGroups.Count
    End If
End Sub

Private Function PlanarViewGroupIds(ByVal viewName As String) As String
    Dim request As Variant, seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    For Each request In planarDimensionRequests.Items
        If request("view") = viewName Then
            seen(CStr(request("source"))) = True
        End If
    Next request
    PlanarViewGroupIds = Join(seen.Keys, ",")
End Function

' While attempts remain, the next attempt thins a crowded overview or splits only the crowded detail.
Private Sub RequestPlanarSplit(ByVal viewName As String, Optional ByVal ids As String = "")
    Dim overview As String, stream As Object
    If jobAttempt >= maxJobAttempts Or planarSplitRequested.Exists(viewName) Then
        Exit Sub
    End If
    overview = "0"
    If viewName = planarOverviewName Then
        If planarSplitOverview Then
            Exit Sub
        End If
        planarSplitOverview = True
        overview = "1"
    Else
        If Len(ids) = 0 Then
            ids = PlanarViewGroupIds(viewName)
        End If
        If Len(ids) = 0 Then
            Exit Sub
        End If
    End If
    planarSplitRequested.Add viewName, True
    planarSplitPending = True
    Set stream = fs.OpenTextFile(fs.BuildPath(runDir, PLANAR_SPLIT_FILENAME), FILE_FOR_APPENDING, True, FILE_UNICODE)
    stream.WriteLine jobNumber & vbTab & overview & vbTab & ids
    stream.Close
    Log "PLANAR SPLIT requested for the next attempt by " & viewName & "; overview=" & (overview = "1") & "; measurements=" & ids
End Sub

Private Function PlanarDetailCellTop() As Double
    PlanarDetailCellTop = sheetHeight - FAMILY_REFERENCE_CENTER_MARGIN_M - FAMILY_REFERENCE_HEIGHT_M / 2
End Function

' Edge lengths kept on the overview plus two location dimensions per hole.
Private Function PlanarOverviewDimensionCount(ByVal plan As Collection, ByVal longest As Double, ByVal overviewEdgeScale As Double, ByVal holes As Collection) As Long
    Dim group As Object
    For Each group In plan
        If CBool(group("selected")) And PlanarKeepsOnOverview(group, longest, overviewEdgeScale) Then
            PlanarOverviewDimensionCount = PlanarOverviewDimensionCount + 1
        End If
    Next group
    PlanarOverviewDimensionCount = PlanarOverviewDimensionCount + 2 * holes.Count
End Function

Private Function PlanarKeepsOnOverview(ByVal group As Object, ByVal longest As Double, ByVal overviewEdgeScale As Double) As Boolean
    PlanarKeepsOnOverview = group("value") >= longest * PLANAR_LOCAL_EDGE_RATIO Or group("value") * overviewEdgeScale >= PLANAR_OVERVIEW_MIN_EDGE_M
End Function

Private Function AssignPlanarDetailOwners(ByVal plan As Collection, ByVal longest As Double, ByVal overviewEdgeScale As Double, ByVal centerX As Double, ByVal dense As Boolean, ByVal groupOwners As Object) As Long
    Dim group As Object, a As Object, b As Object, side As Long, count As Long, index As Long, other As Long
    Dim keys() As Double, ids() As String, swapKey As Double, swapId As String, middle As Double, owner As Long, used As Long, clusterStart As Double
    Dim heights() As Double, clusterOwners() As Long, cluster As Long, firstOwner As Long, members As Long, order() As Long, swapIndex As Long, ownerCount As Long
    Dim lastCluster As Long, flagged As Long, limit As Long, firstPiece As Long
    owner = PLANAR_LEFT_VIEW_INDEX - 1
    For side = 0 To 1
        count = 0
        ReDim keys(plan.Count)
        ReDim ids(plan.Count)
        ReDim heights(plan.Count)
        ReDim clusterOwners(plan.Count)
        For Each group In plan
            If Not PlanarKeepsOnOverview(group, longest, overviewEdgeScale) Then
                Set a = group("first")
                Set b = group("last")
                middle = (a("x") + b("x")) / 2
                If (side = 0 And middle <= centerX) Or (side = 1 And middle > centerX) Then
                    keys(count) = middle
                    If side = 1 Then
                        keys(count) = -middle
                    End If
                    ids(count) = CStr(group("id"))
                    heights(count) = (a("y") + b("y")) / 2
                    count = count + 1
                End If
            End If
        Next group
        For index = 1 To count - 1
            For other = index To 1 Step -1
                If keys(other) < keys(other - 1) Then
                    swapKey = keys(other)
                    keys(other) = keys(other - 1)
                    keys(other - 1) = swapKey
                    swapId = ids(other)
                    ids(other) = ids(other - 1)
                    ids(other - 1) = swapId
                    swapKey = heights(other)
                    heights(other) = heights(other - 1)
                    heights(other - 1) = swapKey
                End If
            Next other
        Next index
        If dense Then
            used = 0
            firstOwner = owner + 1
            For index = 0 To count - 1
                If used = 0 Or keys(index) - clusterStart > PLANAR_DETAIL_MAX_SPAN_M Then
                    owner = owner + 1
                    used = 0
                    clusterStart = keys(index)
                End If
                clusterOwners(index) = owner
                used = used + 1
            Next index
            ' Features at one position along the part but on opposite long sides, separated across the
            ' part by more than the span, get separate details instead of one mostly empty circle.
            ReDim order(count)
            For cluster = firstOwner To owner
                members = 0
                For index = 0 To count - 1
                    If clusterOwners(index) = cluster Then
                        order(members) = index
                        members = members + 1
                    End If
                Next index
                For index = 1 To members - 1
                    For other = index To 1 Step -1
                        If heights(order(other)) < heights(order(other - 1)) Then
                            swapIndex = order(other)
                            order(other) = order(other - 1)
                            order(other - 1) = swapIndex
                        End If
                    Next other
                Next index
                used = cluster
                For index = 0 To members - 1
                    If index > 0 Then
                        If heights(order(index)) - heights(order(index - 1)) > PLANAR_DETAIL_MAX_SPAN_M Then
                            ownerCount = ownerCount + 1
                            used = owner + ownerCount
                        End If
                    End If
                    clusterOwners(order(index)) = used
                Next index
            Next cluster
            ' A detail flagged as crowded by an earlier attempt is cut into 2, 4, ... equal runs along the part.
            lastCluster = owner + ownerCount
            For cluster = firstOwner To lastCluster
                members = 0
                flagged = 0
                For index = 0 To count - 1
                    If clusterOwners(index) = cluster Then
                        members = members + 1
                        If planarSplitGroups.Exists(ids(index)) Then
                            If CLng(planarSplitGroups(ids(index))) > flagged Then
                                flagged = CLng(planarSplitGroups(ids(index)))
                            End If
                        End If
                    End If
                Next index
                If flagged > 0 And members > PLANAR_DETAIL_MIN_GROUP_LIMIT Then
                    limit = -Int(-members / 2 ^ flagged)
                    If limit < PLANAR_DETAIL_MIN_GROUP_LIMIT Then
                        limit = PLANAR_DETAIL_MIN_GROUP_LIMIT
                    End If
                    used = 0
                    firstPiece = owner + ownerCount
                    For index = 0 To count - 1
                        If clusterOwners(index) = cluster Then
                            If used \ limit > 0 Then
                                clusterOwners(index) = firstPiece + used \ limit
                            End If
                            used = used + 1
                        End If
                    Next index
                    ownerCount = ownerCount + (used - 1) \ limit
                End If
            Next cluster
            owner = owner + ownerCount
            ownerCount = 0
            For index = 0 To count - 1
                groupOwners.Add ids(index), clusterOwners(index)
            Next index
        Else
            ' The right-end detail keeps its fixed owner even without a left detail.
            owner = PLANAR_LEFT_VIEW_INDEX + side
            For index = 0 To count - 1
                groupOwners.Add ids(index), owner
            Next index
        End If
    Next side
    AssignPlanarDetailOwners = owner
End Function

Private Function MakePlanarFeatureDrawing(ByRef d As SldWorks.ModelDoc2, ByVal doc As SldWorks.ModelDoc2, ByVal cfg As String, ByVal body As SldWorks.Body2, ByVal outName As String) As Boolean
    Dim face As SldWorks.Face2, raw As Collection, holes As Collection, projected As New Collection, projectedHoles As New Collection
    Dim boundary As SldWorks.Loop2, item As Variant, loops As Variant, node As Object, hole As Object, points As Collection
    Dim first As SldWorks.View, view As SldWorks.View, views As New Collection, plan As Collection, group As Object, reference As Object
    Dim a As Object, b As Object, longest As Double, area As Double, angle As Double, outline As Variant, viewScaleValue As Double
    Dim index As Long, owner As Long, centerX As Double, directions As Object, key As String, rc As Long, display As SldWorks.DisplayDimension
    Dim request As Object, audit As Object, normal As Variant, box As Variant, thickness As Double, angleReference As Object
    Dim error As Long, warning As Long, nativePath As String
    Dim ownerViews As Object, endSource As SldWorks.View, referenceView As SldWorks.View
    Dim firstSheet As String, detailSheet As String, spread As Boolean, localCount As Long
    Dim groupOwners As Object, lastOwner As Long, ownerSheets As Object, cellHeight As Double, createY As Double, detailPage As Long
    Dim circles As New Collection, holeView As SldWorks.View, overviewY As Double, detailBand As Double, overviewEdgeScale As Double
    Dim placements As Object, pageSheets As Object, pageSources As Object, sizes As Object, placement As Variant, lastPage As Long
    Dim packTop As Double, cellWidth As Double, createX As Double, radius As Double, detailScale As Double
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
    Set planarDetailLabelBoxes = CreateObject("Scripting.Dictionary")
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
    LoadPlanarSplit
    ' A split requested by an earlier attempt always moves local features to fragment sheets.
    spread = localCount > PLANAR_PAGE_LOCAL_DIMENSION_LIMIT Or planarSplitOverview Or planarSplitGroups.Count > 0
    If spread Then
        ' The overview sits at the top of the first sheet; the details are packed into the space below it.
        viewScaleValue = FitScale(first, sheetWidth * PLANAR_WIDTH_RATIO, sheetHeight * PLANAR_PACK_OVERVIEW_HEIGHT_RATIO)
    Else
        viewScaleValue = FitScale(first, sheetWidth * PLANAR_WIDTH_RATIO, sheetHeight * PLANAR_HEIGHT_RATIO)
    End If
    SetPaperScale AsDrawing(d).GetCurrentSheet, viewScaleValue
    first.UseSheetScale = True
    d.ForceRebuild3 False
    overviewY = sheetHeight * PLANAR_FIRST_Y_RATIO
    If spread Then
        outline = first.GetOutline
        overviewY = sheetHeight - LEADER_STAMP_TOP_MARGIN_M - PLANAR_PACK_OVERVIEW_RESERVE_M - (outline(OUTLINE_TOP_INDEX) - outline(OUTLINE_BOTTOM_INDEX)) / 2
        packTop = overviewY - (outline(OUTLINE_TOP_INDEX) - outline(OUTLINE_BOTTOM_INDEX)) / 2 - PLANAR_PACK_OVERVIEW_RESERVE_M
    End If
    CenterView first, sheetWidth * PLANAR_CENTER_X_RATIO, overviewY
    ' On fragment sheets the overview also keeps every edge that is readable at its own scale.
    overviewEdgeScale = 0
    ' An overview flagged as crowded keeps only its long edges; the rest moves to details.
    If spread And Not planarSplitOverview Then
        overviewEdgeScale = viewScaleValue
        ' An overview that is crowded by count is thinned at once instead of after a failed attempt.
        If PlanarOverviewDimensionCount(plan, longest, overviewEdgeScale, projectedHoles) > PLANAR_OVERVIEW_DIMENSION_LIMIT Then
            overviewEdgeScale = 0
            planarSplitOverview = True
            Log "PLANAR OVERVIEW crowded by count; short edges go to details and holes to their own sheet"
        End If
    End If
    For Each group In plan
        If PlanarKeepsOnOverview(group, longest, overviewEdgeScale) And CBool(group("selected")) Then
            AddPlanarEdgeDimension d, first, group, area
        End If
    Next group
    If Not planarSplitOverview Then
        rc = AddPlanarHoleDimensions(d, first, projectedHoles, reference)
    End If
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
    planarOverviewName = first.Name
    If planarSplitOverview And projectedHoles.Count > 0 Then
        ' The crowded overview keeps the contour; holes get their own full view on the next sheet.
        NewFormattedSheet d, "Отвори"
        Set holeView = MakeView(d, doc, cfg, body, first.GetOrientationName, sheetWidth * PLANAR_CENTER_X_RATIO, PlanarSpreadCenterY())
        holeView.Angle = first.Angle
        holeView.UseSheetScale = False
        holeView.ScaleDecimal = 1
        d.ForceRebuild3 False
        holeView.ScaleDecimal = FitScale(holeView, sheetWidth * PLANAR_WIDTH_RATIO, PlanarSpreadHeight() * PLANAR_WIDTH_RATIO)
        d.ForceRebuild3 False
        CenterView holeView, sheetWidth * PLANAR_CENTER_X_RATIO, PlanarSpreadCenterY()
        rc = AddPlanarHoleDimensions(d, holeView, projectedHoles, reference)
        views.Add holeView
        Log "PLANAR SPLIT holes moved to " & holeView.Name & "; scale=" & holeView.ScaleDecimal
        AsDrawing(d).ActivateSheet firstSheet
    End If
    Set ownerViews = CreateObject("Scripting.Dictionary")
    ownerViews.Add CStr(PLANAR_OVERVIEW_INDEX), first
    views.Add first
    Set endSource = first
    Set groupOwners = CreateObject("Scripting.Dictionary")
    Set ownerSheets = CreateObject("Scripting.Dictionary")
    lastOwner = AssignPlanarDetailOwners(plan, longest, overviewEdgeScale, centerX, spread, groupOwners)
    cellHeight = sheetHeight * PLANAR_ROW_PITCH_RATIO
    cellWidth = sheetWidth - LEADER_WORK_LEFT_M - LEADER_WORK_EDGE_MARGIN_M
    detailBand = GOST_OBJECT_OFFSET_M
    Set placements = CreateObject("Scripting.Dictionary")
    Set pageSheets = CreateObject("Scripting.Dictionary")
    Set pageSources = CreateObject("Scripting.Dictionary")
    pageSheets.Add CStr(0), firstSheet
    pageSources.Add CStr(0), first
    If spread Then
        ' Each detail gets a square cell of its enlarged circle plus a dimension band on every side.
        detailBand = PLANAR_DETAIL_DIMENSION_BAND_M
        Set sizes = CreateObject("Scripting.Dictionary")
        For owner = PLANAR_LEFT_VIEW_INDEX To lastOwner
            If PlanarOwnerHasFeatures(groupOwners, owner) Then
                radius = DetailProfileRadius(PlanarOwnerBounds(plan, groupOwners, owner))
                detailScale = MaxD(PlanarDetailScale(radius, PLANAR_DETAIL_MAX_PAPER_M), StandardScale(first.ScaleDecimal * PLANAR_DETAIL_MIN_ENLARGEMENT))
                sizes.Add CStr(owner), (2 * radius + PLANAR_DETAIL_BOUNDARY_MARGIN_M) * PLANAR_DETAIL_OUTLINE_ALLOWANCE * detailScale + 2 * detailBand
            End If
        Next owner
        lastPage = PackPlanarDetails(sizes, packTop, placements)
        For detailPage = 1 To lastPage
            ' Each continuation sheet has its own reference view for the detail circles.
            If detailPage = 1 Then
                NewFormattedSheet d, "Місцеві елементи"
            Else
                NewFormattedSheet d, "Місцеві елементи " & detailPage
            End If
            Set referenceView = MakeView(d, doc, cfg, body, first.GetOrientationName, sheetWidth * PLANAR_CENTER_X_RATIO, sheetHeight - FAMILY_REFERENCE_CENTER_MARGIN_M)
            referenceView.Angle = first.Angle
            referenceView.UseSheetScale = False
            referenceView.ScaleDecimal = 1
            d.ForceRebuild3 False
            referenceView.ScaleDecimal = FitScale(referenceView, sheetWidth - LEADER_WORK_LEFT_M - LEADER_WORK_EDGE_MARGIN_M, FAMILY_REFERENCE_HEIGHT_M)
            d.ForceRebuild3 False
            CenterView referenceView, sheetWidth * PLANAR_CENTER_X_RATIO, sheetHeight - FAMILY_REFERENCE_CENTER_MARGIN_M
            pageSheets.Add CStr(detailPage), AsDrawing(d).GetCurrentSheet.GetName
            pageSources.Add CStr(detailPage), referenceView
        Next detailPage
        Log "PLANAR DETAIL PACK details=" & sizes.Count & "; continuation sheets=" & lastPage
    End If
    For owner = PLANAR_LEFT_VIEW_INDEX To lastOwner
        If Not PlanarOwnerHasFeatures(groupOwners, owner) Then
            GoTo NextPlanarDetailOwner
        End If
        createX = sheetWidth * PLANAR_CENTER_X_RATIO
        createY = sheetHeight * (PLANAR_FIRST_Y_RATIO - (owner - 1) * PLANAR_ROW_PITCH_RATIO)
        detailPage = 0
        If spread Then
            placement = placements(CStr(owner))
            detailPage = placement(0)
            createX = placement(1)
            createY = placement(2)
            cellWidth = sizes(CStr(owner))
            cellHeight = cellWidth
        End If
        detailSheet = pageSheets(CStr(detailPage))
        Set endSource = pageSources(CStr(detailPage))
        AsDrawing(d).ActivateSheet detailSheet
        Set view = MakePlanarEndDetail(d, endSource, plan, groupOwners, owner, cellWidth, cellHeight, detailBand, createX, createY, circles)
        views.Add view
        ownerViews.Add CStr(owner), view
        ownerSheets.Add CStr(owner), detailSheet
NextPlanarDetailOwner:
    Next owner
    d.ForceRebuild3 False
    AsDrawing(d).ActivateSheet firstSheet
    CenterView first, sheetWidth * PLANAR_CENTER_X_RATIO, overviewY
    For index = PLANAR_LEFT_VIEW_INDEX To lastOwner
        If Not ownerViews.Exists(CStr(index)) Then
            GoTo NextPlanarFamilyView
        End If
        AsDrawing(d).ActivateSheet ownerSheets(CStr(index))
        Set view = ownerViews(CStr(index))
        If spread Then
            placement = placements(CStr(index))
            CenterView view, placement(1), placement(2)
        Else
            CenterView view, sheetWidth * PLANAR_CENTER_X_RATIO, sheetHeight * (PLANAR_FIRST_Y_RATIO - (index - 1) * PLANAR_ROW_PITCH_RATIO)
        End If
        view.UpdateViewDisplayGeometry
NextPlanarFamilyView:
    Next index
    PlacePlanarDetailLabels d, circles
    Set audit = fs.OpenTextFile(fs.BuildPath(runDir, PLANAR_CONSTRAINTS_FILENAME), FILE_FOR_APPENDING, True, FILE_UNICODE)
    For Each group In plan
        Set a = group("first")
        Set b = group("last")
        audit.WriteLine jobNumber & vbTab & group("id") & vbTab & group("value") & vbTab & group("dx") & vbTab & group("dy") & vbTab & group("count") & vbTab & group("selected")
        owner = PLANAR_OVERVIEW_INDEX
        If groupOwners.Exists(CStr(group("id"))) Then
            owner = CLng(groupOwners(CStr(group("id"))))
        End If
        If Not ownerViews.Exists(CStr(owner)) Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "No planar view owns group " & group("id") & "; owner=" & owner
        End If
        Set view = ownerViews(CStr(owner))
        If owner = PLANAR_OVERVIEW_INDEX Then
            AsDrawing(d).ActivateSheet firstSheet
        Else
            AsDrawing(d).ActivateSheet ownerSheets(CStr(owner))
        End If
        If CBool(group("selected")) And owner <> PLANAR_OVERVIEW_INDEX Then
            Log "PLANAR LOCAL EDGE group=" & group("id") & "; view=" & view.Name
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
