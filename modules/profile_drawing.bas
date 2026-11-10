Attribute VB_Name = "profile_drawing"
Option Explicit

Private Function ProfileReference(ByVal view As SldWorks.View, ByVal point As Variant) As SldWorks.Vertex
    Dim components As Variant, component As Variant, edges As Variant, item As Variant, edge As SldWorks.Edge
    Dim vertex As SldWorks.Vertex, endpoint As Long, target As Variant, actual As Variant
    target = PatternSheetPoint(view, point)
    components = view.GetVisibleComponents
    For Each component In components
        edges = view.GetVisibleEntities2(component, SW_VISIBLE_ENTITY_EDGE)
        For Each item In edges
            Set edge = item
            For endpoint = EDGE_ENDPOINT_FIRST To EDGE_ENDPOINT_LAST
                Set vertex = edge.GetStartVertex
                If endpoint = EDGE_ENDPOINT_LAST Then
                    Set vertex = edge.GetEndVertex
                End If
                If Not vertex Is Nothing Then
                    actual = PatternSheetPoint(view, vertex.GetPoint)
                    If Abs(actual(0) - target(0)) < PROFILE_COORDINATE_TOLERANCE_M * view.ScaleDecimal And Abs(actual(1) - target(1)) < PROFILE_COORDINATE_TOLERANCE_M * view.ScaleDecimal Then
                        Set ProfileReference = vertex
                        Exit Function
                    End If
                End If
            Next endpoint
        Next item
    Next component
    If profilePagedMode Then
        If profilePageCrops.Exists(view.Name) Then
            Set ProfileReference = ProfilePageReference(view, point)
        End If
    End If
End Function

Private Sub ReserveProfileDimensionSpace(ByVal d As SldWorks.ModelDoc2, ByVal views As Collection, ByVal work As Variant)
    Dim item As Variant, view As SldWorks.View, display As SldWorks.DisplayDimension, outline As Variant, region As Variant
    Dim lane As Long, maximum As Long, required As Double, available As Double, shift As Double, headroom As Double, position As Variant
    For Each item In views
        Set view = item
        ' Panel bases and labels are fitted to the isolated region later.
        ' Initial construction lanes must not move the entire view.
        If PanelOwnsDimension(view.GetFirstDisplayDimension5) Then
            GoTo NextProfileSpace
        End If
        region = NativeViewRegion(view, views, work)
        maximum = -1
        Set display = view.GetFirstDisplayDimension5
        Do While Not display Is Nothing
            If NativePlacementSide(display) = NATIVE_LABEL_SIDE_BELOW Then
                lane = 0
                If featureDimensionLanes.Exists(display.GetNameForSelection) Then
                    lane = CLng(featureDimensionLanes(display.GetNameForSelection))
                End If
                If lane > maximum Then
                    maximum = lane
                End If
            End If
            Set display = display.GetNext5
        Loop
        If maximum >= 0 Then
            outline = view.GetOutline
            required = GOST_OBJECT_OFFSET_M + maximum * COORDINATE_LANE_PITCH_M + TEXT_HEIGHT_STANDARD_M + LEADER_LABEL_CLEARANCE_M
            available = outline(OUTLINE_BOTTOM_INDEX) - region(OUTLINE_BOTTOM_INDEX)
            If required > available Then
                shift = 2 * (required - available) + LEADER_LABEL_CLEARANCE_M
                headroom = sheetHeight - LEADER_STAMP_TOP_MARGIN_M - LEADER_BASE_SAFETY_MARGIN_M - GOST_OBJECT_OFFSET_M - outline(OUTLINE_TOP_INDEX)
                shift = MinD(shift, headroom)
                If shift > 0 Then
                    position = view.Position
                    position(SCREEN_Y_AXIS_INDEX) = position(SCREEN_Y_AXIS_INDEX) + shift
                    view.Position = position
                    d.ForceRebuild3 False
                    Log "PROFILE SPACE " & view.Name & "; lanes=" & maximum + 1 & "; shift=" & shift
                End If
            End If
        End If
        ReserveProfileSideText d, view, region
NextProfileSpace:
    Next item
End Sub

Private Sub ReserveProfileSideText(ByVal d As SldWorks.ModelDoc2, ByVal view As SldWorks.View, ByVal region As Variant)
    Dim display As SldWorks.DisplayDimension, request As Object, text As Variant, outline As Variant, position As Variant
    Dim leftWidth As Double, rightWidth As Double, required As Double, available As Double, shift As Double
    Set display = view.GetFirstDisplayDimension5
    Do While Not display Is Nothing
        If ProfileOwnsDimension(display) Then
            Set request = profileDimensionRequests(display.GetNameForSelection)
            text = NativeTextBounds(display.GetDisplayData)
            If NativePlacementSide(display) = NATIVE_LABEL_SIDE_RIGHT Then
                rightWidth = -MinD(-rightWidth, MinD(-(text(2) - text(0)), -(text(3) - text(1))))
            ElseIf request("role") = PROFILE_ROLE_SLOT And CLng(request("axis")) = SCREEN_Y_AXIS_INDEX And request("features").Count > 1 Then
                leftWidth = -MinD(-leftWidth, MinD(-(text(2) - text(0)), -(text(3) - text(1))))
            End If
        End If
        Set display = display.GetNext5
    Loop
    If leftWidth = 0 Then
        Exit Sub
    End If
    outline = view.GetOutline
    required = region(OUTLINE_LEFT_INDEX) + leftWidth + GOST_OBJECT_OFFSET_M + LEADER_MIN_DIAGONAL_M + 2 * LEADER_LABEL_CLEARANCE_M
    available = region(OUTLINE_RIGHT_INDEX) - outline(OUTLINE_RIGHT_INDEX) - GOST_OBJECT_OFFSET_M - rightWidth - 2 * LEADER_LABEL_CLEARANCE_M
    shift = MinD(required - outline(OUTLINE_LEFT_INDEX), available)
    If shift > 0 Then
        position = view.Position
        position(SCREEN_X_AXIS_INDEX) = position(SCREEN_X_AXIS_INDEX) + shift
        view.Position = position
        d.ForceRebuild3 False
        Log "PROFILE SIDE SPACE " & view.Name & "; shift=" & shift
    End If
End Sub

Private Sub ProfileDimension(ByVal d As SldWorks.ModelDoc2, ByVal view As SldWorks.View, ByVal request As Object, ByVal lanes As Collection, ByVal center As Variant)
    Dim first As SldWorks.Vertex, last As SldWorks.Vertex, a As Variant, b As Variant, outline As Variant
    Dim display As SldWorks.DisplayDimension, key As String, side As Long, lane As Long, interval As Variant, row As Variant
    Dim occupied As Boolean, annotation As SldWorks.Annotation, data As SldWorks.TextFormat, audit As Object
    Dim aligned As Boolean, edge As SldWorks.Edge
    If request.Exists("aligned") Then
        aligned = CBool(request("aligned"))
    End If
    Set first = ProfileReference(view, request("first"))
    Set last = ProfileReference(view, request("last"))
    If first Is Nothing Or last Is Nothing Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "No visible model references for profile dimension " & request("key")
    End If
    a = PatternSheetPoint(view, first.GetPoint)
    b = PatternSheetPoint(view, last.GetPoint)
    outline = view.GetOutline
    d.ClearSelection2 True
    If Not SelectFeatureReference(d, view, first) Or Not SelectFeatureReference(d, view, last, True) Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot select profile dimension references."
    End If
    If aligned Then
        Set edge = PlanarVisibleLine(view, request("first"), request("last"))
        If Not edge Is Nothing Then
            d.ClearSelection2 True
            If Not SelectFeatureReference(d, view, edge) Then
                Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot select an aligned panel dimension edge."
            End If
        End If
        Set display = d.AddDimension2((a(0) + b(0)) / 2 + GOST_OBJECT_OFFSET_M, (a(1) + b(1)) / 2 + GOST_OBJECT_OFFSET_M, 0)
        side = NATIVE_LABEL_SIDE_ABOVE
    ElseIf CLng(request("axis")) = SCREEN_X_AXIS_INDEX Then
        Set display = d.AddHorizontalDimension2((a(0) + b(0)) / 2, outline(3) + GOST_OBJECT_OFFSET_M, 0)
        side = NATIVE_LABEL_SIDE_BELOW
        If CBool(request("overall")) Then
            side = NATIVE_LABEL_SIDE_ABOVE
        End If
        If request.Exists("panel") Then
            If request("kind") = "notch-location" Then
                side = NATIVE_LABEL_SIDE_ABOVE
            End If
        End If
    Else
        Set display = d.AddVerticalDimension2(outline(0) - GOST_OBJECT_OFFSET_M, (a(1) + b(1)) / 2, 0)
        side = NATIVE_LABEL_SIDE_LEFT
        If (a(0) + b(0)) / 2 > center(0) And Not CBool(request("overall")) Then
            side = NATIVE_LABEL_SIDE_RIGHT
        End If
    End If
    If display Is Nothing Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot create a native profile dimension."
    End If
    If Abs(display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue - CDbl(request("value"))) > PROFILE_COORDINATE_TOLERANCE_M Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Profile dimension value differs from its model references."
    End If
    ApplyGostDimensionStyle display
    If request("features").Count > 1 Then
        If request.Exists("panel") Then
            display.SetText SW_DIMENSION_TEXT_PREFIX, CStr(request("features").Count) & ChrW(CIRCULAR_MULTIPLICATION_CHARACTER)
        Else
            display.SetText SW_DIMENSION_TEXT_SUFFIX, " (" & request("features").Count & " місця)"
        End If
    End If
    key = display.GetNameForSelection
    profileDimensionRequests.Add key, request
    featureDimensionOwners(key) = "profile|" & request("key")
    featureDimensionPlacement(key) = side
    If side = NATIVE_LABEL_SIDE_BELOW Or (side = NATIVE_LABEL_SIDE_ABOVE And request.Exists("panel") And Not aligned) Then
        interval = Array(MinD(a(0), b(0)) - GOST_ARROW_LENGTH_M - LEADER_LABEL_CLEARANCE_M, -MinD(-a(0), -b(0)) + GOST_ARROW_LENGTH_M + LEADER_LABEL_CLEARANCE_M)
        For lane = 0 To LEADER_BASE_ROW_LIMIT
            occupied = False
            For Each row In lanes
                If CLng(row(0)) = lane And interval(0) < row(2) And interval(1) > row(1) Then
                    If UBound(row) < PANEL_LANE_SIDE_INDEX Then
                        occupied = True
                    ElseIf row(PANEL_LANE_SIDE_INDEX) = side Then
                        occupied = True
                    End If
                End If
            Next row
            If Not occupied Then
                Exit For
            End If
        Next lane
        lanes.Add Array(lane, interval(0), interval(1), side)
        featureDimensionLanes(key) = lane
    End If
    Set audit = fs.OpenTextFile(fs.BuildPath(runDir, PROFILE_AUDIT_FILENAME), FILE_FOR_APPENDING, True, FILE_UNICODE)
    audit.WriteLine "DIM" & vbTab & jobNumber & vbTab & view.Name & vbTab & key & vbTab & request("role") & vbTab & request("axis") & vbTab & request("low") & vbTab & request("high") & vbTab & request("value") & vbTab & side
    audit.Close
End Sub

' Geometry, not filenames, determines whether this drawing can use profile roles.
Private Function MakeProfileRoleDrawing(ByVal d As SldWorks.ModelDoc2, ByVal doc As SldWorks.ModelDoc2, ByVal cfg As String, ByVal body As SldWorks.Body2) As Boolean
    Dim face As SldWorks.Face2, boundaries As Variant, boundary As SldWorks.Loop2, item As Variant
    Dim outer As Collection, points As Collection, openings As New Collection, projected As New Collection
    Dim first As SldWorks.View, second As SldWorks.View, plan As Collection, request As Variant, view As SldWorks.View
    Dim bounds As Variant, center As Variant, viewScale As Double, angle As Double, lanesOuter As New Collection, lanesSlots As New Collection
    Dim normal As Variant, box As Variant, axis As Long, thickness As Double, role As String, audit As Object
    Dim slotCount As Long, centerY As Double, lowerCenterY As Double, widthRatio As Double, heightRatio As Double, extension As Object
    Dim rotateDetail As Boolean, detailScale As Double, overviewCenterX As Double
    Set face = ProfilePlane(body)
    If face Is Nothing Then
        Exit Function
    End If
    boundaries = face.GetLoops
    For Each item In boundaries
        Set boundary = item
        Set points = ProfileLoopPoints(boundary)
        If points Is Nothing Then
            Exit Function
        End If
        If boundary.IsOuter Then
            Set outer = points
        Else
            openings.Add points
        End If
    Next item
    If outer Is Nothing Then
        Exit Function
    End If
    Set first = MakeView(d, doc, cfg, body, BestOrientation(doc, body), PROFILE_VIEW_CENTER_X_RATIO * sheetWidth, PROFILE_UPPER_CENTER_Y_RATIO * sheetHeight)
    If first Is Nothing Then
        Err.Raise vbObjectError + ERR_MAIN_VIEW, , "Cannot create the profile overview."
    End If
    first.UseSheetScale = False
    first.ScaleDecimal = 1
    d.ForceRebuild3 False
    bounds = first.GetOutline
    If bounds(3) - bounds(1) > bounds(2) - bounds(0) Then
        first.Angle = PI / 2
    End If
    d.ForceRebuild3 False
    Set points = ProfileProjection(outer, first)
    For Each item In openings
        projected.Add ProfileProjection(item, first)
    Next item
    Set plan = PlanProfileDimensions(points, projected)
    If plan Is Nothing Then
        d.ClearSelection2 True
        Set extension = d.Extension
        If Not extension.SelectByID2(CStr(first.Name), "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0) Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot remove an unsupported trial profile view."
        End If
        d.EditDelete
        Exit Function
    End If
    For Each request In plan
        If request("role") = PROFILE_ROLE_SLOT Then
            slotCount = slotCount + 1
        End If
    Next request
    If ProfilePageMode() And AsDrawing(d).GetSheetCount = 1 Then
        MakePagedProfileDrawing d, doc, cfg, body, first, plan, outer, openings
        MakeProfileRoleDrawing = True
        Exit Function
    End If
    centerY = sheetHeight * PROFILE_UPPER_CENTER_Y_RATIO
    overviewCenterX = sheetWidth * PROFILE_VIEW_CENTER_X_RATIO
    lowerCenterY = sheetHeight * PROFILE_LOWER_CENTER_Y_RATIO
    widthRatio = PROFILE_VIEW_WIDTH_RATIO
    heightRatio = PROFILE_VIEW_HEIGHT_RATIO
    If plan(1).Exists("panel") Then
        centerY = sheetHeight * PANEL_UPPER_CENTER_Y_RATIO
        lowerCenterY = sheetHeight * PANEL_LOWER_CENTER_Y_RATIO
        widthRatio = PANEL_VIEW_WIDTH_RATIO
        bounds = first.GetOutline
        rotateDetail = (bounds(2) - bounds(0)) / (bounds(3) - bounds(1)) > PANEL_ROTATION_ASPECT_RATIO And sheetHeight > sheetWidth
        If rotateDetail Then
            lowerCenterY = sheetHeight * PANEL_DETAIL_CENTER_Y_RATIO
            centerY = sheetHeight * PANEL_COMPACT_OVERVIEW_CENTER_Y_RATIO
            overviewCenterX = sheetWidth * PANEL_COMPACT_OVERVIEW_CENTER_X_RATIO
            widthRatio = PANEL_COMPACT_OVERVIEW_WIDTH_RATIO
        End If
    End If
    If slotCount = 0 Then
        centerY = sheetHeight * PROFILE_SINGLE_CENTER_Y_RATIO
        widthRatio = PROFILE_SINGLE_WIDTH_RATIO
        heightRatio = PROFILE_SINGLE_HEIGHT_RATIO
    End If
    viewScale = FitScale(first, sheetWidth * widthRatio, sheetHeight * heightRatio)
    SetPaperScale AsDrawing(d).GetCurrentSheet, viewScale
    first.UseSheetScale = True
    CenterView first, overviewCenterX, centerY
    If slotCount > 0 Then
        Set second = AsDrawing(d).CreateDrawViewFromModelView3(doc.GetPathName, first.GetOrientationName, sheetWidth * PROFILE_VIEW_CENTER_X_RATIO, lowerCenterY, 0)
        If second Is Nothing Then
            Err.Raise vbObjectError + ERR_MAIN_VIEW, , "Cannot create the profile feature view."
        End If
        second.ReferencedConfiguration = cfg
        ApplyBody second, body
        second.Angle = first.Angle
        second.UseSheetScale = True
        If rotateDetail Then
            second.Angle = first.Angle + PI / 2
            second.UseSheetScale = False
            second.ScaleDecimal = 1
            d.ForceRebuild3 False
            detailScale = FitScale(second, sheetWidth * PANEL_DETAIL_WIDTH_RATIO, sheetHeight * PANEL_DETAIL_HEIGHT_RATIO)
            second.ScaleDecimal = detailScale
        End If
    End If
    d.ForceRebuild3 False
    CenterView first, overviewCenterX, centerY
    If Not second Is Nothing Then
        CenterView second, sheetWidth * PROFILE_VIEW_CENTER_X_RATIO, lowerCenterY
    End If
    Set featureDimensionOwners = CreateObject("Scripting.Dictionary")
    Set featureDimensionPlacement = CreateObject("Scripting.Dictionary")
    Set featureDimensionLanes = CreateObject("Scripting.Dictionary")
    Set profileDimensionRequests = CreateObject("Scripting.Dictionary")
    Set profileGeometryViews = CreateObject("Scripting.Dictionary")
    Set profileSheetContours = CreateObject("Scripting.Dictionary")
    Set profileRawLoops = New Collection
    profileRawLoops.Add outer
    For Each item In openings
        profileRawLoops.Add item
    Next item
    profileGeometryViews.Add first.Name, True
    If Not second Is Nothing Then
        profileGeometryViews.Add second.Name, True
    End If
    Set featureDimensionRepeatedKeys = Nothing
    Set featureDimensionViewPriorities = Nothing
    Set featurePatternGraphics = Nothing
    For Each request In plan
        role = CStr(request("role"))
        Set view = first
        If role = PROFILE_ROLE_SLOT Then
            Set view = second
            If rotateDetail And CLng(request("axis")) <> PANEL_ALIGNED_AXIS Then
                request("axis") = 1 - CLng(request("axis"))
            End If
        End If
        bounds = view.GetOutline
        center = Array((bounds(0) + bounds(2)) / 2, (bounds(1) + bounds(3)) / 2)
        If role = PROFILE_ROLE_SLOT Then
            ProfileDimension d, view, request, lanesSlots, center
        Else
            ProfileDimension d, view, request, lanesOuter, center
        End If
    Next request
    profileRoleDrawing = True
    normal = face.Normal
    box = body.GetBodyBox
    For axis = SCREEN_X_AXIS_INDEX To SPACE_LAST_AXIS_INDEX
        If Abs(normal(axis)) > FACE_NORMAL_ALIGNMENT_COSINE Then
            thickness = box(axis + BBOX_MAX_OFFSET) - box(axis)
        End If
    Next axis
    AddText d, "Товщина " & CStr(Round(thickness * MILLIMETRES_PER_METRE, LINEAR_DECIMAL_PLACES)) & " мм", CONTENT_LEFT_M, PROFILE_THICKNESS_NOTE_Y_M, TEXT_HEIGHT_STANDARD_M, 0
    If rotateDetail Then
        AddText d, "Дрібні вирізи, повернуто на 90° (1:" & CStr(Round(1 / detailScale, 0)) & ")", CONTENT_LEFT_M, PANEL_DETAIL_NOTE_Y_M, TEXT_HEIGHT_STANDARD_M, 0
    End If
    Log "PROFILE ROLES dimensions=" & plan.Count & "; scale=" & viewScale & "; coordinate precision=" & PROFILE_COORDINATE_TOLERANCE_M
    MakeProfileRoleDrawing = True
End Function
