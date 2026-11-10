Attribute VB_Name = "compact_profiles"
Option Explicit

' Reserve complete feature families before placing individual dimensions.
' A collision may move text within a family, but never create another sheet.
Private Function CompactGroup(ByVal kind As String, ByVal page As Long) As Object
    Dim result As Object, requests As New Collection
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "type", kind
    result.Add "page", page
    result.Add "requests", requests
    result.Add "rotation", 0#
    Set CompactGroup = result
End Function

Private Function CompactFrameBounds() As Variant
    Dim node As Variant, point As Variant, result As Variant
    result = Array(SEARCH_INITIAL_LIMIT, SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT)
    For Each node In profilePanelFrame
        point = ProfilePagePoint(node("point"))
        result = ProfilePageUnion(result, Array(point(0), point(1), point(0), point(1)))
    Next node
    CompactFrameBounds = result
End Function

Private Function CompactMajorCuts() As Collection
    Dim result As New Collection, index As Long, a As Object, b As Object, c As Object, last As Object
    Dim area As Double, span As Double, frameBounds As Variant, group As Object
    frameBounds = CompactFrameBounds()
    span = -MinD(-(frameBounds(2) - frameBounds(0)), -(frameBounds(3) - frameBounds(1)))
    For index = 1 To profilePanelFrame.Count
        Set a = PanelNode(profilePanelFrame, index)
        Set b = PanelNode(profilePanelFrame, index + 1)
        area = area + a("x") * b("y") - b("x") * a("y")
    Next index
    For index = 1 To profilePanelFrame.Count
        Set a = PanelNode(profilePanelFrame, index)
        Set b = PanelNode(profilePanelFrame, index + 1)
        Set c = PanelNode(profilePanelFrame, index + 2)
        Set last = PanelNode(profilePanelFrame, index + 3)
        If ProfileNotch(a, b, c, last, area) And PanelDistance(a, b) > span * PANEL_MINOR_DEPTH_RATIO Then
            Set group = CompactGroup("cut", result.Count + 1)
            group.Add "core", Array(MinD(a("x"), c("x")), MinD(a("y"), c("y")), -MinD(-a("x"), -c("x")), -MinD(-a("y"), -c("y")))
            result.Add group
        End If
    Next index
    Set CompactMajorCuts = result
End Function

Private Function CompactFeatureFamily(ByVal feature As Object, ByVal cuts As Collection, ByVal frameBounds As Variant) As String
    Dim a As Object, b As Object, cut As Variant, box As Variant, x As Double, y As Double, index As Long
    Set a = feature("first")
    Set b = feature("last")
    x = (a("x") + b("x")) / 2
    y = (a("y") + b("y")) / 2
    For Each cut In cuts
        index = index + 1
        box = cut("core")
        If x >= box(0) - COMPACT_CONTEXT_M And x <= box(2) + COMPACT_CONTEXT_M And y >= box(1) - COMPACT_CONTEXT_M And y <= box(3) + COMPACT_CONTEXT_M Then
            CompactFeatureFamily = "cut" & index
            Exit Function
        End If
    Next cut
    If Abs(x - frameBounds(0)) < COMPACT_WALL_CONTEXT_M And ProfileSegmentAxis(a, b) = SCREEN_Y_AXIS_INDEX Then
        CompactFeatureFamily = "left-wall"
    ElseIf Abs(x - frameBounds(2)) < COMPACT_WALL_CONTEXT_M And ProfileSegmentAxis(a, b) = SCREEN_Y_AXIS_INDEX Then
        CompactFeatureFamily = "right-wall"
    ElseIf Abs(y - frameBounds(1)) < PROFILE_COORDINATE_TOLERANCE_M Then
        CompactFeatureFamily = "bottom"
    ElseIf x < (frameBounds(0) + frameBounds(2)) / 2 Then
        CompactFeatureFamily = "left-slope"
    Else
        CompactFeatureFamily = "right-slope"
    End If
End Function

Private Function CompactPlanGroups(ByVal cuts As Collection) As Collection
    Dim families As Object, result As New Collection, group As Variant, request As Variant, feature As Object
    Dim ids As Variant, id As String, kind As String, index As Long, bounds As Variant, a As Variant, b As Variant
    Set families = CreateObject("Scripting.Dictionary")
    For Each group In cuts
        index = index + 1
        families.Add "cut" & index, group
    Next group
    families.Add "depth", CompactGroup("depth", 1)
    families.Add "micro", CompactGroup("micro", 1)
    families.Add "bottom", CompactGroup("bottom", 1)
    families.Add "left-wall", CompactGroup("left-wall", 2)
    families.Add "right-wall", CompactGroup("right-wall", 2)
    families.Add "left-slope", CompactGroup("left-slope", 2)
    families.Add "right-slope", CompactGroup("right-slope", 2)
    bounds = CompactFrameBounds()
    For Each request In profilePagePlan
        If request("role") = PROFILE_ROLE_SLOT Then
            ids = request("features").Keys
            id = CStr(ids(0))
            Set feature = profilePanelFeatures(CLng(Mid$(id, Len("notch") + 1)))
            kind = CompactFeatureFamily(feature, cuts, bounds)
            If request("kind") = "notch-depth" And request("features").Count >= COMPACT_DEPTH_COUNT Then
                kind = "depth"
            ElseIf feature("width") < COMPACT_MICRO_LIMIT_M And feature("depth") < COMPACT_MICRO_LIMIT_M And request("kind") <> "notch-location" Then
                kind = "micro"
            End If
            Set group = families(kind)
            group("requests").Add request
            request("compactFamily") = kind
        End If
    Next request
    For Each group In families.Items
        If group("requests").Count > 0 Then
            If Not group.Exists("core") Then
                For Each request In group("requests")
                    a = ProfilePagePoint(request("first"))
                    b = ProfilePagePoint(request("last"))
                    If Not group.Exists("core") Then
                        group.Add "core", Array(a(0), a(1), a(0), a(1))
                    End If
                    group("core") = ProfilePageUnion(group("core"), Array(MinD(a(0), b(0)), MinD(a(1), b(1)), -MinD(-a(0), -b(0)), -MinD(-a(1), -b(1))))
                Next request
            End If
            result.Add group
        End If
    Next group
    Set CompactPlanGroups = result
End Function

Private Function CompactRotatePoint(ByVal point As Variant, ByVal angle As Double) As Variant
    CompactRotatePoint = Array(point(0) * Cos(angle) - point(1) * Sin(angle), point(0) * Sin(angle) + point(1) * Cos(angle))
End Function

Private Function CompactRotateBounds(ByVal bounds As Variant, ByVal angle As Double) As Variant
    Dim point As Variant, rotated As Variant, result As Variant
    result = Array(SEARCH_INITIAL_LIMIT, SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT)
    For Each point In Array(Array(bounds(0), bounds(1)), Array(bounds(2), bounds(1)), Array(bounds(2), bounds(3)), Array(bounds(0), bounds(3)))
        rotated = CompactRotatePoint(point, angle)
        result = ProfilePageUnion(result, Array(rotated(0), rotated(1), rotated(0), rotated(1)))
    Next point
    CompactRotateBounds = result
End Function

Private Sub CompactConfigureGroup(ByVal group As Object)
    Dim core As Variant, bounds As Variant, frameBounds As Variant, kind As String, viewScaleValue As Double
    Dim x As Double, y As Double, region As Variant, index As Long, middle As Double
    core = group("core")
    bounds = core
    kind = group("type")
    frameBounds = CompactFrameBounds()
    viewScaleValue = SCALE_1_TO_2
    For index = 0 To 1
        bounds(index) = bounds(index) - COMPACT_CONTEXT_M
        bounds(index + 2) = bounds(index + 2) + COMPACT_CONTEXT_M
    Next index
    If kind = "cut" Then
        viewScaleValue = StandardScale(MinD(COMPACT_CUT_PAPER_WIDTH_M / (core(2) - core(0)), COMPACT_CUT_PAPER_HEIGHT_M / (core(3) - core(1))))
        bounds(3) = core(3) + COMPACT_CUT_HEIGHT_CONTEXT_M
        If group("page") = 1 Then
            x = COMPACT_CUT_FIRST_X_M
            y = COMPACT_CUT_FIRST_Y_M + (core(3) - core(1)) * viewScaleValue / 2
            region = Array(LEADER_WORK_LEFT_M, COMPACT_CUT_FIRST_BOTTOM_M, COMPACT_CUT_FIRST_RIGHT_M, COMPACT_CUT_FIRST_TOP_M)
        Else
            x = COMPACT_CUT_SECOND_X_M
            y = COMPACT_CUT_SECOND_Y_M + (core(3) - core(1)) * viewScaleValue / 2
            region = Array(LEADER_WORK_LEFT_M, COMPACT_CUT_SECOND_BOTTOM_M, sheetWidth - LEADER_WORK_EDGE_MARGIN_M, COMPACT_CUT_SECOND_TOP_M)
        End If
    ElseIf kind = "bottom" Then
        viewScaleValue = StandardScale((sheetWidth - LEADER_WORK_LEFT_M - LEADER_WORK_EDGE_MARGIN_M - COMPACT_CONTEXT_M * 2) / (frameBounds(2) - frameBounds(0)))
        bounds = Array(frameBounds(0) - COMPACT_BOTTOM_CONTEXT_M, frameBounds(1) - COMPACT_BOTTOM_CONTEXT_M, frameBounds(2) + COMPACT_BOTTOM_CONTEXT_M, frameBounds(1) + COMPACT_BOTTOM_CONTEXT_M)
        core = Array(frameBounds(0), frameBounds(1), frameBounds(2), frameBounds(1))
        x = (LEADER_WORK_LEFT_M + sheetWidth - LEADER_WORK_EDGE_MARGIN_M) / 2
        y = COMPACT_BOTTOM_Y_M
        region = Array(LEADER_WORK_LEFT_M, COMPACT_BOTTOM_BOTTOM_M, sheetWidth - LEADER_WORK_EDGE_MARGIN_M, COMPACT_BOTTOM_TOP_M)
    ElseIf kind = "depth" Or kind = "micro" Then
        x = COMPACT_MINI_X_M
        y = COMPACT_DEPTH_Y_M
        viewScaleValue = 1
        region = Array(COMPACT_MINI_LEFT_M, COMPACT_DEPTH_BOTTOM_M, sheetWidth - LEADER_WORK_EDGE_MARGIN_M, COMPACT_DEPTH_TOP_M)
        If kind = "micro" Then
            y = COMPACT_MICRO_Y_M
            viewScaleValue = SCALE_10_TO_1
            region = Array(COMPACT_MINI_LEFT_M, COMPACT_MICRO_BOTTOM_M, sheetWidth - LEADER_WORK_EDGE_MARGIN_M, COMPACT_MICRO_TOP_M)
        End If
        For index = 0 To 1
            middle = (core(index) + core(index + 2)) / 2
            bounds(index) = middle - -MinD(-(core(index + 2) - core(index) + COMPACT_MINI_CONTEXT_PAPER_M / viewScaleValue), -COMPACT_MINI_CROP_PAPER_M / viewScaleValue) / 2
            bounds(index + 2) = middle * 2 - bounds(index)
        Next index
    ElseIf kind = "left-wall" Or kind = "right-wall" Then
        group("rotation") = -PI / 2
        x = COMPACT_LEFT_X_M
        region = Array(LEADER_WORK_LEFT_M, COMPACT_WALL_BOTTOM_M, COMPACT_LEFT_REGION_RIGHT_M, COMPACT_WALL_TOP_M)
        If kind = "right-wall" Then
            group("rotation") = PI / 2
            x = COMPACT_RIGHT_X_M
            region = Array(COMPACT_RIGHT_REGION_LEFT_M, COMPACT_WALL_BOTTOM_M, sheetWidth - LEADER_WORK_EDGE_MARGIN_M, COMPACT_WALL_TOP_M)
        End If
        y = COMPACT_WALL_Y_M
        viewScaleValue = StandardScale(COMPACT_WALL_PAPER_LENGTH_M / (core(3) - core(1)))
        bounds = Array(core(0) - COMPACT_WALL_CONTEXT_M, core(1) - COMPACT_WALL_CONTEXT_M, core(2) + COMPACT_WALL_CONTEXT_M, core(3) + COMPACT_WALL_CONTEXT_M)
    Else
        x = COMPACT_LEFT_X_M
        region = Array(LEADER_WORK_LEFT_M, COMPACT_SLOPE_BOTTOM_M, COMPACT_LEFT_REGION_RIGHT_M, COMPACT_SLOPE_TOP_M)
        If kind = "right-slope" Then
            x = COMPACT_RIGHT_X_M
            region = Array(COMPACT_RIGHT_REGION_LEFT_M, COMPACT_SLOPE_BOTTOM_M, sheetWidth - LEADER_WORK_EDGE_MARGIN_M, COMPACT_SLOPE_TOP_M)
        End If
        y = COMPACT_SLOPE_Y_M
        viewScaleValue = StandardScale(MinD(SCALE_1_TO_2, COMPACT_SLOPE_PAPER_WIDTH_M / (core(2) - core(0))))
        bounds = Array(core(0) - COMPACT_WALL_CONTEXT_M, core(1) - COMPACT_WALL_CONTEXT_M, core(2) + COMPACT_WALL_CONTEXT_M, core(3) + COMPACT_WALL_CONTEXT_M)
    End If
    group("core") = core
    group.Add "bounds", bounds
    group.Add "region", region
    group.Add "center", Array(x, y)
    group.Add "scale", viewScaleValue
End Sub

Private Sub CompactCreateView(ByVal d As SldWorks.ModelDoc2, ByVal doc As SldWorks.ModelDoc2, ByVal body As SldWorks.Body2, ByVal group As Object)
    Dim view As SldWorks.View, request As Variant, first As Object, bounds As Variant, core As Variant
    Dim reference As Variant, anchor As Variant, position As Variant, center As Variant, angle As Double, lanes As New Collection
    If Not group.Exists("configured") Then
        CompactConfigureGroup group
    End If
    bounds = group("bounds")
    core = group("core")
    center = group("center")
    angle = group("rotation")
    Set first = group("requests")(1)
    Set view = MakeView(d, doc, profilePageConfiguration, body, profilePageOrientation, center(0), center(1))
    view.Angle = profilePageAngle + angle
    view.UseSheetScale = False
    view.ScaleDecimal = group("scale")
    d.ForceRebuild3 False
    core = CompactRotateBounds(core, angle)
    bounds = CompactRotateBounds(bounds, angle)
    ProfilePageCrop d, view, first, bounds, angle
    d.ForceRebuild3 False
    reference = CompactRotatePoint(ProfilePagePoint(first("first")), angle)
    anchor = PatternSheetPoint(view, first("first"))
    position = view.Position
    position(0) = position(0) + center(0) - anchor(0) - ((core(0) + core(2)) / 2 - reference(0)) * group("scale")
    position(1) = position(1) + center(1) - anchor(1) - ((core(1) + core(3)) / 2 - reference(1)) * group("scale")
    view.Position = position
    d.ForceRebuild3 False
    view.UpdateViewDisplayGeometry
    profilePageCrops.Add view.Name, ProfilePageMapBounds(bounds, reference, PatternSheetPoint(view, first("first")), group("scale"))
    profilePageRegions.Add view.Name, group("region")
    profileGeometryViews.Add view.Name, True
    profilePageViews.Add Array(AsDrawing(d).GetCurrentSheet.GetName, view.Name)
    group.Add "sheet", AsDrawing(d).GetCurrentSheet.GetName
    group.Add "view", view.Name
    group.Add "cell", group("region")
    group.Add "number", profilePageViews.Count - 1
    group.Add "paperCore", ProfilePageMapBounds(core, reference, PatternSheetPoint(view, first("first")), group("scale"))
    AsDrawing(d).ActivateView view.Name
    view.FocusLocked = True
    For Each request In group("requests")
        If Abs(Sin(angle)) > FACE_NORMAL_ALIGNMENT_COSINE Then
            If request("axis") <> PANEL_ALIGNED_AXIS Then
                request("axis") = 1 - request("axis")
            End If
        End If
        ProfileDimension d, view, request, lanes, center
    Next request
    view.FocusLocked = False
    AsDrawing(d).ActivateView ""
    Log "COMPACT family=" & group("type") & "; sheet=" & group("page") & "; requests=" & group("requests").Count & "; scale=" & group("scale")
End Sub

Private Sub BuildCompactProfilePages(ByVal d As SldWorks.ModelDoc2)
    Dim cuts As Collection, doc As SldWorks.ModelDoc2, body As SldWorks.Body2, overview As SldWorks.View
    Dim request As Variant, group As Variant, lanes As New Collection, outline As Variant, center As Variant, page As Long
    Dim supported As Boolean, frameBounds As Variant
    Set cuts = CompactMajorCuts()
    supported = cuts.Count = COMPACT_TARGET_SHEETS And Abs(sheetWidth - A4_SHORT_SIDE_M) < PROFILE_COORDINATE_TOLERANCE_M And Abs(sheetHeight - A4_LONG_SIDE_AND_A3_SHORT_SIDE_M) < PROFILE_COORDINATE_TOLERANCE_M
    frameBounds = CompactFrameBounds()
    For Each group In cuts
        outline = group("core")
        If Abs(outline(1) - frameBounds(1)) > PROFILE_COORDINATE_TOLERANCE_M Then
            supported = False
        End If
    Next group
    If Not supported Then
        profileCompactMode = False
        profileFamilyMode = True
        BuildFamilyProfilePages d
        Exit Sub
    End If
    Set doc = sw.GetOpenDocumentByName(profilePageModel)
    Set body = FindBody(doc, profilePageBody)
    Set featureDimensionOwners = CreateObject("Scripting.Dictionary")
    Set featureDimensionPlacement = CreateObject("Scripting.Dictionary")
    Set featureDimensionLanes = CreateObject("Scripting.Dictionary")
    Set profileDimensionRequests = CreateObject("Scripting.Dictionary")
    Set profileGeometryViews = CreateObject("Scripting.Dictionary")
    Set profileSheetContours = CreateObject("Scripting.Dictionary")
    Set profilePageRegions = CreateObject("Scripting.Dictionary")
    Set profilePageCrops = CreateObject("Scripting.Dictionary")
    Set profilePageViews = New Collection
    Set profilePageSheets = New Collection
    Set profilePageGroups = CompactPlanGroups(cuts)
    For page = 1 To COMPACT_TARGET_SHEETS
        If page = 1 Then
            AsDrawing(d).ActivateSheet profilePageFirstSheet
        Else
            NewFormattedSheet d, "Дрібні вирізи"
            profilePageSheets.Add AsDrawing(d).GetCurrentSheet.GetName
        End If
        Set overview = ProfilePageOverview(d, doc, body)
        If page = 1 Then
            profileCompactEdgeCount = VisibleEdgeCount(overview)
            outline = overview.GetOutline
            center = Array((outline(0) + outline(2)) / 2, (outline(1) + outline(3)) / 2)
            For Each request In profilePagePlan
                If request("role") <> PROFILE_ROLE_SLOT Then
                    ProfileDimension d, overview, request, lanes, center
                End If
            Next request
        Else
            CenterView overview, (LEADER_WORK_LEFT_M + sheetWidth - LEADER_WORK_EDGE_MARGIN_M) / 2, COMPACT_SECOND_OVERVIEW_Y_M
            profilePageRegions(overview.Name) = Array(LEADER_WORK_LEFT_M, COMPACT_SECOND_OVERVIEW_BOTTOM_M, sheetWidth - LEADER_WORK_EDGE_MARGIN_M, sheetHeight - PROFILE_PAGE_MAIN_TOP_MARGIN_M)
        End If
        For Each group In profilePageGroups
            If group("page") = page Then
                CompactCreateView d, doc, body, group
            End If
        Next group
    Next page
    AsDrawing(d).ActivateSheet profilePageFirstSheet
End Sub

' Place the baseline independently from the label. Short intervals retain
' their native arrows while their text can move along that same baseline.
Private Sub CompactPlaceBaseline(ByVal d As SldWorks.ModelDoc2, ByVal display As SldWorks.DisplayDimension, ByVal view As SldWorks.View, ByVal group As Object, ByVal lane As Long)
    Dim request As Object, a As Variant, b As Variant, core As Variant, region As Variant, annotation As SldWorks.Annotation
    Dim dx As Double, dy As Double, length As Double, nx As Double, ny As Double, distance As Double, target As Variant
    Dim position As Variant, arrow As Variant, index As Long, shift As Double, kind As String, textFormat As SldWorks.TextFormat
    Set request = profileDimensionRequests(display.GetNameForSelection)
    a = PatternSheetPoint(view, request("first"))
    b = PatternSheetPoint(view, request("last"))
    dx = b(0) - a(0)
    dy = b(1) - a(1)
    If request("axis") = SCREEN_X_AXIS_INDEX Then
        dy = 0
    ElseIf request("axis") = SCREEN_Y_AXIS_INDEX Then
        dx = 0
    End If
    length = Sqr(dx * dx + dy * dy)
    dx = dx / length
    dy = dy / length
    If dx < -DIRECTION_ZERO_TOLERANCE Or (Abs(dx) < DIRECTION_ZERO_TOLERANCE And dy < 0) Then
        dx = -dx
        dy = -dy
    End If
    nx = -dy
    ny = dx
    core = group("paperCore")
    region = group("region")
    kind = group("type")
    distance = GOST_OBJECT_OFFSET_M
    If request("kind") = "notch-location" Then
        distance = distance + COORDINATE_LANE_PITCH_M
    End If
    If kind = "bottom" Then
        nx = 0
        ny = -1
        If request("kind") = "notch-location" And request("value") > COMPACT_BOTTOM_LONG_GAP_M Then
            distance = distance + COORDINATE_LANE_PITCH_M
        End If
    ElseIf kind = "cut" Then
        If Abs(dx) > Abs(dy) Then
            nx = 0
            ny = -1
        Else
            nx = -1
            ny = 0
            If (a(0) + b(0)) / 2 > (core(0) + core(2)) / 2 Then
                nx = 1
                distance = distance + TEXT_HEIGHT_SMALL_M
            End If
        End If
    ElseIf kind = "left-wall" Or kind = "right-wall" Then
        If Abs(dx) > Abs(dy) Then
            nx = 0
            ny = 1
            distance = distance + (lane Mod COMPACT_ROW_COUNT) * COORDINATE_LANE_PITCH_M
            If request("kind") = "notch-location" Then
                distance = distance + COORDINATE_LANE_PITCH_M
            End If
        Else
            nx = 1
            ny = 0
        End If
    ElseIf kind = "depth" Then
        nx = -1
        ny = 0
    ElseIf kind = "left-slope" Or kind = "right-slope" Then
        If ny < 0 Then
            nx = -nx
            ny = -ny
        End If
    End If
    target = Array(a(0) + nx * distance, a(1) + ny * distance)
    If kind = "cut" Then
        If ny < 0 Then
            target(1) = core(3) - distance
        ElseIf nx < 0 Then
            target(0) = core(0) - distance
        Else
            target(0) = core(2) + distance
        End If
    ElseIf kind = "left-wall" Or kind = "right-wall" Then
        If ny > 0 Then
            target(1) = core(3) + distance
        Else
            target(0) = core(2) + distance
        End If
    End If
    request("view") = view.Name
    request("localBase") = False
    request("externalWitness") = True
    request("sheetFirst") = a
    request("sheetLast") = b
    request("compactDirection") = Array(dx, dy)
    request("compactNormal") = Array(nx, ny)
    Set annotation = display.GetAnnotation
    Set textFormat = annotation.GetTextFormat(SW_ANNOTATION_TEXT_FORMAT_FIRST)
    textFormat.CharHeight = TEXT_HEIGHT_SMALL_M
    annotation.SetTextFormat SW_ANNOTATION_TEXT_FORMAT_FIRST, False, textFormat
    display.OffsetText = False
    display.CenterText = True
    If length >= 2 * GOST_ARROW_HEAD_LENGTH_M Then
        display.ArrowSide = SW_DIMENSION_ARROWS_INSIDE
    Else
        display.ArrowSide = SW_DIMENSION_ARROWS_OUTSIDE
    End If
    annotation.SetPosition2 (a(0) + b(0)) / 2 + nx * distance, (a(1) + b(1)) / 2 + ny * distance, 0
    display.CenterText = True
    For index = 1 To LEADER_CLAMP_ATTEMPTS
        d.GraphicsRedraw2
        arrow = display.GetDisplayData.GetArrowHeadAtIndex2(0)
        shift = (target(0) - arrow(0)) * nx + (target(1) - arrow(1)) * ny
        If Abs(shift) < PROFILE_MATERIAL_TOLERANCE_M Then
            Exit For
        End If
        position = annotation.GetPosition
        annotation.SetPosition2 position(0) + nx * shift, position(1) + ny * shift, position(2)
        display.CenterText = True
    Next index
End Sub

Private Sub FinishCompactProfilePages(ByVal d As SldWorks.ModelDoc2)
    Dim sheetName As Variant, sheetViews As Collection, view As SldWorks.View, item As Variant, group As Variant
    Dim display As SldWorks.DisplayDimension, request As Object, records As Collection, index As Long, textFormat As SldWorks.TextFormat
    Dim other As Variant, box As Variant, counters As Object, kind As String, caption As Variant
    For Each sheetName In AsDrawing(d).GetSheetNames
        AsDrawing(d).ActivateSheet CStr(sheetName)
        AsDrawing(d).ActivateView ""
        Set sheetViews = New Collection
        Set view = AsDrawing(d).GetFirstView.GetNextView
        Do While Not view Is Nothing
            sheetViews.Add view
            Set view = view.GetNextView
        Loop
        CacheProfileContours sheetViews
        For Each item In sheetViews
            Set view = item
            If Not profilePageRegions.Exists(view.Name) Then
                GoTo NextCompactView
            End If
            If Not profilePageCrops.Exists(view.Name) Then
                If VisibleEdgeCount(view) <> profileCompactEdgeCount Then
                    Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "A complete panel overview lost model geometry: " & view.Name
                End If
            End If
            Set group = Nothing
            For Each other In profilePageGroups
                If other("view") = view.Name Then
                    Set group = other
                    Exit For
                End If
            Next other
            Set records = New Collection
            Set counters = CreateObject("Scripting.Dictionary")
            Set display = view.GetFirstDisplayDimension5
            Do While Not display Is Nothing
                If group Is Nothing Then
                    Set textFormat = display.GetAnnotation.GetTextFormat(SW_ANNOTATION_TEXT_FORMAT_FIRST)
                    textFormat.CharHeight = TEXT_HEIGHT_SMALL_M
                    display.GetAnnotation.SetTextFormat SW_ANNOTATION_TEXT_FORMAT_FIRST, False, textFormat
                    ProfilePagePlaceOverview d, display, view
                    Set request = profileDimensionRequests(display.GetNameForSelection)
                    If request("axis") = SCREEN_X_AXIS_INDEX Then
                        request("compactDirection") = Array(1#, 0#)
                    Else
                        request("compactDirection") = Array(0#, 1#)
                    End If
                Else
                    Set request = profileDimensionRequests(display.GetNameForSelection)
                    kind = request("kind")
                    If Not counters.Exists(kind) Then
                        counters.Add kind, 0
                    End If
                    index = counters(kind)
                    CompactPlaceBaseline d, display, view, group, index
                    counters(kind) = index + 1
                End If
                records.Add display
                Set display = display.GetNext5
            Loop
            If group Is Nothing And records.Count > 0 Then
                ProfilePageFitOverview d, view, records, profilePageRegions(view.Name)
            End If
            CompactPackLabels d, view, records, profilePageRegions(view.Name)
            Log "COMPACT placed " & view.Name & "; dimensions=" & records.Count
NextCompactView:
        Next item
    Next sheetName
    For Each group In profilePageGroups
        AsDrawing(d).ActivateSheet CStr(group("sheet"))
        caption = CompactCaption(group)
        AddText d, caption(0) & " (" & ScaleText(group("scale")) & ")", caption(1), caption(2), TEXT_HEIGHT_SMALL_M, 0
    Next group
    AuditProfilePages d
    AsDrawing(d).ActivateSheet profilePageFirstSheet
    Log "COMPACT COMPLETE sheets=" & AsDrawing(d).GetSheetCount & "; requests=" & profilePagePlan.Count
End Sub

Private Function CompactCaption(ByVal group As Object) As Variant
    Dim region As Variant, label As String, y As Double
    region = group("region")
    Select Case group("type")
        Case "cut"
            label = "Виріз А"
            If group("page") = COMPACT_TARGET_SHEETS Then
                label = "Виріз Б"
            End If
            y = region(3) - PROFILE_PAGE_CAPTION_GAP_M
        Case "depth"
            label = "Глибина"
            y = COMPACT_DEPTH_CAPTION_Y_M
        Case "micro"
            label = "Кут"
            y = COMPACT_MICRO_CAPTION_Y_M
        Case "bottom"
            label = "Нижня кромка"
            y = COMPACT_BOTTOM_CAPTION_Y_M
        Case "left-wall", "right-wall"
            label = "В"
            If group("type") = "right-wall" Then
                label = "Г"
            End If
            y = COMPACT_WALL_CAPTION_Y_M
        Case Else
            label = "Д"
            If group("type") = "right-slope" Then
                label = "Е"
            End If
            y = COMPACT_SLOPE_CAPTION_Y_M
    End Select
    CompactCaption = Array(label, region(0), y)
End Function

Private Function CompactLabelScore(ByVal display As SldWorks.DisplayDimension, ByVal records As Collection, ByVal region As Variant) As Long
    Dim box As Variant, other As Variant, otherBox As Variant
    Dim request As Object, loops As Variant, polygons As Collection, lines As Collection, otherPolygons As Collection
    box = NativeTextBounds(display.GetDisplayData, display)
    If Not NativeBoundsInside(NativeDimensionBounds(display), region) Then
        CompactLabelScore = CompactLabelScore + COMPACT_OUTSIDE_COST
    End If
    Set request = profileDimensionRequests(display.GetNameForSelection)
    loops = profileSheetContours(request("view"))
    If PanelBoxTouchesMaterial(Array(box(0) - COMPACT_TEXT_CLEARANCE_M, box(1) - COMPACT_TEXT_CLEARANCE_M, box(2) + COMPACT_TEXT_CLEARANCE_M, box(3) + COMPACT_TEXT_CLEARANCE_M), loops) Then
        CompactLabelScore = CompactLabelScore + COMPACT_OUTSIDE_COST
    End If
    Set polygons = CompactTextPolygons(display)
    Set lines = PlanarDisplayLines(display.GetDisplayData)
    If PlanarLinesHitText(lines, polygons) Then
        CompactLabelScore = CompactLabelScore + 1
    End If
    For Each other In records
        If other.GetNameForSelection <> display.GetNameForSelection Then
            otherBox = NativeTextBounds(other.GetDisplayData, other)
            If NativeBoxesOverlap(box, otherBox, COMPACT_TEXT_CLEARANCE_M) Then
                CompactLabelScore = CompactLabelScore + COMPACT_TEXT_COLLISION_COST
            End If
            If PlanarLinesHitText(PlanarDisplayLines(other.GetDisplayData), polygons) Then
                CompactLabelScore = CompactLabelScore + 1
            End If
            Set otherPolygons = CompactTextPolygons(other)
            If PlanarLinesHitText(lines, otherPolygons) Then
                CompactLabelScore = CompactLabelScore + 1
            End If
        End If
    Next other
End Function

' Native text positions lie on the dimension baseline. Add clearance along
' the string only: a normal expansion would reject that baseline itself.
Private Function CompactTextPolygons(ByVal display As SldWorks.DisplayDimension) As Collection
    Dim result As New Collection, polygon As Variant
    Dim angle As Double, data As SldWorks.DisplayData
    Set data = display.GetDisplayData
    angle = data.GetTextAngleAtIndex(0)
    For Each polygon In PlanarTextPolygons(display, data, 0)
        result.Add CompactExpandTextPolygon(polygon, angle)
    Next polygon
    Set CompactTextPolygons = result
End Function

Private Function CompactExpandTextPolygon(ByVal polygon As Variant, ByVal angle As Double) As Variant
    Dim corner As Long, side As Long, point As Variant, result(3) As Variant
    For corner = 0 To 3
        side = -1
        If corner = 1 Or corner = 2 Then
            side = 1
        End If
        point = polygon(corner)
        result(corner) = Array(point(0) + side * COMPACT_TEXT_CLEARANCE_M * Cos(angle), point(1) + side * COMPACT_TEXT_CLEARANCE_M * Sin(angle))
    Next corner
    CompactExpandTextPolygon = result
End Function

Private Sub CompactPackLabels(ByVal d As SldWorks.ModelDoc2, ByVal view As SldWorks.View, ByVal records As Collection, ByVal region As Variant)
    Dim item As Variant, display As SldWorks.DisplayDimension, request As Object, annotation As SldWorks.Annotation
    Dim direction As Variant, bestPosition As Variant, initialPosition As Variant
    Dim pass As Long, trial As Long, sign As Long, score As Long, best As Long, shift As Double, bestCentered As Boolean
    For pass = 1 To COMPACT_LAYOUT_PASSES
        For Each item In records
            Set display = item
            best = CompactLabelScore(display, records, region)
            If best = 0 Then
                GoTo NextCompactLabel
            End If
            Set request = profileDimensionRequests(display.GetNameForSelection)
            Set annotation = display.GetAnnotation
            direction = request("compactDirection")
            initialPosition = annotation.GetPosition
            bestPosition = initialPosition
            bestCentered = display.CenterText
            For trial = 1 To COMPACT_LABEL_TRIALS
                For sign = -1 To 1 Step 2
                    shift = trial * COMPACT_LABEL_STEP_M * sign
                    display.CenterText = False
                    annotation.SetPosition2 initialPosition(0) + direction(0) * shift, initialPosition(1) + direction(1) * shift, initialPosition(2)
                    display.CenterText = False
                    d.GraphicsRedraw2
                    score = CompactLabelScore(display, records, region)
                    If score < best Then
                        best = score
                        bestPosition = annotation.GetPosition
                        bestCentered = False
                    End If
                    If best = 0 Then
                        Exit For
                    End If
                Next sign
                If best = 0 Then
                    Exit For
                End If
            Next trial
            annotation.SetPosition2 bestPosition(0), bestPosition(1), bestPosition(2)
            display.CenterText = bestCentered
NextCompactLabel:
        Next item
    Next pass
    For Each item In records
        score = CompactLabelScore(item, records, region)
        If score <> 0 Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "The two-sheet panel layout has an unresolved text or view-boundary collision: " & item.GetNameForSelection & "; score=" & score
        End If
    Next item
End Sub
