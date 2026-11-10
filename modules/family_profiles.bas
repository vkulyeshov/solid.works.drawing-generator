Attribute VB_Name = "family_profiles"
Option Explicit

Private Function FamilyReferenceMargin() As Double
    FamilyReferenceMargin = FAMILY_REFERENCE_DETAIL_MARGIN_M
    If sheetWidth > sheetHeight Then
        FamilyReferenceMargin = FAMILY_LANDSCAPE_REFERENCE_MARGIN_M
    End If
End Function

' A support family keeps a complete row of recesses together. The allocator
' reserves text bands for that row before it creates a native drawing view.
Private Function FamilySupportKey(ByVal feature As Object) As String
    Dim a As Object, b As Object, dx As Double, dy As Double, length As Double
    Set a = feature("first")
    Set b = feature("last")
    length = PanelDistance(a, b)
    dx = (b("x") - a("x")) / length
    dy = (b("y") - a("y")) / length
    If dx < -DIRECTION_ZERO_TOLERANCE Or (Abs(dx) <= DIRECTION_ZERO_TOLERANCE And dy < 0) Then
        dx = -dx
        dy = -dy
    End If
    FamilySupportKey = CStr(Round(dx, DIRECTION_KEY_DECIMALS)) & "|" & CStr(Round(dy, DIRECTION_KEY_DECIMALS)) & "|" & ProfileCoordinateKey(-a("x") * dy + a("y") * dx)
End Function

Private Function FamilyProfileGroups() As Collection
    Dim result As New Collection, groups As Object, ids As Object, group As Object, feature As Variant, request As Variant
    Dim node As Variant, points As Variant, a As Object, b As Object, point As Variant, bounds As Variant, names As Variant
    Dim index As Long, id As String, key As String, area As Double, angle As Double
    Set groups = CreateObject("Scripting.Dictionary")
    Set ids = CreateObject("Scripting.Dictionary")
    For index = 1 To profilePanelFrame.Count
        Set a = PanelNode(profilePanelFrame, index)
        Set b = PanelNode(profilePanelFrame, index + 1)
        area = area + a("x") * b("y") - b("x") * a("y")
    Next index
    index = 0
    For Each feature In profilePanelFeatures
        index = index + 1
        id = "notch" & index
        key = FamilySupportKey(feature)
        If Not groups.Exists(key) Then
            Set group = CompactGroup("edge", 1)
            Set a = feature("first")
            Set b = feature("last")
            angle = -PlanarAngle(b("x") - a("x"), b("y") - a("y"))
            If area > 0 Then
                angle = angle + PI
            End If
            group("rotation") = angle
            groups.Add key, group
            result.Add group
        End If
        ids.Add id, groups(key)
    Next feature
    index = 0
    For Each points In profilePanelOpenings
        index = index + 1
        id = "opening" & index
        Set group = CompactGroup("opening", 1)
        ids.Add id, group
        result.Add group
    Next points
    For Each request In profilePagePlan
        If request("role") = PROFILE_ROLE_SLOT Then
            names = request("features").Keys
            id = CStr(names(0))
            Set group = ids(id)
            If request("value") < COMPACT_MICRO_LIMIT_M And request("kind") <> "notch-location" Then
                key = "micro|" & id
                If Not groups.Exists(key) Then
                    groups.Add key, CompactGroup("micro", 1)
                    result.Add groups(key)
                End If
                Set group = groups(key)
            End If
            group("requests").Add request
        End If
    Next request
    For Each group In result
        bounds = Array(SEARCH_INITIAL_LIMIT, SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT)
        For Each request In group("requests")
            point = ProfilePagePoint(request("first"))
            bounds = ProfilePageUnion(bounds, Array(point(0), point(1), point(0), point(1)))
            point = ProfilePagePoint(request("last"))
            bounds = ProfilePageUnion(bounds, Array(point(0), point(1), point(0), point(1)))
            If group("type") <> "micro" Then
                names = request("features").Keys
                id = CStr(names(0))
                If Left$(id, Len("notch")) = "notch" Then
                    Set feature = profilePanelFeatures(CLng(Mid$(id, Len("notch") + 1)))
                    Set points = feature("raw")
                Else
                    Set points = profilePanelOpenings(CLng(Mid$(id, Len("opening") + 1)))
                End If
                For Each node In points
                    point = ProfilePagePoint(node("point"))
                    bounds = ProfilePageUnion(bounds, Array(point(0), point(1), point(0), point(1)))
                Next node
            End If
        Next request
        group.Add "core", bounds
    Next group
    Set FamilyProfileGroups = result
End Function

Private Sub FamilySizeGroup(ByVal group As Object, ByVal work As Variant)
    Dim core As Variant, bounds As Variant, rotated As Variant, width As Double, height As Double, viewScaleValue As Double
    Dim request As Variant, minimum As Double, depthCount As Long, rightBand As Double, index As Long, margin As Double, availableHeight As Double
    core = group("core")
    rotated = CompactRotateBounds(core, group("rotation"))
    minimum = SEARCH_INITIAL_LIMIT
    For Each request In group("requests")
        If request("kind") = "notch-width" Or request("kind") = "opening-size" Then
            minimum = MinD(minimum, request("value"))
        End If
        If request("kind") = "notch-depth" Or request("kind") = "notch-oblique-edge" Then
            depthCount = depthCount + 1
        End If
    Next request
    rightBand = FAMILY_DETAIL_RIGHT_BAND_M + -MinD(0, -(depthCount - 1)) * COORDINATE_LANE_PITCH_M
    width = rotated(2) - rotated(0)
    height = rotated(3) - rotated(1)
    viewScaleValue = FAMILY_DETAIL_TARGET_SCALE
    If minimum < SEARCH_INITIAL_LIMIT Then
        viewScaleValue = -MinD(-viewScaleValue, -MinD(SCALE_10_TO_1, FAMILY_DETAIL_TARGET_INTERVAL_M / minimum))
    End If
    viewScaleValue = MinD(viewScaleValue, (work(2) - work(0) - FAMILY_DETAIL_LEFT_BAND_M - rightBand) / -MinD(-width, -PROFILE_COORDINATE_TOLERANCE_M))
    availableHeight = sheetHeight - FamilyReferenceMargin() - PROFILE_PAGE_REGION_GAP_M - FAMILY_DETAIL_BOTTOM_M - FAMILY_DETAIL_CONTEXT_PAPER_M * 2 - FAMILY_DETAIL_BOTTOM_BAND_M - FAMILY_DETAIL_TOP_BAND_M - PROFILE_PAGE_CAPTION_BAND_M
    viewScaleValue = StandardScale(MinD(viewScaleValue, MinD(FAMILY_DETAIL_HEIGHT_M, availableHeight) / -MinD(-height, -PROFILE_COORDINATE_TOLERANCE_M)))
    margin = FAMILY_DETAIL_CONTEXT_PAPER_M / viewScaleValue
    bounds = core
    For index = 0 To 1
        bounds(index) = bounds(index) - margin
        bounds(index + 2) = bounds(index + 2) + margin
    Next index
    group.Add "bounds", bounds
    group.Add "scale", viewScaleValue
    group.Add "width", width * viewScaleValue + FAMILY_DETAIL_CONTEXT_PAPER_M * 2 + FAMILY_DETAIL_LEFT_BAND_M + rightBand
    group.Add "height", height * viewScaleValue + FAMILY_DETAIL_CONTEXT_PAPER_M * 2 + FAMILY_DETAIL_BOTTOM_BAND_M + FAMILY_DETAIL_TOP_BAND_M + PROFILE_PAGE_CAPTION_BAND_M
End Sub

Private Sub BuildFamilyProfilePages(ByVal d As SldWorks.ModelDoc2)
    Dim doc As SldWorks.ModelDoc2, body As SldWorks.Body2, overview As SldWorks.View, request As Variant, group As Variant
    Dim work As Variant, cell As Variant, outline As Variant, center As Variant, x As Double, y As Double, rowHeight As Double
    Dim lanes As New Collection, page As Long, viewScaleValue As Double, index As Long, sorted As New Collection, largest As Object, used As Object
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
    Set profilePageGroups = FamilyProfileGroups()
    AsDrawing(d).ActivateSheet profilePageFirstSheet
    Set overview = ProfilePageOverview(d, doc, body)
    profileCompactEdgeCount = VisibleEdgeCount(overview)
    overview.ScaleDecimal = 1
    d.ForceRebuild3 False
    viewScaleValue = FitScale(overview, sheetWidth - LEADER_WORK_LEFT_M - LEADER_WORK_EDGE_MARGIN_M - PROFILE_PAGE_MAIN_BAND_M * 2, FAMILY_OVERVIEW_HEIGHT_M)
    overview.ScaleDecimal = viewScaleValue
    SetPaperScale AsDrawing(d).GetCurrentSheet, viewScaleValue
    d.ForceRebuild3 False
    CenterView overview, (LEADER_WORK_LEFT_M + sheetWidth - LEADER_WORK_EDGE_MARGIN_M) / 2, sheetHeight - PROFILE_PAGE_MAIN_CENTER_MARGIN_M
    outline = overview.GetOutline
    center = Array((outline(0) + outline(2)) / 2, (outline(1) + outline(3)) / 2)
    For Each request In profilePagePlan
        If request("role") <> PROFILE_ROLE_SLOT Then
            ProfileDimension d, overview, request, lanes, center
        End If
    Next request
    work = Array(LEADER_WORK_LEFT_M, FAMILY_DETAIL_BOTTOM_M, sheetWidth - LEADER_WORK_EDGE_MARGIN_M, sheetHeight - FAMILY_DETAIL_FIRST_TOP_MARGIN_M)
    For Each group In profilePageGroups
        If group("requests").Count > 0 Then
            FamilySizeGroup group, work
        End If
    Next group
    ' Long support families are placed before small corner details.
    Set used = CreateObject("Scripting.Dictionary")
    Do
        Set largest = Nothing
        For index = 1 To profilePageGroups.Count
            Set group = profilePageGroups(index)
            If group("requests").Count > 0 And Not used.Exists(index) Then
                If largest Is Nothing Then
                    Set largest = group
                ElseIf group("width") * group("height") > largest("width") * largest("height") Then
                    Set largest = group
                End If
            End If
        Next index
        If largest Is Nothing Then
            Exit Do
        End If
        sorted.Add largest
        For index = 1 To profilePageGroups.Count
            If profilePageGroups(index) Is largest Then
                used(index) = True
            End If
        Next index
    Loop
    Set profilePageGroups = sorted
    x = work(0)
    y = work(3)
    page = 1
    For Each group In profilePageGroups
        cell = Empty
        If group("height") <= work(3) - work(1) Then
            cell = ProfilePageCell(group("width"), group("height"), x, y, rowHeight, work)
        End If
        If IsEmpty(cell) Then
            page = page + 1
            NewFormattedSheet d, "Елементи " & page
            profilePageSheets.Add AsDrawing(d).GetCurrentSheet.GetName
            Set overview = ProfilePageOverview(d, doc, body)
            overview.ScaleDecimal = 1
            d.ForceRebuild3 False
            viewScaleValue = FAMILY_REFERENCE_HEIGHT_M
            If sheetWidth > sheetHeight Then
                viewScaleValue = viewScaleValue / 2
            End If
            overview.ScaleDecimal = FitScale(overview, sheetWidth - LEADER_WORK_LEFT_M - LEADER_WORK_EDGE_MARGIN_M, viewScaleValue)
            d.ForceRebuild3 False
            CenterView overview, (LEADER_WORK_LEFT_M + sheetWidth - LEADER_WORK_EDGE_MARGIN_M) / 2, sheetHeight - FamilyReferenceMargin() / 2
            profilePageRegions(overview.Name) = Array(LEADER_WORK_LEFT_M, sheetHeight - FamilyReferenceMargin(), sheetWidth - LEADER_WORK_EDGE_MARGIN_M, sheetHeight - PROFILE_PAGE_MAIN_TOP_MARGIN_M)
            work(3) = sheetHeight - FamilyReferenceMargin() - PROFILE_PAGE_REGION_GAP_M
            x = work(0)
            y = work(3)
            rowHeight = 0
            cell = ProfilePageCell(group("width"), group("height"), x, y, rowHeight, work)
        End If
        If IsEmpty(cell) Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "A complete support family exceeds the selected paper."
        End If
        group("page") = page
        group.Add "region", Array(cell(0), cell(1), cell(2), cell(3) - PROFILE_PAGE_CAPTION_BAND_M)
        outline = CompactRotateBounds(group("core"), group("rotation"))
        group.Add "center", Array(cell(0) + FAMILY_DETAIL_LEFT_BAND_M + FAMILY_DETAIL_CONTEXT_PAPER_M + (outline(2) - outline(0)) * group("scale") / 2, cell(1) + FAMILY_DETAIL_BOTTOM_BAND_M + FAMILY_DETAIL_CONTEXT_PAPER_M + (outline(3) - outline(1)) * group("scale") / 2)
        group.Add "configured", True
        CompactCreateView d, doc, body, group
    Next group
    AsDrawing(d).ActivateSheet profilePageFirstSheet
    Log "FAMILY PROFILE sheets=" & page & "; complete support groups=" & profilePageGroups.Count & "; dimensions=" & profilePagePlan.Count
End Sub

Private Sub FinishFamilyProfilePages(ByVal d As SldWorks.ModelDoc2)
    Dim sheetName As Variant, views As Collection, view As SldWorks.View, item As Variant, group As Variant, candidate As Variant
    Dim display As SldWorks.DisplayDimension, request As Object, records As Collection, bases As Collection, format As SldWorks.TextFormat
    Dim counters As Object, kind As String, index As Long, core As Variant, a As Variant, b As Variant, position As Variant, coordinate As Double
    Dim axis As Long, row As Long, attempt As Long, arrow As Variant, annotation As SldWorks.Annotation, region As Variant, length As Double
    For Each sheetName In AsDrawing(d).GetSheetNames
        AsDrawing(d).ActivateSheet CStr(sheetName)
        AsDrawing(d).ActivateView ""
        Set views = New Collection
        Set view = AsDrawing(d).GetFirstView.GetNextView
        Do While Not view Is Nothing
            views.Add view
            Set view = view.GetNextView
        Loop
        CacheProfileContours views
        For Each item In views
            Set view = item
            Set group = Nothing
            For Each candidate In profilePageGroups
                If candidate("view") = view.Name Then
                    Set group = candidate
                    Exit For
                End If
            Next candidate
            If group Is Nothing And VisibleEdgeCount(view) <> profileCompactEdgeCount Then
                Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "A full feature-family overview lost geometry."
            End If
            region = profilePageRegions(view.Name)
            Set records = New Collection
            Set bases = New Collection
            Set counters = CreateObject("Scripting.Dictionary")
            Set display = view.GetFirstDisplayDimension5
            Do While Not display Is Nothing
                Set request = profileDimensionRequests(display.GetNameForSelection)
                Set annotation = display.GetAnnotation
                Set format = annotation.GetTextFormat(SW_ANNOTATION_TEXT_FORMAT_FIRST)
                format.CharHeight = TEXT_HEIGHT_SMALL_M
                annotation.SetTextFormat SW_ANNOTATION_TEXT_FORMAT_FIRST, False, format
                If group Is Nothing Then
                    If Not ProfilePagePlaceOverview(d, display, view) Then
                        PlaceProfileDimensionBase d, display, view, region, bases
                    End If
                Else
                    a = PatternSheetPoint(view, request("first"))
                    b = PatternSheetPoint(view, request("last"))
                    axis = SCREEN_X_AXIS_INDEX
                    If Abs(b(0) - a(0)) < Abs(b(1) - a(1)) Then
                        axis = SCREEN_Y_AXIS_INDEX
                    End If
                    request("axis") = axis
                    request("sheetFirst") = a
                    request("sheetLast") = b
                    request("view") = view.Name
                    request("localBase") = False
                    request("externalWitness") = True
                    kind = request("kind")
                    If kind = "notch-oblique-edge" Then
                        request("axis") = SPACE_LAST_AXIS_INDEX
                        PlacePanelAlignedBase d, display, view, request, region
                        GoTo FamilyDirectionReady
                    End If
                    If Not counters.Exists(kind) Then
                        counters.Add kind, 0
                    End If
                    row = counters(kind)
                    counters(kind) = row + 1
                    core = group("paperCore")
                    coordinate = core(3) + GOST_OBJECT_OFFSET_M + (row Mod FAMILY_ROW_LIMIT) * COORDINATE_LANE_PITCH_M
                    If kind = "notch-location" Then
                        coordinate = coordinate + FAMILY_ROW_LIMIT * COORDINATE_LANE_PITCH_M
                    End If
                    If axis = SCREEN_Y_AXIS_INDEX Then
                        coordinate = core(2) + GOST_OBJECT_OFFSET_M + TEXT_HEIGHT_SMALL_M + row * COORDINATE_LANE_PITCH_M
                    End If
                    display.OffsetText = False
                    display.CenterText = True
                    If Abs(b(axis) - a(axis)) >= 2 * GOST_ARROW_HEAD_LENGTH_M Then
                        display.ArrowSide = SW_DIMENSION_ARROWS_INSIDE
                    Else
                        display.ArrowSide = SW_DIMENSION_ARROWS_OUTSIDE
                    End If
                    position = Array((a(0) + b(0)) / 2, (a(1) + b(1)) / 2, 0#)
                    position(1 - axis) = coordinate
                    annotation.SetPosition2 position(0), position(1), 0
                    display.CenterText = True
                    For attempt = 1 To LEADER_CLAMP_ATTEMPTS
                        arrow = display.GetDisplayData.GetArrowHeadAtIndex2(0)
                        If Abs(arrow(1 - axis) - coordinate) < PROFILE_MATERIAL_TOLERANCE_M Then
                            Exit For
                        End If
                        position = annotation.GetPosition
                        position(1 - axis) = position(1 - axis) + coordinate - arrow(1 - axis)
                        annotation.SetPosition2 position(0), position(1), position(2)
                        display.CenterText = True
                    Next attempt
                End If
FamilyDirectionReady:
                If request("axis") = SPACE_LAST_AXIS_INDEX Then
                    a = PatternSheetPoint(view, request("first"))
                    b = PatternSheetPoint(view, request("last"))
                    length = Sqr((b(0) - a(0)) ^ 2 + (b(1) - a(1)) ^ 2)
                    request("compactDirection") = Array((b(0) - a(0)) / length, (b(1) - a(1)) / length)
                ElseIf request("axis") = SCREEN_X_AXIS_INDEX Then
                    request("compactDirection") = Array(1#, 0#)
                Else
                    request("compactDirection") = Array(0#, 1#)
                End If
                records.Add display
                Set display = display.GetNext5
            Loop
            If group Is Nothing And records.Count > 0 Then
                ProfilePageFitOverview d, view, records, region
            End If
            CompactPackLabels d, view, records, region
            Log "FAMILY fitted " & view.Name & "; dimensions=" & records.Count
        Next item
    Next sheetName
    For Each group In profilePageGroups
        AsDrawing(d).ActivateSheet CStr(group("sheet"))
        region = group("region")
        AddText d, "Елементи " & group("number") & " (" & ScaleText(group("scale")) & ")", region(0), region(3) + PROFILE_PAGE_CAPTION_GAP_M, TEXT_HEIGHT_SMALL_M, 0
    Next group
    AuditProfilePages d
    AsDrawing(d).ActivateSheet profilePageFirstSheet
    Log "FAMILY COMPLETE sheets=" & AsDrawing(d).GetSheetCount & "; dimensions=" & profilePagePlan.Count
End Sub
