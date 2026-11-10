Attribute VB_Name = "profile_pages"
Option Explicit

Private Function ProfilePageMode() As Boolean
    Dim mode As String
    mode = LCase$(Setting("ProfileSheets", "auto"))
    If mode <> "auto" And mode <> "legacy" Then
        Err.Raise vbObjectError + ERR_DRAWING_SETTINGS, , "ProfileSheets must be auto or legacy."
    End If
    ProfilePageMode = mode = "auto"
End Function

Private Function ProfilePagePoint(ByVal point As Variant) As Variant
    Dim m As Variant
    m = profilePageMatrix
    ProfilePagePoint = Array(point(0) * m(MATRIX_XX_INDEX) + point(1) * m(MATRIX_YX_INDEX) + point(2) * m(MATRIX_ZX_INDEX), point(0) * m(MATRIX_XY_INDEX) + point(1) * m(MATRIX_YY_INDEX) + point(2) * m(MATRIX_ZY_INDEX))
End Function

' Cropped views can expose trimmed display edges without their original end
' vertices. Select the saved body topology in that drawing-view context.
Private Function ProfilePageReference(ByVal view As SldWorks.View, ByVal point As Variant) As SldWorks.Vertex
    Dim doc As SldWorks.ModelDoc2, body As SldWorks.Body2, vertices As Variant, item As Variant
    Dim vertex As SldWorks.Vertex, target As Variant, actual As Variant, bounds As Variant
    target = PatternSheetPoint(view, point)
    bounds = profilePageCrops(view.Name)
    If Not NativeBoundsInside(Array(target(0), target(1), target(0), target(1)), bounds) Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "A cropped reference is outside the native viewport: " & view.Name & "; point=" & CStr(target(0)) & "/" & CStr(target(1)) & "; viewport=" & CStr(bounds(0)) & "/" & CStr(bounds(1)) & "/" & CStr(bounds(2)) & "/" & CStr(bounds(3))
    End If
    Set doc = sw.GetOpenDocumentByName(profilePageModel)
    Set body = FindBody(doc, profilePageBody)
    vertices = body.GetVertices
    For Each item In vertices
        Set vertex = item
        actual = PatternSheetPoint(view, vertex.GetPoint)
        If Abs(actual(0) - target(0)) <= PROFILE_COORDINATE_TOLERANCE_M * view.ScaleDecimal And Abs(actual(1) - target(1)) <= PROFILE_COORDINATE_TOLERANCE_M * view.ScaleDecimal Then
            Set ProfilePageReference = vertex
            Exit Function
        End If
    Next item
End Function

Private Function ProfilePageRequestBounds(ByVal request As Object) As Variant
    Dim a As Variant, b As Variant, bounds As Variant, raw As Collection, point As Variant, neighbor As Variant, projected As Variant
    Dim index As Long, offset As Long, distance As Double, limit As Double, path As Variant
    a = ProfilePagePoint(request("first"))
    b = ProfilePagePoint(request("last"))
    bounds = Array(MinD(a(0), b(0)), MinD(a(1), b(1)), -MinD(-a(0), -b(0)), -MinD(-a(1), -b(1)))
    path = ProfilePageBoundaryBounds(request("first"), request("last"))
    If Not IsEmpty(path) Then
        bounds = ProfilePageUnion(bounds, path)
    End If
    limit = MinD(PROFILE_PAGE_NEIGHBOR_LIMIT_M, request("value") * PROFILE_PAGE_NEIGHBOR_FACTOR)
    ' Include the adjacent slot walls, even when their shared depth dimension
    ' is owned by a different representative view.
    For Each raw In profileRawLoops
        For index = 1 To raw.Count
            point = raw(index)
            If ProfilePointClose(point, request("first")) Or ProfilePointClose(point, request("last")) Then
                For offset = -1 To 1 Step 2
                    neighbor = raw((index + offset - 1 + raw.Count) Mod raw.Count + 1)
                    distance = Sqr((point(0) - neighbor(0)) ^ 2 + (point(1) - neighbor(1)) ^ 2 + (point(2) - neighbor(2)) ^ 2)
                    If distance <= limit Then
                        projected = ProfilePagePoint(neighbor)
                        bounds = ProfilePageUnion(bounds, Array(projected(0), projected(1), projected(0), projected(1)))
                    End If
                Next offset
            End If
        Next index
    Next raw
    ProfilePageRequestBounds = bounds
End Function

' A straight distance can span an open major cutout. Include the shortest
' real boundary path, so its cropped context retains the connecting bridge.
Private Function ProfilePageBoundaryBounds(ByVal first As Variant, ByVal last As Variant) As Variant
    Dim raw As Collection, a As Variant, b As Variant, p As Variant, bounds As Variant, bestBounds As Variant
    Dim i As Long, start As Long, finish As Long, direction As Long, index As Long, nextIndex As Long, length As Double, best As Double
    best = SEARCH_INITIAL_LIMIT
    For Each raw In profileRawLoops
        start = 0
        finish = 0
        For i = 1 To raw.Count
            If ProfilePointClose(raw(i), first) Then
                start = i
            End If
            If ProfilePointClose(raw(i), last) Then
                finish = i
            End If
        Next i
        If start > 0 And finish > 0 Then
            For direction = -1 To 1 Step 2
                p = ProfilePagePoint(raw(start))
                bounds = Array(p(0), p(1), p(0), p(1))
                length = 0
                index = start
                Do While index <> finish
                    nextIndex = (index + direction - 1 + raw.Count) Mod raw.Count + 1
                    a = raw(index)
                    b = raw(nextIndex)
                    length = length + Sqr((a(0) - b(0)) ^ 2 + (a(1) - b(1)) ^ 2 + (a(2) - b(2)) ^ 2)
                    p = ProfilePagePoint(b)
                    bounds = ProfilePageUnion(bounds, Array(p(0), p(1), p(0), p(1)))
                    index = nextIndex
                Loop
                If length < best Then
                    best = length
                    bestBounds = bounds
                End If
            Next direction
        End If
    Next raw
    ProfilePageBoundaryBounds = bestBounds
End Function

Private Function ProfilePageSupport(ByVal request As Object) As String
    Dim a As Variant, b As Variant, dx As Double, dy As Double, length As Double
    a = ProfilePagePoint(request("first"))
    b = ProfilePagePoint(request("last"))
    dx = b(0) - a(0)
    dy = b(1) - a(1)
    length = Sqr(dx * dx + dy * dy)
    dx = dx / length
    dy = dy / length
    If dx < -DIRECTION_ZERO_TOLERANCE Or (Abs(dx) <= DIRECTION_ZERO_TOLERANCE And dy < 0) Then
        dx = -dx
        dy = -dy
    End If
    ProfilePageSupport = CStr(Round(dx, DIRECTION_KEY_DECIMALS)) & "|" & CStr(Round(dy, DIRECTION_KEY_DECIMALS)) & "|" & ProfileCoordinateKey(-a(0) * dy + a(1) * dx)
End Function

Private Function ProfilePageUnion(ByVal a As Variant, ByVal b As Variant) As Variant
    ProfilePageUnion = Array(MinD(a(0), b(0)), MinD(a(1), b(1)), -MinD(-a(2), -b(2)), -MinD(-a(3), -b(3)))
End Function

' Group nearby measured features, never the complete panel bounding box.
' A split preserves the same request objects and their unique physical keys.
Private Function ProfilePagePlanGroups(ByVal plan As Collection, ByVal limit As Long) As Collection
    Dim groups As New Collection, request As Variant, group As Object, chosen As Object
    Dim bounds As Variant, combined As Variant, requests As Collection, cost As Double, best As Double
    Dim supports As Object, ids As Variant, id As String, support As String, partition As String
    Set supports = CreateObject("Scripting.Dictionary")
    For Each request In plan
        If request.Exists("panel") And request("role") = PROFILE_ROLE_SLOT Then
            If request("kind") = "notch-width" Or request("kind") = "notch-location" Then
                ids = request("features").Keys
                id = CStr(ids(0))
                supports(id) = ProfilePageSupport(request)
            End If
        End If
    Next request
    For Each request In plan
        If request("role") = PROFILE_ROLE_SLOT Then
            ids = request("features").Keys
            id = CStr(ids(0))
            support = id
            If supports.Exists(id) Then
                support = supports(id)
            End If
            bounds = ProfilePageRequestBounds(request)
            partition = ""
            If request.Exists("pagePartition") Then
                partition = request("pagePartition")
            End If
            best = SEARCH_INITIAL_LIMIT
            Set chosen = Nothing
            For Each group In groups
                Set requests = group("requests")
                If requests.Count < limit And group("support") = support And group("partition") = partition Then
                    combined = ProfilePageUnion(bounds, group("bounds"))
                    cost = (combined(2) - combined(0)) + (combined(3) - combined(1))
                    If cost < best And cost <= PROFILE_PAGE_GROUP_SPAN_M Then
                        Set chosen = group
                        best = cost
                    End If
                End If
            Next group
            If chosen Is Nothing Then
                Set chosen = CreateObject("Scripting.Dictionary")
                Set requests = New Collection
                chosen.Add "requests", requests
                chosen.Add "bounds", bounds
                chosen.Add "support", support
                chosen.Add "partition", partition
                chosen.Add "padding", 0
                groups.Add chosen
            End If
            If request.Exists("pagePadding") Then
                chosen("padding") = -MinD(-chosen("padding"), -request("pagePadding"))
            End If
            chosen("bounds") = ProfilePageUnion(chosen("bounds"), bounds)
            chosen("requests").Add request
        End If
    Next request
    Set ProfilePagePlanGroups = groups
End Function

' A shelf allocator reserves geometry AND dimension bands. An oversized cell
' is an error; a full page starts another sheet without reducing text height.
Private Function ProfilePageCell(ByVal width As Double, ByVal height As Double, ByRef x As Double, ByRef y As Double, ByRef rowHeight As Double, ByVal area As Variant) As Variant
    If width > area(2) - area(0) Or height > area(3) - area(1) Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "An enlarged profile cell exceeds the selected sheet."
    End If
    If x + width > area(2) Then
        x = area(0)
        y = y - rowHeight - PROFILE_PAGE_CELL_GAP_M
        rowHeight = 0
    End If
    If y - height < area(1) Then
        Exit Function
    End If
    ProfilePageCell = Array(x, y - height, x + width, y)
    x = x + width + PROFILE_PAGE_CELL_GAP_M
    rowHeight = -MinD(-rowHeight, -height)
End Function

Private Function ProfilePageBands(ByVal group As Object, ByVal band As Double) As Variant
    Dim request As Variant, horizontal As Boolean, vertical As Boolean, aligned As Boolean, bands As Variant
    For Each request In group("requests")
        If request("axis") = SCREEN_X_AXIS_INDEX Then
            horizontal = True
        ElseIf request("axis") = SCREEN_Y_AXIS_INDEX Then
            vertical = True
        Else
            aligned = True
        End If
    Next request
    bands = Array(band, band)
    If Not vertical And Not aligned Then
        bands(0) = PROFILE_PAGE_CONTEXT_BAND_M
    End If
    If Not horizontal And Not aligned Then
        bands(1) = PROFILE_PAGE_CONTEXT_BAND_M
    End If
    ProfilePageBands = bands
End Function

Private Function ProfilePageMapBounds(ByVal bounds As Variant, ByVal reference As Variant, ByVal anchor As Variant, ByVal viewScaleValue As Double) As Variant
    ProfilePageMapBounds = Array(anchor(0) + (bounds(0) - reference(0)) * viewScaleValue, anchor(1) + (bounds(1) - reference(1)) * viewScaleValue, anchor(0) + (bounds(2) - reference(0)) * viewScaleValue, anchor(1) + (bounds(3) - reference(1)) * viewScaleValue)
End Function

Private Sub ProfilePageCrop(ByVal d As SldWorks.ModelDoc2, ByVal view As SldWorks.View, ByVal request As Object, ByVal bounds As Variant, Optional ByVal rotation As Double = 0)
    Dim anchor As Variant, reference As Variant, corner As Variant, lower As Variant, upper As Variant, result As Long
    Dim util As SldWorks.MathUtility, sketch As SldWorks.Sketch, point As SldWorks.MathPoint, rectangle As Variant
    Dim coordinates(2) As Double, viewScaleValue As Double
    Dim manager As SldWorks.SketchManager, oldAddToDB As Boolean, oldFocus As Boolean
    If AsDrawing(d).ActivateView(view.Name) = False Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot activate the profile crop view."
    End If
    anchor = PatternSheetPoint(view, request("first"))
    reference = ProfilePagePoint(request("first"))
    If Abs(rotation) > DIRECTION_ZERO_TOLERANCE Then
        reference = CompactRotatePoint(reference, rotation)
    End If
    viewScaleValue = view.ScaleDecimal
    lower = Array(anchor(0) + (bounds(0) - reference(0)) * viewScaleValue, anchor(1) + (bounds(1) - reference(1)) * viewScaleValue)
    upper = Array(anchor(0) + (bounds(2) - reference(0)) * viewScaleValue, anchor(1) + (bounds(3) - reference(1)) * viewScaleValue)
    Set sketch = view.GetSketch
    Set util = sw.GetMathUtility
    coordinates(0) = lower(0)
    coordinates(1) = lower(1)
    Set point = util.CreatePoint(coordinates)
    Set point = point.MultiplyTransform(sketch.ModelToSketchTransform)
    lower = point.ArrayData
    coordinates(0) = upper(0)
    coordinates(1) = upper(1)
    Set point = util.CreatePoint(coordinates)
    Set point = point.MultiplyTransform(sketch.ModelToSketchTransform)
    upper = point.ArrayData
    Log "PROFILE CROP " & view.Name & "; scale=" & viewScaleValue & "; model bounds=" & Join(bounds, ",") & "; sketch lower=" & CStr(lower(0)) & "/" & CStr(lower(1)) & "; upper=" & CStr(upper(0)) & "/" & CStr(upper(1))
    d.ClearSelection2 True
    AsDrawing(d).ActivateView view.Name
    ' Drawing sketch commands can dynamically activate a neighboring view.
    ' Cropping must attach its rectangle to this view, including after rotation.
    oldFocus = view.FocusLocked
    view.FocusLocked = True
    Set manager = d.SketchManager
    oldAddToDB = manager.AddToDB
    manager.AddToDB = True
    rectangle = manager.CreateCornerRectangle(lower(0), upper(1), 0, upper(0), lower(1), 0)
    manager.AddToDB = oldAddToDB
    ' A jagged crop can cut off vertices near its nominal rectangle. Keep
    ' the exact rectangular mask and hide its non-model boundary instead.
    result = view.Crop2(False, True, PROFILE_PAGE_CROP_INTENSITY)
    If result <> PROFILE_PAGE_CROP_OK Then
        result = view.Crop
    End If
    view.FocusLocked = oldFocus
    d.ClearSelection2 True
    If result <> PROFILE_PAGE_CROP_OK Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Native profile crop failed: " & result
    End If
End Sub

Private Function ProfilePageOverview(ByVal d As SldWorks.ModelDoc2, ByVal doc As SldWorks.ModelDoc2, ByVal body As SldWorks.Body2) As SldWorks.View
    Dim view As SldWorks.View, area As Variant, viewScaleValue As Double
    If profilePageInitialView Is Nothing Then
        Set view = MakeView(d, doc, profilePageConfiguration, body, profilePageOrientation, sheetWidth / 2, sheetHeight - PROFILE_PAGE_MAIN_CENTER_MARGIN_M)
    Else
        Set view = profilePageInitialView
        Set profilePageInitialView = Nothing
    End If
    If view Is Nothing Then
        Err.Raise vbObjectError + ERR_MAIN_VIEW, , "Cannot create the complete profile overview."
    End If
    view.Angle = profilePageAngle
    view.UseSheetScale = False
    view.ScaleDecimal = 1
    d.ForceRebuild3 False
    viewScaleValue = FitScale(view, sheetWidth - LEADER_WORK_LEFT_M - LEADER_WORK_EDGE_MARGIN_M - PROFILE_PAGE_MAIN_BAND_M * 2, PROFILE_PAGE_MAIN_HEIGHT_M)
    view.ScaleDecimal = viewScaleValue
    SetPaperScale AsDrawing(d).GetCurrentSheet, viewScaleValue
    d.ForceRebuild3 False
    CenterView view, (LEADER_WORK_LEFT_M + sheetWidth - LEADER_WORK_EDGE_MARGIN_M) / 2, sheetHeight - PROFILE_PAGE_MAIN_CENTER_MARGIN_M
    area = Array(LEADER_WORK_LEFT_M, sheetHeight - PROFILE_PAGE_DETAIL_TOP_MARGIN_M + PROFILE_PAGE_REGION_GAP_M, sheetWidth - LEADER_WORK_EDGE_MARGIN_M, sheetHeight - PROFILE_PAGE_MAIN_TOP_MARGIN_M)
    profilePageRegions.Add view.Name, area
    profileGeometryViews.Add view.Name, True
    profilePageViews.Add Array(AsDrawing(d).GetCurrentSheet.GetName, view.Name)
    Set ProfilePageOverview = view
End Function

Private Sub BuildProfilePages(ByVal d As SldWorks.ModelDoc2)
    Dim doc As SldWorks.ModelDoc2, body As SldWorks.Body2, overview As SldWorks.View, view As SldWorks.View
    Dim group As Variant, request As Variant, first As Object, bounds As Variant, cell As Variant, area As Variant, outline As Variant, center As Variant, bands As Variant
    Dim anchor As Variant, reference As Variant, position As Variant
    Dim viewScaleValue As Double, band As Double, span As Double, width As Double, height As Double, x As Double, y As Double, rowHeight As Double, cropMargin As Double
    Dim page As Long, count As Long, lanes As Collection, index As Long, audit As Object
    Dim groupIndex As Long, splitIndex As Long, splitGroups As Collection, splitGroup As Variant
    Dim savedX As Double, savedY As Double, savedRowHeight As Double, failure As Long, failureMessage As String
    Set doc = sw.GetOpenDocumentByName(profilePageModel)
    If doc Is Nothing Then
        Err.Raise vbObjectError + ERR_OPEN_JOB_MODEL, , "The copied profile model is no longer open."
    End If
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
    Set profilePageGroups = ProfilePagePlanGroups(profilePagePlan, profilePageLimit)
    For Each request In profilePagePlan
        If request.Exists("materialBounds") Then
            request.Remove "materialBounds"
        End If
    Next request
    Log "PROFILE PAGE build groups=" & profilePageGroups.Count
    AsDrawing(d).ActivateSheet profilePageFirstSheet
    Set overview = ProfilePageOverview(d, doc, body)
    Log "PROFILE PAGE overview=" & overview.Name
    Set lanes = New Collection
    outline = overview.GetOutline
    center = Array((outline(0) + outline(2)) / 2, (outline(1) + outline(3)) / 2)
    For Each request In profilePagePlan
        If request("role") <> PROFILE_ROLE_SLOT Then
            ProfileDimension d, overview, request, lanes, center
        End If
    Next request
    ' Reject an invalid complete overview before building its detail sheets.
    FinishPagedDimensionLayout d
    area = Array(LEADER_WORK_LEFT_M, PROFILE_PAGE_BOTTOM_M, sheetWidth - LEADER_WORK_EDGE_MARGIN_M, sheetHeight - PROFILE_PAGE_DETAIL_TOP_MARGIN_M)
    x = area(0)
    y = area(3)
    page = 1
    groupIndex = 1
    Do While groupIndex <= profilePageGroups.Count
        Set group = profilePageGroups(groupIndex)
        count = count + 1
        Set first = group("requests")(1)
        bounds = group("bounds")
        cropMargin = PROFILE_PAGE_CROP_MARGIN_M
        If group.Exists("cropRetries") Then
            cropMargin = cropMargin * PROFILE_PAGE_CROP_MARGIN_FACTOR ^ group("cropRetries")
        End If
        For index = 0 To 1
            bounds(index) = bounds(index) - cropMargin
            bounds(index + 2) = bounds(index + 2) + cropMargin
        Next index
        band = PROFILE_PAGE_BASE_BAND_M - Int(-group("requests").Count / PROFILE_PAGE_SIDES_PER_BAND) * COORDINATE_LANE_PITCH_M + group("padding") * COORDINATE_LANE_PITCH_M
        bands = ProfilePageBands(group, band)
        span = SEARCH_INITIAL_LIMIT
        For Each request In group("requests")
            span = MinD(span, request("value"))
        Next request
        viewScaleValue = MinD(SCALE_10_TO_1, -MinD(-SCALE_1_TO_4, -PROFILE_PAGE_MIN_SPAN_M / span))
        viewScaleValue = MinD(viewScaleValue, (area(2) - area(0) - bands(0) * 2) / (bounds(2) - bounds(0)))
        viewScaleValue = StandardScale(MinD(viewScaleValue, (area(3) - area(1) - bands(1) * 2 - PROFILE_PAGE_CAPTION_BAND_M) / (bounds(3) - bounds(1))))
        For index = 0 To 1
            If (bounds(index + 2) - bounds(index)) * viewScaleValue < PROFILE_PAGE_MIN_CROP_PAPER_M Then
                span = (bounds(index) + bounds(index + 2)) / 2
                bounds(index) = span - PROFILE_PAGE_MIN_CROP_PAPER_M / viewScaleValue / 2
                bounds(index + 2) = span + PROFILE_PAGE_MIN_CROP_PAPER_M / viewScaleValue / 2
            End If
        Next index
        width = (bounds(2) - bounds(0)) * viewScaleValue + bands(0) * 2
        height = (bounds(3) - bounds(1)) * viewScaleValue + bands(1) * 2 + PROFILE_PAGE_CAPTION_BAND_M
        savedX = x
        savedY = y
        savedRowHeight = rowHeight
        cell = ProfilePageCell(width, height, x, y, rowHeight, area)
        If IsEmpty(cell) Then
            page = page + 1
            NewFormattedSheet d, "Вирізи " & page
            profilePageSheets.Add AsDrawing(d).GetCurrentSheet.GetName
            Set overview = ProfilePageOverview(d, doc, body)
            x = area(0)
            y = area(3)
            rowHeight = 0
            savedX = x
            savedY = y
            savedRowHeight = rowHeight
            cell = ProfilePageCell(width, height, x, y, rowHeight, area)
        End If
        Set view = MakeView(d, doc, profilePageConfiguration, body, profilePageOrientation, (cell(0) + cell(2)) / 2, (cell(1) + cell(3)) / 2)
        view.Angle = profilePageAngle
        view.UseSheetScale = False
        view.ScaleDecimal = viewScaleValue
        d.ForceRebuild3 False
        anchor = PatternSheetPoint(view, first("first"))
        reference = ProfilePagePoint(first("first"))
        position = view.Position
        position(0) = position(0) + (cell(0) + cell(2)) / 2 - anchor(0) - ((bounds(0) + bounds(2)) / 2 - reference(0)) * viewScaleValue
        position(1) = position(1) + (cell(1) + cell(3) - PROFILE_PAGE_CAPTION_BAND_M) / 2 - anchor(1) - ((bounds(1) + bounds(3)) / 2 - reference(1)) * viewScaleValue
        view.Position = position
        d.ForceRebuild3 False
        Log "PROFILE PAGE creating group=" & count & "; scale=" & viewScaleValue & "; bounds=" & Join(bounds, ",")
        On Error Resume Next
        ProfilePageCrop d, view, first, bounds
        failure = Err.Number
        failureMessage = Err.Description
        Err.Clear
        On Error GoTo 0
        If failure <> 0 Then
            If failure <> vbObjectError + ERR_DIMENSION_LAYOUT Then
                Err.Raise failure, "BuildProfilePages", failureMessage
            End If
            If Not group.Exists("cropRetries") Then
                group.Add "cropRetries", 0
            End If
            If group("cropRetries") >= PROFILE_PAGE_CROP_RETRIES Then
                Err.Raise failure, "BuildProfilePages", failureMessage
            End If
            group("cropRetries") = group("cropRetries") + 1
            Log "PROFILE PAGE retrying native crop with more context; retry=" & group("cropRetries")
            RemoveProfilePageView d, view, False
            x = savedX
            y = savedY
            rowHeight = savedRowHeight
            count = count - 1
            GoTo NextProfilePageGroup
        End If
        d.ForceRebuild3 False
        view.UpdateViewDisplayGeometry
        ' GetOutline is the view boundary, not the crop mask. Map the actual
        ' rectangle through the current native transform after cropping.
        profilePageCrops.Add view.Name, ProfilePageMapBounds(bounds, ProfilePagePoint(first("first")), PatternSheetPoint(view, first("first")), viewScaleValue)
        profilePageRegions.Add view.Name, Array(cell(0), cell(1), cell(2), cell(3) - PROFILE_PAGE_CAPTION_BAND_M)
        profileGeometryViews.Add view.Name, True
        profilePageViews.Add Array(AsDrawing(d).GetCurrentSheet.GetName, view.Name)
        group.Add "sheet", AsDrawing(d).GetCurrentSheet.GetName
        group.Add "view", view.Name
        group.Add "cell", cell
        group.Add "scale", viewScaleValue
        group.Add "number", count
        Set lanes = New Collection
        outline = view.GetOutline
        center = Array((outline(0) + outline(2)) / 2, (outline(1) + outline(3)) / 2)
        For Each request In group("requests")
            ProfileDimension d, view, request, lanes, center
        Next request
        On Error Resume Next
        FinishPagedDimensionLayout d, view.Name
        failure = Err.Number
        failureMessage = Err.Description
        Err.Clear
        On Error GoTo 0
        If failure <> 0 Then
            If failure <> vbObjectError + ERR_DIMENSION_LAYOUT Then
                Err.Raise failure, "BuildProfilePages", failureMessage
            End If
            If Not SplitProfilePageGroup(view.Name) Then
                Err.Raise failure, "BuildProfilePages", failureMessage
            End If
            Log "PROFILE PAGE immediate overflow: " & failureMessage
            RemoveProfilePageView d, view
            Set splitGroups = ProfilePagePlanGroups(group("requests"), profilePageLimit)
            profilePageGroups.Remove groupIndex
            For splitIndex = 1 To splitGroups.Count
                Set splitGroup = splitGroups(splitIndex)
                If groupIndex + splitIndex - 1 > profilePageGroups.Count Then
                    profilePageGroups.Add splitGroup
                Else
                    profilePageGroups.Add splitGroup, Before:=groupIndex + splitIndex - 1
                End If
            Next splitIndex
            x = savedX
            y = savedY
            rowHeight = savedRowHeight
            count = count - 1
            GoTo NextProfilePageGroup
        End If
        Log "PROFILE PAGE sheet=" & page & "; group=" & count & "; scale=" & viewScaleValue & "; requests=" & group("requests").Count
        groupIndex = groupIndex + 1
NextProfilePageGroup:
    Loop
    AsDrawing(d).ActivateSheet profilePageFirstSheet
End Sub

Private Sub RemoveProfilePageView(ByVal d As SldWorks.ModelDoc2, ByVal view As SldWorks.View, Optional ByVal registered As Boolean = True)
    Dim display As SldWorks.DisplayDimension, key As String, name As String, request As Object
    name = view.Name
    Set display = view.GetFirstDisplayDimension5
    Do While Not display Is Nothing
        key = display.GetNameForSelection
        featureDimensionOwners.Remove key
        featureDimensionPlacement.Remove key
        If featureDimensionLanes.Exists(key) Then
            featureDimensionLanes.Remove key
        End If
        Set request = profileDimensionRequests(key)
        If request.Exists("materialBounds") Then
            request.Remove "materialBounds"
        End If
        profileDimensionRequests.Remove key
        Set display = display.GetNext5
    Loop
    AsDrawing(d).ActivateView ""
    d.ClearSelection2 True
    If d.Extension.SelectByID2(name, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0) = False Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot select a rejected profile detail view."
    End If
    If d.Extension.DeleteSelection2(PROFILE_PAGE_DELETE_OPTIONS) = False Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot remove a rejected profile detail view."
    End If
    If registered Then
        profilePageViews.Remove profilePageViews.Count
        profilePageCrops.Remove name
        profilePageRegions.Remove name
        profileGeometryViews.Remove name
        If profileSheetContours.Exists(name) Then
            profileSheetContours.Remove name
        End If
    End If
End Sub

Private Sub MakePagedProfileDrawing(ByVal d As SldWorks.ModelDoc2, ByVal doc As SldWorks.ModelDoc2, ByVal cfg As String, ByVal body As SldWorks.Body2, ByVal trial As SldWorks.View, ByVal plan As Collection, ByVal outer As Collection, ByVal openings As Collection)
    Dim item As Variant, face As SldWorks.Face2, normal As Variant, box As Variant, axis As Long
    profilePagedMode = True
    profileRoleDrawing = True
    profilePageFirstSheet = AsDrawing(d).GetCurrentSheet.GetName
    profilePageModel = doc.GetPathName
    profilePageConfiguration = cfg
    profilePageBody = body.Name
    profilePageOrientation = trial.GetOrientationName
    profilePageAngle = trial.Angle
    profilePageMatrix = trial.ModelToViewTransform.ArrayData
    Set face = ProfilePlane(body)
    normal = face.Normal
    box = body.GetBodyBox
    For axis = SCREEN_X_AXIS_INDEX To SPACE_LAST_AXIS_INDEX
        If Abs(normal(axis)) > FACE_NORMAL_ALIGNMENT_COSINE Then
            profilePageThickness = box(axis + BBOX_MAX_OFFSET) - box(axis)
        End If
    Next axis
    Set profilePagePlan = plan
    Set profileRawLoops = New Collection
    profileRawLoops.Add outer
    For Each item In openings
        profileRawLoops.Add item
    Next item
    Set profilePageInitialView = trial
    profilePageLimit = PROFILE_PAGE_INITIAL_REQUESTS
    profilePageSplitSerial = 0
    Log "PROFILE PAGE configured; reusing complete planning view"
    If Not profilePanelFrame Is Nothing Then
        profileCompactMode = True
        BuildCompactProfilePages d
    Else
        BuildProfilePages d
    End If
End Sub

Private Sub ClearProfilePages(ByVal d As SldWorks.ModelDoc2)
    Dim item As Variant
    For Each item In profilePageViews
        AsDrawing(d).ActivateSheet CStr(item(0))
        AsDrawing(d).ActivateView ""
        d.ClearSelection2 True
        If d.Extension.SelectByID2(CStr(item(1)), "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0) Then
            If d.Extension.DeleteSelection2(PROFILE_PAGE_DELETE_OPTIONS) = False Then
                Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot remove a rejected generated view."
            End If
        Else
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot select a rejected generated view."
        End If
    Next item
    For Each item In profilePageSheets
        d.ClearSelection2 True
        If d.Extension.SelectByID2(CStr(item), "SHEET", 0, 0, 0, False, 0, Nothing, 0) = False Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot select a rejected generated sheet."
        End If
        If d.Extension.DeleteSelection2(PROFILE_PAGE_DELETE_OPTIONS) = False Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot remove a rejected generated sheet."
        End If
    Next item
    AsDrawing(d).ActivateSheet profilePageFirstSheet
End Sub

Private Function SplitProfilePageGroup(ByVal viewName As String) As Boolean
    Dim group As Variant, request As Variant, count As Long, half As Long, index As Long
    For Each group In profilePageGroups
        If group.Exists("view") Then
          If group("view") = viewName Then
            count = group("requests").Count
            If count <= 1 Then
                Set request = group("requests")(1)
                If Not request.Exists("pagePadding") Then
                    request.Add "pagePadding", 0
                End If
                If request("pagePadding") >= PROFILE_PAGE_SINGLE_RETRIES Then
                    Exit Function
                End If
                request("pagePadding") = request("pagePadding") + 1
                Log "PROFILE PAGE increasing single-dimension clearance; retry=" & request("pagePadding")
                SplitProfilePageGroup = True
                Exit Function
            End If
            half = -Int(-count / 2)
            profilePageSplitSerial = profilePageSplitSerial + 1
            For Each request In group("requests")
                request("pagePartition") = CStr(profilePageSplitSerial) & "|" & CStr(index \ half)
                index = index + 1
            Next request
            Log "PROFILE PAGE splitting crowded group=" & group("number") & "; requests=" & count
            SplitProfilePageGroup = True
            Exit Function
          End If
        End If
    Next group
End Function

Private Sub FinishProfilePages(ByVal d As SldWorks.ModelDoc2)
    Dim number As Long, message As String, group As Variant, cell As Variant
    If profileFamilyMode Then
        FinishFamilyProfilePages d
        Exit Sub
    End If
    If profileCompactMode Then
        FinishCompactProfilePages d
        Exit Sub
    End If
    Do
        On Error Resume Next
        FinishPagedDimensionLayout d
        number = Err.Number
        message = Err.Description
        Err.Clear
        On Error GoTo 0
        If number = 0 Then
            Exit Do
        End If
        If number <> vbObjectError + ERR_DIMENSION_LAYOUT Then
            Err.Raise number, "FinishProfilePages", message
        End If
        If Not profilePageCrops.Exists(profilePageFailureView) Then
            Err.Raise number, "FinishProfilePages", "Complete overview: " & message
        End If
        If Not SplitProfilePageGroup(profilePageFailureView) Then
            Err.Raise number, "FinishProfilePages", message
        End If
        Log "PROFILE PAGE overflow; repacking after local split: " & message
        ClearProfilePages d
        BuildProfilePages d
    Loop
    For Each group In profilePageGroups
        AsDrawing(d).ActivateSheet CStr(group("sheet"))
        cell = group("cell")
        AddText d, "Б" & group("number") & " (" & ScaleText(group("scale")) & ")", cell(0), cell(3) - PROFILE_PAGE_CAPTION_GAP_M, TEXT_HEIGHT_SMALL_M, 0
    Next group
    AuditProfilePages d
    AsDrawing(d).ActivateSheet profilePageFirstSheet
    AddText d, "Товщина " & CStr(Round(profilePageThickness * MILLIMETRES_PER_METRE, LINEAR_DECIMAL_PLACES)) & " мм", CONTENT_LEFT_M, PROFILE_PAGE_THICKNESS_NOTE_Y_M, TEXT_HEIGHT_STANDARD_M, 0
    Log "PROFILE PAGE COMPLETE sheets=" & AsDrawing(d).GetSheetCount & "; requests=" & profilePagePlan.Count
End Sub

' Keep native aligned text above its dimension line. The legacy free-label
' grid can take minutes and disconnect text from a dense panel. Here a failed
' bounded view is split into more enlarged views instead of searching forever.
Private Sub FinishPagedDimensionLayout(ByVal d As SldWorks.ModelDoc2, Optional ByVal selectedView As String = "")
    Dim name As Variant, item As Variant, other As Variant, labelItem As Variant, sheetViews As Collection, records As Collection, bases As Collection
    Dim view As SldWorks.View, display As SldWorks.DisplayDimension, request As Object, region As Variant, cellRegion As Variant, box As Variant, bounds As Variant
    Dim labels As Collection, a As Variant, b As Variant, side As Long, lane As Long, kind As String
    For Each name In AsDrawing(d).GetSheetNames
        AsDrawing(d).ActivateSheet CStr(name)
        AsDrawing(d).ActivateView ""
        Set sheetViews = New Collection
        Set view = AsDrawing(d).GetFirstView.GetNextView
        Do While Not view Is Nothing
            sheetViews.Add view
            Set view = view.GetNextView
        Loop
        CacheProfileContours sheetViews
        Set bases = New Collection
        For Each item In sheetViews
            Set view = item
            If profilePageRegions.Exists(view.Name) And (selectedView = "" Or selectedView = view.Name) Then
                profilePageFailureView = view.Name
                region = profilePageRegions(view.Name)
                cellRegion = region
                region = Array(region(0) + LEADER_BASE_SAFETY_MARGIN_M, region(1) + LEADER_BASE_SAFETY_MARGIN_M, region(2) - LEADER_BASE_SAFETY_MARGIN_M, region(3) - LEADER_BASE_SAFETY_MARGIN_M)
                Set records = New Collection
                Set display = view.GetFirstDisplayDimension5
                Do While Not display Is Nothing
                    Set request = profileDimensionRequests(display.GetNameForSelection)
                    display.OffsetText = False
                    display.CenterText = True
                    If request("role") = PROFILE_ROLE_OUTER And request.Exists("panel") Then
                        kind = CStr(request("kind"))
                        If kind = "overall-width" Or kind = "major-cut-width" Or kind = "major-cut-location" Then
                            featureDimensionPlacement(display.GetNameForSelection) = NATIVE_LABEL_SIDE_BELOW
                            lane = 0
                            If kind = "overall-width" Then
                                lane = PROFILE_PAGE_OVERALL_LANE
                            ElseIf kind = "major-cut-location" Then
                                lane = PROFILE_PAGE_LOCATION_LANE
                            End If
                            featureDimensionLanes(display.GetNameForSelection) = lane
                        End If
                    End If
                    If Not ProfilePagePlaceOverview(d, display, view) Then
                        PlaceProfileDimensionBase d, display, view, region, bases
                    End If
                    ' Moving an annotation can clear native text centering.
                    display.CenterText = True
                    records.Add display
                    Set display = display.GetNext5
                Loop
                d.GraphicsRedraw2
                If Not profilePageCrops.Exists(view.Name) And records.Count > 0 Then
                    If ProfilePageFitOverview(d, view, records, cellRegion) Then
                        CacheProfileContours sheetViews
                    End If
                End If
                For Each other In records
                    Set display = other
                    bounds = NativeDimensionBounds(display)
                    If Not NativeBoundsInside(bounds, cellRegion) Then
                        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Profile dimension graphics exceed the reserved view cell: " & display.GetNameForSelection & "; bounds=" & CStr(bounds(0)) & "/" & CStr(bounds(1)) & "/" & CStr(bounds(2)) & "/" & CStr(bounds(3)) & "; cell=" & CStr(cellRegion(0)) & "/" & CStr(cellRegion(1)) & "/" & CStr(cellRegion(2)) & "/" & CStr(cellRegion(3))
                    End If
                    If ProfileDimensionCutsMaterial(display, True) Then
                        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Profile dimension graphics cross visible material: " & display.GetNameForSelection
                    End If
                    If PanelOwnsDimension(display) Then
                        If PanelLabelTouchesMaterial(display, NativeTextBounds(display.GetDisplayData, display)) Then
                            box = NativeTextBounds(display.GetDisplayData, display)
                            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Profile dimension text touches visible material: " & display.GetNameForSelection & "; text bounds=" & CStr(box(0)) & "/" & CStr(box(1)) & "/" & CStr(box(2)) & "/" & CStr(box(3))
                        End If
                    End If
                    Set labels = New Collection
                    For Each labelItem In records
                        If labelItem.GetNameForSelection <> display.GetNameForSelection Then
                            box = NativeTextBounds(labelItem.GetDisplayData, labelItem)
                            If NativeBoxesOverlap(box, NativeTextBounds(display.GetDisplayData, display), LEADER_LABEL_CLEARANCE_M) Then
                                Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Profile dimension labels overlap: " & display.GetNameForSelection
                            End If
                            labels.Add box
                        End If
                    Next labelItem
                    If NativeLinesHitLabels(display, labels, True) Then
                        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Profile dimension lines cross another label: " & display.GetNameForSelection
                    End If
                Next other
                Log "PROFILE PAGE fitted view=" & view.Name & "; dimensions=" & records.Count
            End If
        Next item
    Next name
End Sub

Private Function ProfilePageFitOverview(ByVal d As SldWorks.ModelDoc2, ByVal view As SldWorks.View, ByVal records As Collection, ByVal region As Variant) As Boolean
    Dim bounds As Variant, item As Variant, axis As Long, shift As Variant, position As Variant, request As Object
    bounds = view.GetOutline
    For Each item In records
        bounds = ProfilePageUnion(bounds, NativeDimensionBounds(item))
    Next item
    shift = Array(0#, 0#)
    For axis = SCREEN_X_AXIS_INDEX To SCREEN_Y_AXIS_INDEX
        If bounds(axis + BBOX_MAX_2D_OFFSET) - bounds(axis) > region(axis + BBOX_MAX_2D_OFFSET) - region(axis) Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "The complete overview and its dimensions exceed the reserved region."
        End If
        shift(axis) = -MinD(0, bounds(axis) - region(axis)) + MinD(0, region(axis + BBOX_MAX_2D_OFFSET) - bounds(axis + BBOX_MAX_2D_OFFSET))
    Next axis
    If Abs(shift(0)) + Abs(shift(1)) <= PROFILE_MATERIAL_TOLERANCE_M Then
        Exit Function
    End If
    position = view.Position
    position(0) = position(0) + shift(0)
    position(1) = position(1) + shift(1)
    view.Position = position
    d.ForceRebuild3 False
    For Each item In records
        Set request = profileDimensionRequests(item.GetNameForSelection)
        request("sheetFirst") = PatternSheetPoint(view, request("first"))
        request("sheetLast") = PatternSheetPoint(view, request("last"))
        If request.Exists("materialBounds") Then
            request.Remove "materialBounds"
        End If
    Next item
    ProfilePageFitOverview = True
End Function

Private Function ProfilePagePlaceOverview(ByVal d As SldWorks.ModelDoc2, ByVal display As SldWorks.DisplayDimension, ByVal view As SldWorks.View) As Boolean
    Dim request As Object, item As Variant, outline As Variant, a As Variant, b As Variant, loops As Variant, axis As Long
    Dim coordinate As Double, x As Double, y As Double, kind As String, width As Double, attempt As Long
    Dim annotation As SldWorks.Annotation, data As SldWorks.DisplayData, arrow As Variant, position As Variant, textBounds As Variant
    Dim textFormat As SldWorks.TextFormat
    Set request = profileDimensionRequests(display.GetNameForSelection)
    If request("role") <> PROFILE_ROLE_OUTER Or Not request.Exists("panel") Then
        Exit Function
    End If
    kind = request("kind")
    outline = view.GetOutline
    loops = profileSheetContours(view.Name)
    a = PatternSheetPoint(view, request("first"))
    b = PatternSheetPoint(view, request("last"))
    axis = request("axis")
    x = (a(0) + b(0)) / 2
    y = (a(1) + b(1)) / 2
    If kind = "overall-width" Then
        coordinate = outline(1) - GOST_OBJECT_OFFSET_M - PROFILE_PAGE_OVERALL_LANE * COORDINATE_LANE_PITCH_M
    ElseIf kind = "major-cut-width" Then
        coordinate = outline(1) - GOST_OBJECT_OFFSET_M
    ElseIf kind = "major-cut-location" Then
        coordinate = outline(1) - GOST_OBJECT_OFFSET_M - PROFILE_PAGE_LOCATION_LANE * COORDINATE_LANE_PITCH_M
    ElseIf kind = "corner-width" Then
        coordinate = outline(3) + GOST_OBJECT_OFFSET_M
    ElseIf kind = "overall-height" Then
        coordinate = outline(0) - GOST_OBJECT_OFFSET_M - PROFILE_PAGE_LOCATION_LANE * COORDINATE_LANE_PITCH_M
    ElseIf kind = "corner-height" Then
        coordinate = outline(0) - GOST_OBJECT_OFFSET_M
        If x > (outline(0) + outline(2)) / 2 Then
            coordinate = outline(2) + GOST_OBJECT_OFFSET_M
        End If
    ElseIf kind = "major-cut-depth" Then
        For Each item In profilePagePlan
            If item("kind") = "major-cut-width" Then
                width = item("value") * view.ScaleDecimal / 2
                Exit For
            End If
        Next item
        coordinate = x + width
        If Not ProfileBaseClear(a, b, axis, coordinate, loops) Then
            coordinate = x - width
        End If
    Else
        Exit Function
    End If
    If axis = SCREEN_X_AXIS_INDEX Then
        y = coordinate
    Else
        x = coordinate
    End If
    request("view") = view.Name
    request("localBase") = kind = "major-cut-depth"
    request("externalWitness") = False
    request("sheetFirst") = a
    request("sheetLast") = b
    display.OffsetText = False
    display.CenterText = True
    Set data = display.GetDisplayData
    If data.GetArrowHeadCount = 2 Then
        If PlanarArrowSpanFits(data.GetArrowHeadAtIndex2(0), data.GetArrowHeadAtIndex2(1)) Then
            display.ArrowSide = SW_DIMENSION_ARROWS_INSIDE
        Else
            display.ArrowSide = SW_DIMENSION_ARROWS_OUTSIDE
        End If
    End If
    Set annotation = display.GetAnnotation
    If CBool(request("localBase")) Then
        textBounds = NativeTextBounds(display.GetDisplayData, display)
        If textBounds(axis + BBOX_MAX_2D_OFFSET) - textBounds(axis) + 2 * LEADER_GEOMETRY_CLEARANCE_M > Abs(a(axis) - b(axis)) Then
            Set textFormat = annotation.GetTextFormat(SW_ANNOTATION_TEXT_FORMAT_FIRST)
            textFormat.CharHeight = TEXT_HEIGHT_SMALL_M
            annotation.SetTextFormat SW_ANNOTATION_TEXT_FORMAT_FIRST, False, textFormat
        End If
    End If
    annotation.SetPosition2 x, y, 0
    display.CenterText = True
    For attempt = 1 To LEADER_CLAMP_ATTEMPTS
        d.GraphicsRedraw2
        Set data = display.GetDisplayData
        arrow = data.GetArrowHeadAtIndex2(0)
        If Abs(arrow(1 - axis) - coordinate) <= PROFILE_MATERIAL_TOLERANCE_M Then
            Exit For
        End If
        position = annotation.GetPosition
        position(1 - axis) = position(1 - axis) + coordinate - arrow(1 - axis)
        annotation.SetPosition2 position(0), position(1), position(2)
        display.CenterText = True
    Next attempt
    ProfilePagePlaceOverview = True
End Function

Private Sub AuditProfilePages(ByVal d As SldWorks.ModelDoc2)
    Dim audit As Object, expected As Object, retained As Object, request As Variant, key As Variant
    Dim sheetName As Variant, view As SldWorks.View, display As SldWorks.DisplayDimension, bounds As Variant, kind As String
    Set expected = CreateObject("Scripting.Dictionary")
    Set retained = CreateObject("Scripting.Dictionary")
    For Each request In profilePagePlan
        If expected.Exists(request("key")) Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "A planned profile measurement is duplicated."
        End If
        expected.Add request("key"), True
    Next request
    Set audit = fs.CreateTextFile(fs.BuildPath(runDir, Format$(jobNumber, JOB_NUMBER_FORMAT) & "_" & PROFILE_PAGE_AUDIT_FILENAME), True, True)
    audit.WriteLine "JOB" & vbTab & jobNumber & vbTab & profilePagePlan.Count
    For Each sheetName In AsDrawing(d).GetSheetNames
        AsDrawing(d).ActivateSheet CStr(sheetName)
        Set view = AsDrawing(d).GetFirstView.GetNextView
        Do While Not view Is Nothing
            If profilePageRegions.Exists(view.Name) Then
                bounds = view.GetOutline
                kind = "full"
                If profilePageCrops.Exists(view.Name) Then
                    kind = "crop"
                End If
                audit.WriteLine "VIEW" & vbTab & jobNumber & vbTab & sheetName & vbTab & view.Name & vbTab & kind & vbTab & view.ScaleDecimal & vbTab & VisibleEdgeCount(view) & vbTab & bounds(0) & vbTab & bounds(1) & vbTab & bounds(2) & vbTab & bounds(3)
                If kind = "crop" Then
                    bounds = profilePageCrops(view.Name)
                    audit.WriteLine "MASK" & vbTab & jobNumber & vbTab & sheetName & vbTab & view.Name & vbTab & bounds(0) & vbTab & bounds(1) & vbTab & bounds(2) & vbTab & bounds(3)
                End If
                Set display = view.GetFirstDisplayDimension5
                Do While Not display Is Nothing
                    key = display.GetNameForSelection
                    If Not profileDimensionRequests.Exists(key) Then
                        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "A profile dimension has no planned owner."
                    End If
                    Set request = profileDimensionRequests(key)
                    If retained.Exists(request("key")) Then
                        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "A profile measurement is repeated across views."
                    End If
                    retained.Add request("key"), True
                    If Abs(display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue - request("value")) > PROFILE_COORDINATE_TOLERANCE_M Then
                        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "A retained profile value differs from its model geometry."
                    End If
                    audit.WriteLine "DIM" & vbTab & jobNumber & vbTab & sheetName & vbTab & view.Name & vbTab & request("key") & vbTab & request("role") & vbTab & request("value") & vbTab & key
                    Set display = display.GetNext5
                Loop
            End If
            Set view = view.GetNextView
        Loop
    Next sheetName
    audit.Close
    If retained.Count <> expected.Count Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Profile page coverage is incomplete."
    End If
End Sub

' Clip the cached material silhouette to the native cropped viewport.
' Dimension placement must never see invisible material outside that viewport.
Private Function ProfileClipPolygon(ByVal polygon As Variant, ByVal box As Variant) As Variant
    Dim points As Collection, output As Collection, a As Variant, b As Variant, item As Variant
    Dim edge As Long, axis As Long, limit As Double, side As Double, insideA As Boolean, insideB As Boolean, fraction As Double
    Dim result() As Variant, index As Long
    Set points = New Collection
    For Each item In polygon
        points.Add item
    Next item
    For edge = 0 To 3
        axis = edge Mod 2
        limit = box(edge)
        side = 1
        If edge >= 2 Then
            side = -1
        End If
        Set output = New Collection
        If points.Count > 0 Then
            a = points(points.Count)
            For Each b In points
                insideA = (a(axis) - limit) * side >= 0
                insideB = (b(axis) - limit) * side >= 0
                If insideA <> insideB Then
                    fraction = (limit - a(axis)) / (b(axis) - a(axis))
                    output.Add Array(a(0) + fraction * (b(0) - a(0)), a(1) + fraction * (b(1) - a(1)))
                End If
                If insideB Then
                    output.Add b
                End If
                a = b
            Next b
        End If
        Set points = output
    Next edge
    If points.Count = 0 Then
        ProfileClipPolygon = Array()
        Exit Function
    End If
    ReDim result(points.Count - 1)
    For index = 1 To points.Count
        result(index - 1) = points(index)
    Next index
    ProfileClipPolygon = result
End Function
