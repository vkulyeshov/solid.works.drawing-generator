Attribute VB_Name = "native_layout"
Option Explicit
Private Function NativeSameBox(ByVal first As Variant, ByVal last As Variant) As Boolean
    Dim index As Long
    For index = OUTLINE_LEFT_INDEX To OUTLINE_TOP_INDEX
        If Abs(first(index) - last(index)) > NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
            Exit Function
        End If
    Next index
    NativeSameBox = True
End Function

Private Sub ReserveLowerViewSpace(ByVal d As SldWorks.ModelDoc2, ByVal views As Collection, ByVal work As Variant)
    Dim item As Variant, v As SldWorks.View, bounds As Variant, position As Variant
    Dim highest As Double, lowest As Double, highestTop As Double, center As Double, shift As Double, count As Long
    highest = -SEARCH_INITIAL_LIMIT
    lowest = SEARCH_INITIAL_LIMIT
    For Each item In views
        Set v = item
        If Not v.GetFirstDisplayDimension5 Is Nothing Then
            bounds = v.GetOutline
            center = (bounds(1) + bounds(3)) / 2
            highest = -MinD(-highest, -center)
            lowest = MinD(lowest, center)
            count = count + 1
        End If
    Next item
    If count < 3 Or highest - lowest < NATIVE_VIEW_ROW_TOLERANCE_M Then
        Exit Sub
    End If
    For Each item In views
        Set v = item
        If Not v.GetFirstDisplayDimension5 Is Nothing Then
            bounds = v.GetOutline
            If (bounds(1) + bounds(3)) / 2 >= highest - NATIVE_VIEW_ROW_TOLERANCE_M Then
                highestTop = -MinD(-highestTop, -bounds(3))
            End If
        End If
    Next item
    shift = MinD(NATIVE_MAX_UPPER_ROW_SHIFT_M, work(3) - NATIVE_UPPER_ROW_TOP_CLEARANCE_M - highestTop)
    If shift <= NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
        Exit Sub
    End If
    For Each item In views
        Set v = item
        If Not v.GetFirstDisplayDimension5 Is Nothing Then
            bounds = v.GetOutline
            If (bounds(1) + bounds(3)) / 2 >= highest - NATIVE_VIEW_ROW_TOLERANCE_M Then
                position = v.Position
                position(1) = position(1) + shift
                v.Position = position
            End If
        End If
    Next item
    d.ForceRebuild3 False
    d.GraphicsRedraw2
    Log "NATIVE LAYOUT upper row shifted by " & shift * MILLIMETRES_PER_METRE & " mm to reserve lower-view dimension space"
End Sub

Private Sub ValidateNativeViewIsolation(ByVal display As SldWorks.DisplayDimension, ByVal owner As String, ByVal views As Collection)
    Dim item As Variant, v As SldWorks.View, bounds As Variant, data As SldWorks.DisplayData, line As Variant, index As Long
    Set data = display.GetDisplayData
    For Each item In views
        Set v = item
        If v.Name <> owner Then
            bounds = v.GetOutline
            If NativeBoxesOverlap(NativeTextBounds(data), bounds, 0) Then
                Log "NATIVE OVERLAP owner=" & owner & "; other=" & v.Name & "; text=" & Join(NativeTextBounds(data), ",") & "; other bounds=" & Join(bounds, ",")
                Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Dimension text overlaps another view: " & display.GetNameForSelection
            End If
            For index = 0 To data.GetLineCount - 1
                line = data.GetLineAtIndex2(index)
                If NativeLeaderCrossesBox(line(DISPLAY_LINE_START_X), line(DISPLAY_LINE_START_Y), line(DISPLAY_LINE_END_X), line(DISPLAY_LINE_END_Y), bounds) Then
                    Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Dimension graphics cross another view: " & display.GetNameForSelection & " / " & v.Name
                End If
            Next index
        End If
    Next item
End Sub

Private Sub ReserveIsometricSpace(ByVal d As SldWorks.ModelDoc2, ByVal views As Collection, ByVal work As Variant)
    Dim item As Variant, view As SldWorks.View, matrix As Variant, bounds As Variant, position As Variant
    Dim axis As Long, orthographic As Boolean, shift As Double
    For Each item In views
        Set view = item
        If view.GetFirstDisplayDimension5 Is Nothing Then
            matrix = view.ModelToViewTransform.ArrayData
            orthographic = False
            For axis = MATRIX_XZ_INDEX To MATRIX_ZZ_INDEX Step MATRIX_ROW_STRIDE
                If Abs(matrix(axis)) >= 1 - DIRECTION_ZERO_TOLERANCE Then
                    orthographic = True
                End If
            Next axis
            If Not orthographic Then
                bounds = view.GetOutline
                shift = work(OUTLINE_RIGHT_INDEX) - LEADER_GEOMETRY_CLEARANCE_M - bounds(OUTLINE_RIGHT_INDEX)
                If shift > DIMENSION_GEOMETRY_GAP_M Then
                    position = view.Position
                    position(SCREEN_X_AXIS_INDEX) = position(SCREEN_X_AXIS_INDEX) + shift
                    view.Position = position
                    Log "NATIVE LAYOUT isometric shifted right by " & shift * MILLIMETRES_PER_METRE & " mm to release dimension space"
                End If
            End If
        End If
    Next item
    d.ForceRebuild3 False
    d.GraphicsRedraw2
End Sub

Private Function NativeViewRegion(ByVal own As SldWorks.View, ByVal views As Collection, ByVal work As Variant) As Variant
    Dim result As Variant, a As Variant, b As Variant, item As Variant, other As SldWorks.View, gapX As Double, gapY As Double
    result = Array(work(0), work(1), work(2), work(3))
    a = own.GetOutline
    For Each item In views
        Set other = item
        If other.Name <> own.Name Then
            If other.GetFirstDisplayDimension5 Is Nothing Then
                GoTo NextRegionView
            End If
            b = other.GetOutline
            gapX = -MinD(-(b(0) - a(2)), -(a(0) - b(2)))
            gapY = -MinD(-(b(1) - a(3)), -(a(1) - b(3)))
            If gapX > 0 Or gapY > 0 Then
                If gapX > gapY Then
                    If b(0) > a(2) Then
                        result(2) = MinD(result(2), (a(2) + b(0) - NATIVE_VIEW_REGION_GAP_M) / 2)
                    Else
                        result(0) = -MinD(-result(0), -(a(0) + b(2) + NATIVE_VIEW_REGION_GAP_M) / 2)
                    End If
                Else
                    If b(1) > a(3) Then
                        result(3) = MinD(result(3), (a(3) + b(1) - NATIVE_VIEW_REGION_GAP_M) / 2)
                    Else
                        result(1) = -MinD(-result(1), -(a(1) + b(3) + NATIVE_VIEW_REGION_GAP_M) / 2)
                    End If
                End If
            Else
                Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Drawing view geometry overlaps: " & own.Name & " / " & other.Name
            End If
        End If
NextRegionView:
    Next item
    If Not NativeBoundsInside(a, result) Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "A view does not fit its isolated dimension region: " & own.Name
    End If
    NativeViewRegion = result
End Function

Private Sub PrepareNativeCalloutText(ByVal d As SldWorks.ModelDoc2)
    Dim v As SldWorks.View, display As SldWorks.DisplayDimension, data As SldWorks.DisplayData
    Dim ann As SldWorks.Annotation, textFormat As SldWorks.TextFormat, first As Variant, point As Variant, index As Long
    Set v = AsDrawing(d).GetFirstView.GetNextView
    Do While Not v Is Nothing
        Set display = v.GetFirstDisplayDimension5
        Do While Not display Is Nothing
            If display.Type2 = SW_DIMENSION_DIAMETER Then
                Set data = display.GetDisplayData
                If data.GetTextCount > 1 Then
                    first = data.GetTextPositionAtIndex(0)
                    For index = 1 To data.GetTextCount - 1
                        point = data.GetTextPositionAtIndex(index)
                        If Abs(point(1) - first(1)) > TEXT_HEIGHT_STANDARD_M / 2 Then
                            Set ann = display.GetAnnotation
                            Set textFormat = ann.GetTextFormat(SW_ANNOTATION_TEXT_FORMAT_FIRST)
                            textFormat.CharHeight = TEXT_HEIGHT_SMALL_M
                            ann.SetTextFormat SW_ANNOTATION_TEXT_FORMAT_FIRST, False, textFormat
                            Exit For
                        End If
                    Next index
                End If
            End If
            Set display = display.GetNext5
        Loop
        Set v = v.GetNextView
    Loop
    d.GraphicsRedraw2
End Sub

Private Sub AddNativeSymmetryCenterline(ByVal d As SldWorks.ModelDoc2, ByVal v As SldWorks.View, ByVal symmetry As Object)
    Dim center As Variant, flags As Variant, transform As Variant, outline As Variant, x As Double, y As Double, axis As Long
    Dim dx As Double, dy As Double, segment As SldWorks.SketchSegment
    Dim sheet As SldWorks.Sheet, properties As Variant, sketchScale As Double, audit As Object, endpoints As Variant
    If v.GetFirstDisplayDimension5 Is Nothing Then
        Exit Sub
    End If
    center = symmetry("center")
    flags = symmetry("axes")
    transform = v.ModelToViewTransform.ArrayData
    outline = v.GetOutline
    x = (center(0) * transform(0) + center(1) * transform(3) + center(2) * transform(6)) * transform(MATRIX_SCALE_INDEX) + transform(9)
    y = (center(0) * transform(1) + center(1) * transform(4) + center(2) * transform(7)) * transform(MATRIX_SCALE_INDEX) + transform(10)
    AsDrawing(d).ActivateView ""
    d.ClearSelection2 True
    Set sheet = AsDrawing(d).GetCurrentSheet
    properties = sheet.GetProperties
    sketchScale = properties(SHEET_SCALE_NUMERATOR_PROPERTY) / properties(SHEET_SCALE_DENOMINATOR_PROPERTY)
    If sketchScale <= 0 Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Invalid drawing scale for a symmetry centerline."
    End If
    For axis = 0 To 2
        If CBool(flags(axis)) Then
            dx = transform(axis * 3)
            dy = transform(axis * 3 + 1)
            If Abs(dx) > DIRECTION_ZERO_TOLERANCE And Abs(dy) < DIRECTION_ZERO_TOLERANCE Then
                endpoints = Array(x, outline(1) - NATIVE_CENTERLINE_EXTENSION_M, x, outline(3) + NATIVE_CENTERLINE_EXTENSION_M)
                Set segment = d.SketchManager.CreateCenterLine(endpoints(0) / sketchScale, endpoints(1) / sketchScale, 0, endpoints(2) / sketchScale, endpoints(3) / sketchScale, 0)
            ElseIf Abs(dy) > DIRECTION_ZERO_TOLERANCE And Abs(dx) < DIRECTION_ZERO_TOLERANCE Then
                endpoints = Array(outline(0) - NATIVE_CENTERLINE_EXTENSION_M, y, outline(2) + NATIVE_CENTERLINE_EXTENSION_M, y)
                Set segment = d.SketchManager.CreateCenterLine(endpoints(0) / sketchScale, endpoints(1) / sketchScale, 0, endpoints(2) / sketchScale, endpoints(3) / sketchScale, 0)
            Else
                endpoints = Empty
            End If
            If Not IsEmpty(endpoints) Then
                segment.Color = DIMENSION_COLOR_BLACK
                Set audit = fs.OpenTextFile(fs.BuildPath(runDir, NATIVE_CENTERLINES_FILENAME), FILE_FOR_APPENDING, True, FILE_UNICODE)
                audit.WriteLine CStr(jobNumber) & vbTab & v.Name & vbTab & Join(endpoints, vbTab)
                audit.Close
            End If
        End If
    Next axis
    d.ClearSelection2 True
End Sub

Private Sub FinishNativeDimensionLayout(ByVal d As SldWorks.ModelDoc2)
    Dim sheetName As Variant, sheet As SldWorks.Sheet, properties As Variant, region As Variant, outline As Variant
    Dim v As SldWorks.View, display As SldWorks.DisplayDimension, records As Collection, obstacles As Collection, labels As Collection, lines As Collection, item As Variant
    Dim ann As SldWorks.Annotation, before As Object, value As Double, count As Long, baseRegion As Variant
    Dim recordViews As Collection, rows As Collection, index As Long, orderedLabels As Collection
    Dim allViews As Collection, regions As Object, dimensionRegions As Object, symmetry As Object, ownRegion As Variant, regionFile As Object
    Dim dimensionOutlines As Object, dimensionOwners As Object, otherLabels As Collection, other As Variant
    Dim planarLines As Collection
    If SixViewMode() And Not profileRoleDrawing Then
        AssignSixViewDimensions d
    End If
    For Each sheetName In AsDrawing(d).GetSheetNames
        AsDrawing(d).ActivateSheet CStr(sheetName)
        AsDrawing(d).ActivateView ""
        Set sheet = AsDrawing(d).GetCurrentSheet
        properties = sheet.GetProperties
        region = Array(LEADER_WORK_LEFT_M, LEADER_WORK_BOTTOM_M, properties(SHEET_WIDTH_PROPERTY) - LEADER_WORK_EDGE_MARGIN_M, properties(SHEET_HEIGHT_PROPERTY) - LEADER_WORK_EDGE_MARGIN_M)
        Set records = New Collection
        Set recordViews = New Collection
        Set rows = New Collection
        Set planarLines = New Collection
        Set obstacles = New Collection
        Set labels = New Collection
        Set lines = New Collection
        Set before = CreateObject("Scripting.Dictionary")
        Set allViews = New Collection
        Set regions = CreateObject("Scripting.Dictionary")
        Set dimensionRegions = CreateObject("Scripting.Dictionary")
        Set dimensionOutlines = CreateObject("Scripting.Dictionary")
        Set dimensionOwners = CreateObject("Scripting.Dictionary")
        Set v = AsDrawing(d).GetFirstView.GetNextView
        Do While Not v Is Nothing
            If Not planarDimensionRequests Is Nothing Then
                v.DisableAutoUpdate = False
                v.UpdateViewDisplayGeometry
            End If
            allViews.Add v
            Set v = v.GetNextView
        Loop
        If allViews.Count = 0 Then
            Exit Sub
        End If
        If profileRoleDrawing Then
            ReserveProfileDimensionSpace d, allViews, region
        Else
            ReserveLowerViewSpace d, allViews, region
            ReserveIsometricSpace d, allViews, region
        End If
        CacheProfileContours allViews
        Set v = allViews(1)
        Set symmetry = NativeSymmetryEvidence(v)
        If Not SixViewMode() And Not profileRoleDrawing Then
            DeduplicateNativeDimensions d, symmetry
        End If
        PrepareNativeCalloutText d
        For Each item In allViews
            Set v = item
            ownRegion = NativeViewRegion(v, allViews, region)
            regions.Add v.Name, ownRegion
            Set regionFile = fs.OpenTextFile(fs.BuildPath(runDir, NATIVE_REGIONS_FILENAME), FILE_FOR_APPENDING, True, FILE_UNICODE)
            regionFile.WriteLine CStr(jobNumber) & vbTab & v.Name & vbTab & Join(ownRegion, vbTab)
            regionFile.Close
        Next item
        If Not planarDimensionRequests Is Nothing Then
            FinishPlanarDimensionLayout d, allViews, regions, region
            Exit Sub
        End If
        obstacles.Add Array(region(0), properties(SHEET_HEIGHT_PROPERTY) - LEADER_STAMP_TOP_MARGIN_M, LEADER_STAMP_WIDTH_M, properties(SHEET_HEIGHT_PROPERTY))
        Set v = AsDrawing(d).GetFirstView.GetNextView
        d.ClearSelection2 True
        Do While Not v Is Nothing
            obstacles.Add v.GetOutline
            Set display = v.GetFirstDisplayDimension5
            Do While Not display Is Nothing
                records.Add display
                recordViews.Add v
                dimensionRegions.Add display.GetNameForSelection, regions(v.Name)
                dimensionOutlines.Add display.GetNameForSelection, v.GetOutline
                dimensionOwners.Add display.GetNameForSelection, v.Name
                before.Add display.GetNameForSelection, display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue
                Set ann = display.GetAnnotation
                ann.Select3 True, Nothing
                Set display = display.GetNext5
            Loop
            Set v = v.GetNextView
        Loop
        If records.Count > 1 Then
            If Not d.Extension.AlignDimensions(SW_AUTO_ARRANGE_DIMENSIONS, 0) Then
                Log "NATIVE Auto Arrange reported no update; checking the existing dimension graphics"
            End If
        End If
        ' Native chain alignment must not move other bases while labels are routed.
        d.ClearSelection2 True
        For Each item In records
            Set display = item
            Set ann = display.GetAnnotation
            ann.Select3 True, Nothing
        Next item
        d.BreakDimensionAlignment
        d.ClearSelection2 True
        For index = 1 To records.Count
            Set display = records(index)
            Set v = recordViews(index)
            ownRegion = regions(v.Name)
            baseRegion = Array(ownRegion(0) + LEADER_BASE_SAFETY_MARGIN_M, ownRegion(1) + LEADER_BASE_SAFETY_MARGIN_M, ownRegion(2) - LEADER_BASE_SAFETY_MARGIN_M, ownRegion(3) - LEADER_BASE_SAFETY_MARGIN_M)
            If PlanarOwnsDimension(display) Then
                If planarLines Is Nothing Then
                    Set planarLines = New Collection
                End If
                PlacePlanarDimensionBase d, display, v, baseRegion, planarLines, rows
            ElseIf ProfileOwnsDimension(display) Then
                PlaceProfileDimensionBase d, display, v
            Else
                PlaceNativeDimensionBase display, v, baseRegion, obstacles, rows
                ClampNativeDimension display, baseRegion
            End If
            If NativePlacementSide(display) = NATIVE_LABEL_SIDE_RIGHT And display.Type2 = SW_DIMENSION_DIAMETER Then
                PlaceDirectedPitchLabel d, display, v, baseRegion
            End If
        Next index
        d.GraphicsRedraw2
        For index = 1 To records.Count
            Set display = records(index)
            Set v = recordViews(index)
            If display.Type2 = SW_DIMENSION_LINEAR Then
                ' Stabilize native bases before reserving their tails and witnesses.
                If PlanarOwnsDimension(display) Then
                    StabilizePlanarDimensionText d, display, v
                ElseIf NativePlacementSide(display) = NATIVE_LABEL_SIDE_BELOW And Not ProfileLocalBase(display) Then
                    PlaceCenteredBelowDimension d, display, v
                Else
                    display.CenterText = False
                    display.OffsetText = True
                End If
            End If
        Next index
        d.GraphicsRedraw2
        For Each item In records
            Set display = item
            NativeCollectLines display, lines, True
            If NativeFixedLabel(display) Then
                If NativeBoxesOverlap(NativeTextBounds(display.GetDisplayData), dimensionOutlines(display.GetNameForSelection), LEADER_GEOMETRY_CLEARANCE_M) And Not (PlanarOwnsDimension(display) And display.Type2 = SW_DIMENSION_ANGULAR) Then
                    Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Directed dimension text overlaps its own view."
                End If
                labels.Add NativeTextBounds(display.GetDisplayData)
            End If
        Next item
        Set orderedLabels = NativeLabelsByWidth(records)
        For Each item In orderedLabels
            Set display = item
            Log "NATIVE LABEL " & display.GetNameForSelection
            ownRegion = dimensionRegions(display.GetNameForSelection)
            outline = dimensionOutlines(display.GetNameForSelection)
            ' Side policies guide the first search pass; the isolated view region
            ' remains available when that preferred side cannot hold the text.
            If NativeFixedLabel(display) Then
                ' These directed labels were positioned and reserved before the other labels.
            ElseIf Not KeepReadableNativeText(display, ownRegion, obstacles, labels, lines, outline) Then
                PlaceNativeDimensionLabel d, display, ownRegion, obstacles, labels, lines, dimensionOutlines(display.GetNameForSelection)
            End If
        Next item
        d.GraphicsRedraw2
        For Each item In records
            Set display = item
            value = display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue
            If Abs(value - CDbl(before(display.GetNameForSelection))) > LEADER_VALUE_TOLERANCE_M Then
                Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Dimension value changed during leader placement."
            End If
            ownRegion = dimensionRegions(display.GetNameForSelection)
            If Not NativeBoundsInside(NativeDimensionBounds(display), ownRegion) Then
                outline = NativeDimensionBounds(display)
                Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Dimension graphics outside the drawing area: " & display.GetNameForSelection & "; bounds=" & Join(outline, ",")
            End If
            ValidateNativeViewIsolation display, CStr(dimensionOwners(display.GetNameForSelection)), allViews
            If ProfileDimensionCutsMaterial(display) Then
                Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Dimension graphics cross profile material: " & display.GetNameForSelection
            End If
            Set otherLabels = New Collection
            For Each other In records
                If other.GetNameForSelection <> display.GetNameForSelection Then
                    otherLabels.Add NativeTextBounds(other.GetDisplayData)
                End If
            Next other
            If NativeLinesHitLabels(display, otherLabels, True) Then
                Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Dimension graphics cross another dimension's text: " & display.GetNameForSelection
            End If
        Next item
        count = count + records.Count
        For Each item In allViews
            Set v = item
            AddNativeSymmetryCenterline d, v, symmetry
        Next item
    Next sheetName
    If SixViewMode() And Not profileRoleDrawing Then
        RenderHolePatternCircles d
        AuditFinalSixViews d
    End If
    Log "NATIVE LEADERS dimensions=" & count & "; all graphics inside the drawing area"
End Sub

