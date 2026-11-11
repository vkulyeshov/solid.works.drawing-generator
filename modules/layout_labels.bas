Attribute VB_Name = "layout_labels"
Option Explicit

' The same bounded text pass is used by profile, prismatic and ordinary
' solid-part planners. Ownership and source measurements are never changed.
Private Function LayoutMaterial(ByVal display As SldWorks.DisplayDimension, ByVal view As SldWorks.View) As Variant
    Dim raw() As Variant, index As Long, polygon As Variant, request As Object
    If ProfileOwnsDimension(display) Then
        Set request = profileDimensionRequests(display.GetNameForSelection)
        If profileSheetContours.Exists(view.Name) Then
            LayoutMaterial = profileSheetContours(view.Name)
            Exit Function
        End If
    End If
    If Not planarContourPoints Is Nothing Then
        ReDim raw(planarContourPoints.Count - 1)
        For index = 1 To planarContourPoints.Count
            raw(index - 1) = PatternSheetPoint(view, planarContourPoints(index))
        Next index
        polygon = PlanarClipPolygon(raw, view.GetOutline)
        If Not IsEmpty(polygon) Then
            LayoutMaterial = Array(polygon)
        End If
    End If
End Function

Private Function LayoutLabelScore(ByVal display As SldWorks.DisplayDimension, ByVal view As SldWorks.View, ByVal records As Collection, ByVal region As Variant, ByVal material As Variant) As Long
    Dim bounds As Variant, box As Variant, otherBox As Variant, other As Variant, key As Variant, name As String, cached As Object
    Dim polygons As Collection, lines As Collection, otherPolygons As Collection
    bounds = NativeDimensionBounds(display)
    Set polygons = LayoutInkPolygons(display)
    box = LayoutInkBounds(polygons)
    If Not NativeBoundsInside(bounds, region) Or Not NativeBoundsInside(box, region) Then
        LayoutLabelScore = LayoutLabelScore + COMPACT_OUTSIDE_COST
    End If
    If NativeGraphicsHitObstacles(display, semanticLayoutObstacles, view.GetOutline) Then
        LayoutLabelScore = LayoutLabelScore + COMPACT_OUTSIDE_COST
    End If
    For Each other In semanticLayoutObstacles
        If Not NativeSameBox(other, view.GetOutline) Then
            If NativeBoxesOverlap(box, other, COMPACT_TEXT_CLEARANCE_M) Then
                LayoutLabelScore = LayoutLabelScore + COMPACT_OUTSIDE_COST
            End If
        End If
    Next other
    If IsEmpty(material) Then
        If NativeBoxesOverlap(box, view.GetOutline, COMPACT_TEXT_CLEARANCE_M) Then
            LayoutLabelScore = LayoutLabelScore + COMPACT_OUTSIDE_COST
        End If
    ElseIf PanelBoxTouchesMaterial(box, material) Then
        LayoutLabelScore = LayoutLabelScore + COMPACT_OUTSIDE_COST
    End If
    If LayoutBaselineCutsMaterial(display, view, material) Then
        LayoutLabelScore = LayoutLabelScore + COMPACT_OUTSIDE_COST
    End If
    Set lines = PlanarDisplayLines(display.GetDisplayData)
    If display.Type2 = SW_DIMENSION_LINEAR Or display.Type2 = SW_DIMENSION_ANGULAR Then
        If LayoutOwnLinesHitText(display, lines, polygons) Then
            LayoutLabelScore = LayoutLabelScore + 1
        End If
    ElseIf PlanarLinesHitText(lines, CompactTextPolygons(display)) Then
        LayoutLabelScore = LayoutLabelScore + 1
    End If
    If LayoutHeadsHitText(display.GetDisplayData, polygons) Then
        LayoutLabelScore = LayoutLabelScore + 1
    End If
    name = display.GetNameForSelection
    For Each key In semanticLayoutGeometry.Keys
        If CStr(key) <> name Then
            Set cached = semanticLayoutGeometry(key)
            otherBox = cached("bounds")
            If NativeBoxesOverlap(box, otherBox, COMPACT_TEXT_CLEARANCE_M) Then
                LayoutLabelScore = LayoutLabelScore + COMPACT_TEXT_COLLISION_COST
            End If
            If PlanarLinesHitText(cached("lines"), polygons) Then
                LayoutLabelScore = LayoutLabelScore + 1
            End If
            If LayoutHeadPolygonsHitText(cached("heads"), polygons) Then
                LayoutLabelScore = LayoutLabelScore + 1
            End If
            Set otherPolygons = cached("text")
            If PlanarLinesHitText(lines, otherPolygons) Then
                LayoutLabelScore = LayoutLabelScore + 1
            End If
            ' Planar drawings use the joint solver's DSTU GOST 2.307 crossing rule.
            If Not planarDimensionRequests Is Nothing Then
                If PlanarBasesCrossLines(PlanarDimensionBases(display), cached("lines")) Or PlanarBasesCrossLines(cached("bases"), lines) Then
                    LayoutLabelScore = LayoutLabelScore + 1
                End If
            End If
        End If
    Next key
End Function

Private Sub LayoutCacheDimension(ByVal display As SldWorks.DisplayDimension)
    Dim record As Object, heads As New Collection, index As Long, triangle As Variant, data As SldWorks.DisplayData
    Set record = CreateObject("Scripting.Dictionary")
    Set data = display.GetDisplayData
    record.Add "bounds", LayoutInkBounds(LayoutInkPolygons(display))
    record.Add "text", LayoutInkPolygons(display)
    record.Add "lines", PlanarDisplayLines(data)
    record.Add "bases", PlanarDimensionBases(display)
    For index = 0 To data.GetArrowHeadCount - 1
        triangle = PlanarArrowTriangle(data.GetArrowHeadAtIndex2(index))
        If Not IsEmpty(triangle) Then
            heads.Add triangle
        End If
    Next index
    record.Add "heads", heads
    If semanticLayoutGeometry.Exists(display.GetNameForSelection) Then
        semanticLayoutGeometry.Remove display.GetNameForSelection
    End If
    semanticLayoutGeometry.Add display.GetNameForSelection, record
End Sub

' DisplayData positions describe the character cell, not the complete printed
' ink. Include descenders and export rounding when checking another leader.
Private Function LayoutInkPolygons(ByVal display As SldWorks.DisplayDimension) As Collection
    Set LayoutInkPolygons = PlanarTextPolygons(display, display.GetDisplayData, LAYOUT_TEXT_INK_MARGIN_M, True)
End Function

Private Function LayoutInkBounds(ByVal polygons As Collection) As Variant
    Dim result As Variant, polygon As Variant, point As Variant
    result = Array(SEARCH_INITIAL_LIMIT, SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT)
    For Each polygon In polygons
        For Each point In polygon
            NativeExpandBounds result, point(0), point(1)
        Next point
    Next polygon
    LayoutInkBounds = result
End Function

Private Function LayoutOwnLinesHitText(ByVal display As SldWorks.DisplayDimension, ByVal lines As Collection, ByVal polygons As Collection) As Boolean
    Dim segment As Variant, candidates As New Collection, data As SldWorks.DisplayData
    Dim point As Variant, angle As Double, firstDistance As Double, lastDistance As Double, baseline As Boolean
    Set data = display.GetDisplayData
    point = data.GetTextPositionAtIndex(0)
    angle = data.GetTextAngleAtIndex(0)
    For Each segment In lines
        baseline = False
        firstDistance = -(segment(0) - point(0)) * Sin(angle) + (segment(1) - point(1)) * Cos(angle)
        lastDistance = -(segment(2) - point(0)) * Sin(angle) + (segment(3) - point(1)) * Cos(angle)
        baseline = Abs(firstDistance) < LAYOUT_BASELINE_TOLERANCE_M And Abs(lastDistance) < LAYOUT_BASELINE_TOLERANCE_M
        If Not baseline Then
            candidates.Add segment
        End If
    Next segment
    LayoutOwnLinesHitText = PlanarLinesHitText(candidates, polygons)
End Function

Private Function LayoutHeadPolygonsHitText(ByVal heads As Collection, ByVal polygons As Collection) As Boolean
    Dim triangle As Variant, polygon As Variant
    For Each triangle In heads
        For Each polygon In polygons
            If PlanarPolygonsOverlap(triangle, polygon) Then
                LayoutHeadPolygonsHitText = True
                Exit Function
            End If
        Next polygon
    Next triangle
End Function

Private Function LayoutHeadsHitText(ByVal data As SldWorks.DisplayData, ByVal polygons As Collection) As Boolean
    Dim index As Long, triangle As Variant, polygon As Variant
    For index = 0 To data.GetArrowHeadCount - 1
        triangle = PlanarArrowTriangle(data.GetArrowHeadAtIndex2(index))
        If Not IsEmpty(triangle) Then
            For Each polygon In polygons
                If PlanarPolygonsOverlap(triangle, polygon) Then
                    LayoutHeadsHitText = True
                    Exit Function
                End If
            Next polygon
        End If
    Next index
End Function

Private Function LayoutBaselineCutsMaterial(ByVal display As SldWorks.DisplayDimension, ByVal view As SldWorks.View, ByVal material As Variant) As Boolean
    Dim request As Object, first As Variant, last As Variant, line As Variant
    If IsEmpty(material) Or display.Type2 <> SW_DIMENSION_LINEAR Then
        Exit Function
    End If
    If ProfileOwnsDimension(display) Then
        LayoutBaselineCutsMaterial = ProfileDimensionCutsMaterial(display)
    ElseIf PlanarOwnsDimension(display) Then
        Set request = planarDimensionRequests(display.GetNameForSelection)
        If request("kind") = "edge" Then
            first = PatternSheetPoint(view, request("first"))
            last = PatternSheetPoint(view, request("last"))
            For Each line In PlanarDisplayLines(display.GetDisplayData)
                If ProfileLineCutsMaterial(Array(line(0), line(1)), Array(line(2), line(3)), material) Then
                    If Not PanelReferenceWitness(Array(line(0), line(1)), Array(line(2), line(3)), first, last) Then
                        LayoutBaselineCutsMaterial = True
                        Exit Function
                    End If
                End If
            Next line
        End If
    End If
End Function

Private Function LayoutLinearDirection(ByVal display As SldWorks.DisplayDimension) As Variant
    Dim data As SldWorks.DisplayData, a As Variant, b As Variant, length As Double
    Set data = display.GetDisplayData
    If data.GetArrowHeadCount <> DIMENSION_REFERENCE_SECOND Then
        Exit Function
    End If
    a = data.GetArrowHeadAtIndex2(0)
    b = data.GetArrowHeadAtIndex2(1)
    length = Sqr((b(0) - a(0)) ^ 2 + (b(1) - a(1)) ^ 2)
    If length > DIRECTION_ZERO_TOLERANCE Then
        LayoutLinearDirection = Array((b(0) - a(0)) / length, (b(1) - a(1)) / length)
    End If
End Function

Private Function LayoutRouteLabel(ByVal d As SldWorks.ModelDoc2, ByVal display As SldWorks.DisplayDimension, ByVal view As SldWorks.View, ByVal records As Collection, ByVal region As Variant, ByVal material As Variant) As Boolean
    Dim annotation As SldWorks.Annotation, start As Variant, bestPosition As Variant, direction As Variant, position As Variant, outline As Variant
    Dim best As Long, score As Long, trial As Long, sign As Long, side As Long, row As Long, shift As Double, normalShift As Double
    Dim centered As Boolean, offset As Boolean, value As Double, linear As Boolean, attempts As Long
    Dim textSignature As String, bounds As Variant, adjustment As Variant, fit As Long
    best = LayoutLabelScore(display, view, records, region, material)
    If best = 0 Then
        LayoutRouteLabel = True
        Exit Function
    End If
    If display.Type2 = SW_DIMENSION_ANGULAR And PlanarOwnsDimension(display) Then
        Exit Function
    End If
    If display.Type2 = SW_DIMENSION_DIAMETER Or display.Type2 = SW_DIMENSION_RADIAL Then
        If NativePlacementSide(display) > 0 Then
            ' Directed circle leaders retain the source planner's side.
            Exit Function
        End If
    End If
    linear = display.Type2 = SW_DIMENSION_LINEAR
    Set annotation = display.GetAnnotation
    start = annotation.GetPosition
    bestPosition = start
    value = display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue
    textSignature = LayoutTextSignature(display)
    centered = display.CenterText
    offset = display.OffsetText
    outline = view.GetOutline
    If linear Then
        direction = LayoutLinearDirection(display)
        If IsEmpty(direction) Then
            Exit Function
        End If
        display.OffsetText = False
    Else
        display.SetBrokenLeader2 False, SW_BROKEN_LEADER_HORIZONTAL_TEXT
        If display.Type2 = SW_DIMENSION_DIAMETER Or display.Type2 = SW_DIMENSION_RADIAL Then
            display.SolidLeader = False
            display.SetBentLeaderLength False, LEADER_SHOULDER_LENGTH_M
        ElseIf display.Type2 = SW_DIMENSION_ANGULAR Then
            display.CenterText = False
            display.OffsetText = True
        End If
    End If
    For row = 0 To LAYOUT_NORMAL_RETRIES
        For trial = 1 To LAYOUT_LABEL_TRIALS
            For sign = -1 To 1 Step 2
                attempts = attempts + 1
                If attempts > LAYOUT_LABEL_CANDIDATE_LIMIT Then
                    GoTo RestoreLayoutLabel
                End If
                If linear Then
                    shift = trial * COMPACT_LABEL_STEP_M * sign
                    normalShift = row * COORDINATE_LANE_PITCH_M * sign
                    position = Array(start(0) + direction(0) * shift - direction(1) * normalShift, start(1) + direction(1) * shift + direction(0) * normalShift, start(2))
                    display.CenterText = False
                    annotation.SetPosition2 position(0), position(1), position(2)
                    display.CenterText = False
                Else
                    side = (trial - 1) Mod LAYOUT_LABEL_SIDE_COUNT
                    shift = ((trial - 1) \ LAYOUT_LABEL_SIDE_COUNT + row) * COORDINATE_LANE_PITCH_M * sign
                    position = Array((outline(0) + outline(2)) / 2 + shift, (outline(1) + outline(3)) / 2 + shift, start(2))
                    Select Case side
                        Case 0
                            position(0) = outline(2) + GOST_OBJECT_OFFSET_M
                        Case 1
                            position(0) = outline(0) - GOST_OBJECT_OFFSET_M
                        Case 2
                            position(1) = outline(3) + GOST_OBJECT_OFFSET_M
                        Case Else
                            position(1) = outline(1) - GOST_OBJECT_OFFSET_M
                    End Select
                    annotation.SetPosition2 position(0), position(1), position(2)
                    d.GraphicsRedraw2
                    If display.Type2 = SW_DIMENSION_DIAMETER Or display.Type2 = SW_DIMENSION_RADIAL Then
                        bounds = LayoutInkBounds(LayoutInkPolygons(display))
                        position = annotation.GetPosition
                        Select Case side
                            Case 0
                                position(0) = position(0) + outline(2) + LEADER_GEOMETRY_CLEARANCE_M - bounds(0)
                            Case 1
                                position(0) = position(0) + outline(0) - LEADER_GEOMETRY_CLEARANCE_M - bounds(2)
                            Case 2
                                position(1) = position(1) + outline(3) + LEADER_GEOMETRY_CLEARANCE_M - bounds(1)
                            Case Else
                                position(1) = position(1) + outline(1) - LEADER_GEOMETRY_CLEARANCE_M - bounds(3)
                        End Select
                        annotation.SetPosition2 position(0), position(1), position(2)
                    End If
                End If
                For fit = 1 To LAYOUT_TEXT_CLAMP_ATTEMPTS
                    d.GraphicsRedraw2
                    bounds = LayoutInkBounds(LayoutInkPolygons(display))
                    adjustment = LayoutTextTranslation(bounds, region)
                    If Abs(adjustment(0)) + Abs(adjustment(1)) < PROFILE_MATERIAL_TOLERANCE_M Then
                        Exit For
                    End If
                    position = annotation.GetPosition
                    annotation.SetPosition2 position(0) + adjustment(0), position(1) + adjustment(1), position(2)
                Next fit
                d.GraphicsRedraw2
                If Abs(display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue - value) > LEADER_VALUE_TOLERANCE_M Then
                    Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "A native label move changed its physical measurement."
                End If
                score = LayoutLabelScore(display, view, records, region, material)
                If LayoutTextSignature(display) <> textSignature Then
                    score = score + COMPACT_OUTSIDE_COST
                End If
                If Setting("DimensionLayoutAudit", "0") = "1" And display.IsHoleCallout Then
                    Log "SEMANTIC CALLOUT CANDIDATE " & display.GetNameForSelection & "; side=" & side & "; row=" & row & "; trial=" & trial & "; score=" & score & "; bounds=" & Join(LayoutInkBounds(LayoutInkPolygons(display)), ",") & "; same text=" & (LayoutTextSignature(display) = textSignature)
                End If
                If score < best Then
                    best = score
                    bestPosition = annotation.GetPosition
                End If
                If best = 0 Then
                    LayoutRouteLabel = True
                    Exit Function
                End If
            Next sign
        Next trial
    Next row
RestoreLayoutLabel:
    annotation.SetPosition2 bestPosition(0), bestPosition(1), bestPosition(2)
    If bestPosition(0) = start(0) And bestPosition(1) = start(1) Then
        display.CenterText = centered
        display.OffsetText = offset
    End If
End Function

Private Function LayoutTextTranslation(ByVal bounds As Variant, ByVal region As Variant) As Variant
    Dim shift(1) As Double, axis As Long
    For axis = SCREEN_X_AXIS_INDEX To SCREEN_Y_AXIS_INDEX
        If bounds(axis) < region(axis) + COMPACT_TEXT_CLEARANCE_M Then
            shift(axis) = region(axis) + COMPACT_TEXT_CLEARANCE_M - bounds(axis)
        ElseIf bounds(axis + BBOX_MAX_2D_OFFSET) > region(axis + BBOX_MAX_2D_OFFSET) - COMPACT_TEXT_CLEARANCE_M Then
            shift(axis) = region(axis + BBOX_MAX_2D_OFFSET) - COMPACT_TEXT_CLEARANCE_M - bounds(axis + BBOX_MAX_2D_OFFSET)
        End If
    Next axis
    LayoutTextTranslation = shift
End Function

Private Function LayoutTextSignature(ByVal display As SldWorks.DisplayDimension) As String
    Dim data As SldWorks.DisplayData, index As Long
    Set data = display.GetDisplayData
    For index = 0 To data.GetTextCount - 1
        LayoutTextSignature = LayoutTextSignature & Trim$(data.GetTextAtIndex(index))
    Next index
    LayoutTextSignature = Replace(Replace(Replace(LayoutTextSignature, " ", ""), vbCr, ""), vbLf, "")
End Function

Private Function TrySemanticLabelLayout(ByVal d As SldWorks.ModelDoc2, ByVal views As Collection, ByVal regions As Object, Optional ByVal prepare As Boolean = True, Optional ByVal passes As Long = LAYOUT_LABEL_PASSES) As Boolean
    Dim records As New Collection, owners As New Collection, materials As New Collection, item As Variant
    Dim view As SldWorks.View, display As SldWorks.DisplayDimension, format As SldWorks.TextFormat, direction As Variant
    Dim initial As New Collection, annotation As SldWorks.Annotation, data As SldWorks.DisplayData, a As Variant, b As Variant
    Dim index As Long, pass As Long, unresolved As Long, state As Variant, span As Double, text As Variant, value As Double
    Dim region As Variant
    Dim notes As Variant, note As Variant, extent As Variant
    Set semanticLayoutObstacles = New Collection
    Set semanticLayoutGeometry = CreateObject("Scripting.Dictionary")
    semanticLayoutObstacles.Add Array(LEADER_WORK_LEFT_M, sheetHeight - LEADER_STAMP_TOP_MARGIN_M, LEADER_STAMP_WIDTH_M, sheetHeight)
    For Each item In views
        Set view = item
        semanticLayoutObstacles.Add view.GetOutline
        notes = view.GetNotes
        If Not IsEmpty(notes) Then
            For Each note In notes
                extent = note.GetExtent
                If Not IsEmpty(extent) Then
                    semanticLayoutObstacles.Add Array(extent(BBOX_X_MIN_INDEX), extent(BBOX_Y_MIN_INDEX), extent(BBOX_X_MAX_INDEX), extent(BBOX_Y_MAX_INDEX))
                End If
            Next note
        End If
        Set display = view.GetFirstDisplayDimension5
        Do While Not display Is Nothing
            Set annotation = display.GetAnnotation
            Set format = annotation.GetTextFormat(SW_ANNOTATION_TEXT_FORMAT_FIRST)
            initial.Add Array(annotation.GetPosition, display.CenterText, display.OffsetText, display.ArrowSide, display.GetUseDocBrokenLeader, display.GetBrokenLeader2, display.SolidLeader, display.GetUseDocBentLeaderLength, display.GetBentLeaderLength)
            records.Add display
            owners.Add view
            materials.Add LayoutMaterial(display, view)
            region = regions(view.Name)
            If prepare Then
                text = NativeTextBounds(display.GetDisplayData, display)
                If display.IsHoleCallout Then
                    ' Auto Arrange can restore the document height after the
                    ' earlier callout preparation. Fit it after that native step.
                    Set format = annotation.GetTextFormat(SW_ANNOTATION_TEXT_FORMAT_FIRST)
                    format.CharHeight = TEXT_HEIGHT_SMALL_M
                    annotation.SetTextFormat SW_ANNOTATION_TEXT_FORMAT_FIRST, False, format
                ElseIf display.Type2 <> SW_DIMENSION_LINEAR And text(2) - text(0) > (region(2) - region(0)) / 2 Then
                    Set format = annotation.GetTextFormat(SW_ANNOTATION_TEXT_FORMAT_FIRST)
                    format.CharHeight = TEXT_HEIGHT_SMALL_M
                    annotation.SetTextFormat SW_ANNOTATION_TEXT_FORMAT_FIRST, False, format
                End If
                If display.Type2 = SW_DIMENSION_ANGULAR Then
                    Set data = display.GetDisplayData
                    If data.GetArrowHeadCount = DIMENSION_REFERENCE_SECOND Then
                        a = data.GetArrowHeadAtIndex2(0)
                        b = data.GetArrowHeadAtIndex2(1)
                        text = LayoutInkBounds(LayoutInkPolygons(display))
                        span = Sqr((b(0) - a(0)) ^ 2 + (b(1) - a(1)) ^ 2)
                        If span < text(2) - text(0) + 2 * GOST_ARROW_HEAD_LENGTH_M Then
                            Set format = annotation.GetTextFormat(SW_ANNOTATION_TEXT_FORMAT_FIRST)
                            format.CharHeight = TEXT_HEIGHT_SMALL_M
                            annotation.SetTextFormat SW_ANNOTATION_TEXT_FORMAT_FIRST, False, format
                        End If
                        If PlanarArrowSpanFits(a, b) Then
                            display.ArrowSide = SW_DIMENSION_ARROWS_INSIDE
                        Else
                            display.ArrowSide = SW_DIMENSION_ARROWS_OUTSIDE
                        End If
                    End If
                ElseIf display.Type2 = SW_DIMENSION_LINEAR Then
                    Set data = display.GetDisplayData
                    If data.GetArrowHeadCount = DIMENSION_REFERENCE_SECOND Then
                        a = data.GetArrowHeadAtIndex2(0)
                        b = data.GetArrowHeadAtIndex2(1)
                        span = Sqr((b(0) - a(0)) ^ 2 + (b(1) - a(1)) ^ 2)
                        text = NativeTextBounds(data, display)
                        direction = LayoutLinearDirection(display)
                        If span < Abs(direction(0)) * (text(2) - text(0)) + Abs(direction(1)) * (text(3) - text(1)) + 2 * GOST_ARROW_HEAD_LENGTH_M Then
                            Set format = annotation.GetTextFormat(SW_ANNOTATION_TEXT_FORMAT_FIRST)
                            format.CharHeight = TEXT_HEIGHT_SMALL_M
                            annotation.SetTextFormat SW_ANNOTATION_TEXT_FORMAT_FIRST, False, format
                        End If
                        If span >= 2 * GOST_ARROW_HEAD_LENGTH_M + DIMENSION_GEOMETRY_GAP_M Then
                            display.SetArrowHeadStyle2 False, SW_ARROW_CLOSED, SW_ARROW_CLOSED
                            display.ArrowSide = SW_DIMENSION_ARROWS_INSIDE
                        Else
                            display.ArrowSide = SW_DIMENSION_ARROWS_OUTSIDE
                        End If
                    End If
                End If
            End If
            Set display = display.GetNext5
        Loop
    Next item
    d.ForceRebuild3 False
    d.GraphicsRedraw2
    For Each item In records
        LayoutCacheDimension item
    Next item
    For pass = 1 To passes
        Log "SEMANTIC LABEL PASS " & pass & "/" & passes & "; dimensions=" & records.Count
        For index = 1 To records.Count
            Set display = records(index)
            Set view = owners(index)
            LayoutRouteLabel d, display, view, records, regions(view.Name), materials(index)
            LayoutCacheDimension display
        Next index
    Next pass
    For index = 1 To records.Count
        Set display = records(index)
        Set view = owners(index)
        If LayoutLabelScore(display, view, records, regions(view.Name), materials(index)) <> 0 Then
            unresolved = unresolved + 1
            Log "SEMANTIC UNRESOLVED " & display.GetNameForSelection & "; type=" & display.Type2 & "; score=" & LayoutLabelScore(display, view, records, regions(view.Name), materials(index)) & "; bounds=" & Join(NativeDimensionBounds(display), ",") & "; region=" & Join(regions(view.Name), ",")
        End If
    Next index
    TrySemanticLabelLayout = unresolved = 0
    Log "SEMANTIC LABELS dimensions=" & records.Count & "; unresolved=" & unresolved & "; bounded passes=" & passes
    If unresolved > 0 Then
        For index = 1 To records.Count
            Set display = records(index)
            If planarDimensionRequests Is Nothing Then
                Set view = owners(index)
                If LayoutLabelScore(display, view, records, regions(view.Name), materials(index)) = 0 Then
                    ' Preserve solved labels as the fallback's starting point.
                    GoTo NextLayoutRestore
                End If
            End If
            state = initial(index)
            a = state(0)
            display.SetBrokenLeader2 CBool(state(4)), CLng(state(5))
            display.SolidLeader = CBool(state(6))
            display.SetBentLeaderLength CBool(state(7)), CDbl(state(8))
            Set annotation = display.GetAnnotation
            ' Standard font preparation also applies to the fallback solver.
            display.CenterText = state(1)
            display.OffsetText = state(2)
            display.ArrowSide = state(3)
            annotation.SetPosition2 a(0), a(1), a(2)
NextLayoutRestore:
        Next index
    End If
End Function

' A persisted drawing exposes the font and fragment positions actually used
' by PDF export. Repair labels against that data without Auto Arrange or
' another style pass invalidating it again.
Private Sub FinishPersistedNativeLabels(ByRef d As SldWorks.ModelDoc2)
    Dim path As String, er As Long, wr As Long, sheetName As Variant, sheetNames As Variant, work As Variant
    Dim view As SldWorks.View, views As Collection, regions As Object, item As Variant, index As Long, active As Object, saved As Boolean
    path = d.GetPathName
    Set active = sw.ActivateDoc3(d.GetTitle, False, SW_REBUILD_USER_DECISION, er)
    d.SetSaveFlag
    saved = d.Save3(SW_SAVE_SILENT, er, wr)
    Log "PERSISTED LABEL SAVE saved=" & saved & "; errors=" & er & "; dirty=" & d.GetSaveFlag
    If saved = False Then
        Err.Raise vbObjectError + ERR_SAVE_DRAWING, , "Cannot persist native labels before their final check: " & er
    End If
    sw.CloseDoc d.GetTitle
    Set d = sw.OpenDoc6(path, SW_DOC_DRAWING, SW_OPEN_SILENT, "", er, wr)
    If d Is Nothing Then
        Err.Raise vbObjectError + ERR_REOPEN_DRAWING, , "Cannot reopen the drawing for its final text check: " & er
    End If
    Set currentDrawing = d
    PrepareDocumentWindow d
    Set active = sw.ActivateDoc3(d.GetTitle, False, SW_REBUILD_USER_DECISION, er)
    d.ActiveView.EnableGraphicsUpdate = True
    d.ForceRebuild3 False
    d.GraphicsRedraw2
    sheetNames = AsDrawing(d).GetSheetNames
    For index = LBound(sheetNames) To UBound(sheetNames)
        sheetName = sheetNames(index)
        AsDrawing(d).ActivateSheet CStr(sheetName)
        AsDrawing(d).ActivateView ""
        work = Array(LEADER_WORK_LEFT_M, LEADER_WORK_BOTTOM_M, sheetWidth - LEADER_WORK_EDGE_MARGIN_M, sheetHeight - LEADER_WORK_EDGE_MARGIN_M)
        If index > LBound(sheetNames) Then
            work(1) = LAYOUT_CONTINUATION_BOTTOM_M
        End If
        Set views = New Collection
        Set regions = CreateObject("Scripting.Dictionary")
        Set view = AsDrawing(d).GetFirstView.GetNextView
        Do While Not view Is Nothing
            views.Add view
            Set view = view.GetNextView
        Loop
        For Each item In views
            Set view = item
            regions.Add view.Name, NativeViewRegion(view, views, work)
        Next item
        If Not TrySemanticLabelLayout(d, views, regions, False) Then
            If Setting("StrictDimensionLayout", "1") = "1" Then
                Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "The final persisted drawing still has conflicting native labels."
            End If
            Log "WARNING persisted label conflicts remain on " & sheetName
        End If
    Next index
    AsDrawing(d).ActivateSheet CStr(sheetNames(LBound(sheetNames)))
End Sub
