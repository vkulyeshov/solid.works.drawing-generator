Attribute VB_Name = "memory_layout"
Option Explicit

Private Function MemoryLayoutEnabled() As Boolean
    Dim engine As String
    engine = LCase$(Setting("LayoutEngine", DEFAULT_LAYOUT_ENGINE))
    If engine <> MEMORY_LAYOUT_ENGINE And engine <> LEGACY_LAYOUT_ENGINE Then
        Err.Raise vbObjectError + ERR_DRAWING_SETTINGS, , "LayoutEngine must be memory or native."
    End If
    MemoryLayoutEnabled = engine = MEMORY_LAYOUT_ENGINE
End Function

' Read the actual font, references and graphics once. The search never calls CAD.
Private Function MemoryCapture(ByVal display As SldWorks.DisplayDimension, ByVal view As SldWorks.View) As Object
    Dim result As Object, data As SldWorks.DisplayData, heads As New Collection, triangle As Variant
    Dim first As Variant, last As Variant, direction As Variant, index As Long, arc As Variant
    Set data = display.GetDisplayData
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "name", display.GetNameForSelection
    result.Add "owner", view.Name
    result.Add "type", LayoutDimensionKind(display.Type2)
    result.Add "sectorState", 0
    result.Add "position", display.GetAnnotation.GetPosition
    result.Add "controlPoint", display.GetAnnotation.GetPosition
    result.Add "centerText", display.CenterText
    result.Add "offsetText", display.OffsetText
    result.Add "arrowSide", display.ArrowSide
    result.Add "value", display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue
    result.Add "text", LayoutInkPolygons(display)
    result.Add "lines", PlanarDisplayLines(data)
    result.Add "bases", PlanarDimensionBases(display)
    For index = 0 To data.GetArrowHeadCount - 1
        triangle = PlanarArrowTriangle(data.GetArrowHeadAtIndex2(index))
        If Not IsEmpty(triangle) Then
            heads.Add triangle
        End If
    Next index
    result.Add "heads", heads
    If LayoutDimensionKind(display.Type2) = SW_DIMENSION_LINEAR And data.GetArrowHeadCount > 0 Then
        first = data.GetArrowHeadAtIndex2(0)
        last = first
        If data.GetArrowHeadCount = DIMENSION_REFERENCE_SECOND Then
            last = data.GetArrowHeadAtIndex2(1)
        End If
        direction = LayoutLinearDirection(display)
        If Not IsEmpty(direction) Then
            result.Add "direction", direction
            result.Add "first", MemoryWitnessFoot(first, result("lines"), direction(0), direction(1))
            result.Add "last", MemoryWitnessFoot(last, result("lines"), direction(0), direction(1))
            result.Add "baseline", Array((first(0) + last(0)) / 2, (first(1) + last(1)) / 2)
            ' SetPosition2 expects the native annotation anchor. For rotated
            ' centered text that point differs from the arrow midpoint by the
            ' text offset; replacing it with the baseline makes every repair drift.
        End If
    ElseIf display.Type2 = SW_DIMENSION_ANGULAR And data.GetArcCount > 0 Then
        arc = data.GetArcAtIndex2(0)
        result.Add "arcCenter", Array(arc(DISPLAY_ARC_CENTER_X), arc(DISPLAY_ARC_CENTER_Y))
        result.Add "arcRadius", Sqr((arc(DISPLAY_LINE_START_X) - arc(DISPLAY_ARC_CENTER_X)) ^ 2 + (arc(DISPLAY_LINE_START_Y) - arc(DISPLAY_ARC_CENTER_Y)) ^ 2)
    ElseIf data.GetArrowHeadCount > 0 Then
        first = data.GetArrowHeadAtIndex2(0)
        result.Add "anchor", Array(first(0), first(1))
    End If
    result.Add "cost", PlanarGraphicLength(result("lines"))
    MemoryRebound result
    Set MemoryCapture = result
End Function

Private Sub MemoryPrepare(ByVal display As SldWorks.DisplayDimension)
    Dim format As SldWorks.TextFormat, data As SldWorks.DisplayData, first As Variant, last As Variant
    Dim span As Double, width As Double, text As Variant, direction As Variant, position As Variant
    Set format = display.GetAnnotation.GetTextFormat(SW_ANNOTATION_TEXT_FORMAT_FIRST)
    format.TypeFaceName = Setting("Font", DEFAULT_FONT)
    format.CharHeight = TEXT_HEIGHT_SMALL_M
    If format.CharHeight <= 0 Or display.IsHoleCallout Then
        format.CharHeight = TEXT_HEIGHT_SMALL_M
    End If
    display.GetAnnotation.SetTextFormat SW_ANNOTATION_TEXT_FORMAT_FIRST, False, format
    If LayoutDimensionKind(display.Type2) = SW_DIMENSION_LINEAR Or display.Type2 = SW_DIMENSION_ANGULAR Then
        display.OffsetText = False
        display.CenterText = True
        Set data = display.GetDisplayData
        If data.GetArrowHeadCount = DIMENSION_REFERENCE_SECOND Then
            first = data.GetArrowHeadAtIndex2(0)
            last = data.GetArrowHeadAtIndex2(1)
            span = Sqr((first(0) - last(0)) ^ 2 + (first(1) - last(1)) ^ 2)
            text = LayoutInkBounds(LayoutInkPolygons(display))
            width = text(2) - text(0)
            If LayoutDimensionKind(display.Type2) = SW_DIMENSION_LINEAR Then
                direction = LayoutLinearDirection(display)
                If Not IsEmpty(direction) Then
                    width = Abs(direction(0)) * (text(2) - text(0)) + Abs(direction(1)) * (text(3) - text(1))
                End If
            End If
            If span >= 2 * GOST_ARROW_HEAD_LENGTH_M + DIMENSION_GEOMETRY_GAP_M Then
                display.ArrowSide = SW_DIMENSION_ARROWS_INSIDE
            Else
                display.ArrowSide = SW_DIMENSION_ARROWS_OUTSIDE
            End If
            If LayoutDimensionKind(display.Type2) = SW_DIMENSION_LINEAR And span < width + 2 * GOST_ARROW_HEAD_LENGTH_M + DIMENSION_GEOMETRY_GAP_M And Not IsEmpty(direction) Then
                display.CenterText = False
                position = Array((first(0) + last(0)) / 2 + direction(0) * (span / 2 + GOST_ARROW_LENGTH_M + width / 2), (first(1) + last(1)) / 2 + direction(1) * (span / 2 + GOST_ARROW_LENGTH_M + width / 2), 0)
                display.GetAnnotation.SetPosition2 position(0), position(1), position(2)
            ElseIf LayoutDimensionKind(display.Type2) = SW_DIMENSION_LINEAR Then
                ' Some newly created projected dimensions retain their old
                ' text anchor until explicitly placed on the measured midpoint.
                display.GetAnnotation.SetPosition2 (first(0) + last(0)) / 2, (first(1) + last(1)) / 2, 0
                display.CenterText = True
            End If
        End If
    Else
        display.CenterText = False
        display.SolidLeader = False
        display.SetBrokenLeader2 False, SW_BROKEN_LEADER_HORIZONTAL_TEXT
    End If
End Sub

Private Sub MemoryAddCandidate(ByVal candidate As Object, ByVal record As Object, ByVal obstacles As Collection)
    Dim existing As Object, first As Variant, last As Variant
    PollExecutionBudget
    If Not MemoryCandidateFits(candidate, record("region"), obstacles, record("material"), record("outline")) Then
        Exit Sub
    End If
    first = candidate("position")
    For Each existing In record("candidates")
        last = existing("position")
        If Abs(first(0) - last(0)) + Abs(first(1) - last(1)) < PROFILE_MATERIAL_TOLERANCE_M Then
            Exit Sub
        End If
    Next existing
    candidate("id") = candidate("name") & "|" & CStr(record("candidates").Count)
    PlanarInsertCandidate record("candidates"), candidate
End Sub

Private Sub MemoryBuildCandidates(ByVal record As Object, ByVal obstacles As Collection)
    Dim seed As Object, candidate As Object, direction As Variant, baseline As Variant, outline As Variant, alternate As Object, textSide As Long
    Dim nx As Double, ny As Double, low As Double, high As Double, coordinate As Double, point As Variant
    Dim row As Long, side As Long, index As Long, delta As Double, angle As Double, radius As Double
    Dim center As Variant, width As Double, height As Double, distance As Double
    Dim localLow As Double, localHigh As Double, origin As Long
    Set record("candidates") = New Collection
    Set planarLayoutRejections = CreateObject("Scripting.Dictionary")
    Set seed = record("seed")
    If record.Exists("reference") Then
        Set seed = ReferenceLabelCandidate(seed, record("reference"), seed("position")(0), seed("position")(1))
    End If
    seed("id") = seed("name") & "|original"
    seed("cost") = CDbl(seed("cost")) + MEMORY_UNCHANGED_COST_M
    MemoryAddCandidate seed, record, obstacles
    If record.Exists("reference") Then
        MemoryReferenceCandidates record, obstacles
        Exit Sub
    End If
    outline = record("outline")
    If seed.Exists("direction") Then
        direction = seed("direction")
        nx = -direction(1): ny = direction(0)
        baseline = seed("baseline")
        low = SEARCH_INITIAL_LIMIT: high = -SEARCH_INITIAL_LIMIT
        For Each point In MemoryBoxPolygon(outline)
            coordinate = point(0) * nx + point(1) * ny
            low = MinD(low, coordinate): high = MaxD(high, coordinate)
        Next point
        point = seed("first")
        localLow = point(0) * nx + point(1) * ny
        point = seed("last")
        localHigh = point(0) * nx + point(1) * ny
        coordinate = localLow
        localLow = MinD(coordinate, localHigh)
        localHigh = MaxD(coordinate, localHigh)
        ' Local lanes can occupy a notch or a gap within the view. The material
        ' test, rather than the view bounding rectangle, decides whether it fits.
        For origin = 0 To 1
          For row = 0 To MEMORY_LAYOUT_ROWS
            For side = 0 To 1
                coordinate = low - GOST_OBJECT_OFFSET_M - row * MEMORY_LAYOUT_ROW_PITCH_M
                If side = 1 Then
                    coordinate = high + GOST_OBJECT_OFFSET_M + row * MEMORY_LAYOUT_ROW_PITCH_M
                End If
                If origin = 1 Then
                    coordinate = localLow - GOST_OBJECT_OFFSET_M - row * MEMORY_LOCAL_ROW_PITCH_M
                    If side = 1 Then
                        coordinate = localHigh + GOST_OBJECT_OFFSET_M + row * MEMORY_LOCAL_ROW_PITCH_M
                    End If
                End If
                delta = coordinate - baseline(0) * nx - baseline(1) * ny
                Set candidate = MemoryTranslateLinearCandidate(seed, nx, ny, delta)
                candidate("baseline") = Array(baseline(0) + nx * delta, baseline(1) + ny * delta)
                MemoryAddCandidate candidate, record, obstacles
                ' Keep a fitting centered label centered. Native imported
                ' dimensions can rebuild their baseline differently when a
                ' previously centered label is made free. Short dimensions
                ' already prepared as free text still offer both end positions.
                If seed("centerText") = False And seed("heads").Count = DIMENSION_REFERENCE_SECOND Then
                    For textSide = -1 To 1 Step 2
                        Set alternate = MemoryLinearTextCandidate(candidate, textSide)
                        MemoryAddCandidate alternate, record, obstacles
                    Next textSide
                End If
            Next side
          Next row
        Next origin
    ElseIf seed.Exists("arcCenter") Then
        If seed("arcRadius") > MEMORY_ANGULAR_MIN_RADIUS_M Then
            For row = 0 To MEMORY_LAYOUT_ROWS
                radius = GOST_OBJECT_OFFSET_M + row * MEMORY_LAYOUT_ROW_PITCH_M
                Set candidate = MemoryAngularCandidate(seed, radius)
                MemoryAddCandidate candidate, record, obstacles
                Set alternate = MemoryOppositeAngularCandidate(candidate)
                MemoryAddCandidate alternate, record, obstacles
            Next row
        End If
    ElseIf seed.Exists("anchor") Then
        center = seed("anchor")
        For row = 0 To MEMORY_LAYOUT_ROWS
            distance = GOST_OBJECT_OFFSET_M + row * MEMORY_LAYOUT_ROW_PITCH_M
            For index = 0 To MEMORY_LAYOUT_DIRECTIONS - 1
                angle = index * FULL_REVOLUTION_RAD / MEMORY_LAYOUT_DIRECTIONS
                Set candidate = MemoryLeaderCandidate(seed, center(0) + distance * Cos(angle), center(1) + distance * Sin(angle))
                MemoryAddCandidate candidate, record, obstacles
                If record("candidates").Count >= MEMORY_CANDIDATE_LIMIT Then
                    Exit For
                End If
            Next index
            If record("candidates").Count >= MEMORY_CANDIDATE_LIMIT Then
                Exit For
            End If
        Next row
    End If
End Sub

Private Sub MemoryApply(ByVal d As SldWorks.ModelDoc2, ByVal display As SldWorks.DisplayDimension, ByVal candidate As Object, ByVal record As Object)
    Dim point As Variant, changed As Boolean
    If display.Type2 = SW_DIMENSION_ANGULAR Then
        If CLng(candidate("sectorState")) <> CLng(record("sectorState")) Then
            SelectPlanarAngularDimension d, display
            changed = display.VerticallyOppositeAngle
            If changed = False Then
                Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot apply the calculated angular sector."
            End If
            record("sectorState") = CLng(candidate("sectorState"))
        End If
    End If
    point = candidate("controlPoint")
    display.OffsetText = candidate("offsetText")
    display.CenterText = candidate("centerText")
    display.ArrowSide = CLng(candidate("arrowSide"))
    display.GetAnnotation.SetPosition2 point(0), point(1), point(2)
    display.CenterText = candidate("centerText")
End Sub

Private Function MemoryCanSampleNatively(ByVal seed As Object) As Boolean
    If seed("type") = SW_DIMENSION_LINEAR Then
        MemoryCanSampleNatively = seed.Exists("direction")
    ElseIf seed("type") = SW_DIMENSION_DIAMETER Or seed("type") = SW_DIMENSION_RADIAL Then
        MemoryCanSampleNatively = seed.Exists("anchor")
    End If
End Function

' Native anchors can clamp to a measured axis or reattach to a circle.
' Sample only annotations whose calculated representation proved unstable.
Private Sub MemoryNativeCandidates(ByVal d As SldWorks.ModelDoc2, ByVal display As SldWorks.DisplayDimension, ByVal view As SldWorks.View, ByVal record As Object)
    Dim seed As Object, direction As Variant, point As Variant, target As Variant
    Dim row As Long, side As Long, textSide As Long, distance As Double, nx As Double, ny As Double
    Dim index As Long, angle As Double
    Set seed = record("seed")
    Set record("candidates") = New Collection
    Set planarLayoutRejections = CreateObject("Scripting.Dictionary")
    If seed("type") = SW_DIMENSION_LINEAR Then
        direction = seed("direction")
        nx = -direction(1): ny = direction(0)
        point = seed("controlPoint")
        For row = 0 To MEMORY_NATIVE_LINEAR_ROWS
            For side = -1 To 1 Step 2
                distance = side * row * MEMORY_LAYOUT_ROW_PITCH_M
                For textSide = -1 To 1
                    target = Array(point(0) + nx * distance + direction(0) * textSide * MEMORY_NATIVE_LINEAR_TEXT_SHIFT_M, point(1) + ny * distance + direction(1) * textSide * MEMORY_NATIVE_LINEAR_TEXT_SHIFT_M, point(2))
                    MemorySampleNativeCandidate d, display, view, record, target
                Next textSide
            Next side
        Next row
    Else
        point = seed("anchor")
        For row = 1 To MEMORY_NATIVE_LINEAR_ROWS
            distance = row * MEMORY_NATIVE_LINEAR_TEXT_SHIFT_M
            For index = 0 To MEMORY_NATIVE_CIRCULAR_DIRECTIONS - 1
                angle = index * FULL_REVOLUTION_RAD / MEMORY_NATIVE_CIRCULAR_DIRECTIONS
                target = Array(point(0) + distance * Cos(angle), point(1) + distance * Sin(angle), seed("controlPoint")(2))
                MemorySampleNativeCandidate d, display, view, record, target
            Next index
        Next row
    End If
    MemoryApply d, display, seed, record
    Log "MEMORY native exceptional dimension samples " & record("name") & "; candidates=" & record("candidates").Count
End Sub

Private Sub MemorySampleNativeCandidate(ByVal d As SldWorks.ModelDoc2, ByVal display As SldWorks.DisplayDimension, ByVal view As SldWorks.View, ByVal record As Object, ByVal target As Variant)
    Dim seed As Object, candidate As Object
    CheckExecutionBudget
    Set seed = record("seed")
    display.OffsetText = False
    display.CenterText = False
    display.ArrowSide = CLng(seed("arrowSide"))
    display.GetAnnotation.SetPosition2 target(0), target(1), target(2)
    d.GraphicsRedraw2
    Set candidate = MemoryCapture(display, view)
    If Abs(CDbl(candidate("value")) - CDbl(seed("value"))) > LEADER_VALUE_TOLERANCE_M Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "A model measurement changed during native verification."
    End If
    candidate("controlPoint") = target
    MemoryAddCandidate candidate, record, record("obstacles")
End Sub

Private Function MemoryRecordRegion(ByVal view As SldWorks.View, ByVal frame As Variant) As Variant
    ' Joint ownership and actual occupied areas replace artificial half-sheet
    ' cells. Dimensions may use any free space inside this template's frame.
    MemoryRecordRegion = frame
End Function

' Native detail captions inherit the application's UI language. Replace only
' generated detail captions; model notes and linked properties remain intact.
Private Sub MemoryPrepareViewCaption(ByVal view As SldWorks.View)
    Dim notes As Variant, item As Variant, note As SldWorks.Note, format As SldWorks.TextFormat
    If planarDetailLabels Is Nothing Then
        Exit Sub
    End If
    If Not planarDetailLabels.Exists(view.Name) Then
        Exit Sub
    End If
    notes = view.GetNotes
    If IsEmpty(notes) Then
        Exit Sub
    End If
    For Each item In notes
        Set note = item
        note.SetText CStr(planarDetailLabels(view.Name)) & " (" & ScaleText(view.ScaleDecimal) & ")"
        Set format = note.GetAnnotation.GetTextFormat(SW_ANNOTATION_TEXT_FORMAT_FIRST)
        format.TypeFaceName = Setting("Font", DEFAULT_FONT)
        format.CharHeight = TEXT_HEIGHT_STANDARD_M
        format.Italic = False
        format.Escapement = 0
        note.GetAnnotation.SetTextFormat SW_ANNOTATION_TEXT_FORMAT_FIRST, False, format
    Next item
End Sub

' A source-area outline is annotation geometry, not material. In particular,
' a detail marker must not enlarge the part's dimension lanes or block a
' neighbouring detail through IView.GetOutline's sketch-inclusive bounds.
Private Function MemoryViewOutline(ByVal view As SldWorks.View) As Variant
    Dim material As Variant, polygon As Variant, point As Variant, bounds As Variant, raw() As Variant, index As Long
    MemoryViewOutline = view.GetOutline
    If Not profileSheetContours Is Nothing Then
        If profileSheetContours.Exists(view.Name) Then
            material = profileSheetContours(view.Name)
        End If
    End If
    If IsEmpty(material) And Not planarContourPoints Is Nothing Then
        ReDim raw(planarContourPoints.Count - 1)
        For index = 1 To planarContourPoints.Count
            raw(index - 1) = PatternSheetPoint(view, planarContourPoints(index))
        Next index
        polygon = PlanarClipPolygon(raw, view.GetOutline)
        If Not IsEmpty(polygon) Then
            material = Array(polygon)
        End If
    End If
    If IsEmpty(material) Then
        Exit Function
    End If
    bounds = Array(SEARCH_INITIAL_LIMIT, SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT)
    For Each polygon In material
        For Each point In polygon
            NativeExpandBounds bounds, point(0), point(1)
        Next point
    Next polygon
    If bounds(2) > bounds(0) And bounds(3) > bounds(1) Then
        MemoryViewOutline = bounds
    End If
End Function

' Shared entry point for panels, inclined profiles, solids, sheet metal and assemblies.
' selectedView is used while a paged planner builds one local feature family.
Private Sub FinishMemoryDimensionLayout(ByVal d As SldWorks.ModelDoc2, Optional ByVal selectedView As String = "", Optional ByVal prepare As Boolean = True)
    Dim sheetName As Variant, sheetNames As Variant, views As Collection, view As SldWorks.View
    Dim display As SldWorks.DisplayDimension, records As Collection, record As Object, workspace As Object
    Dim obstacles As Collection, ownObstacles As Collection, obstacle As Variant, item As Variant, other As Variant
    Dim originals As Object, displays As Object, owners As Object, expected As Object, snapshot As Object, recordMap As Object
    Dim candidates As Collection, chosen As Collection, pending As Collection, actual As Collection, valid As Object
    Dim pass As Long, steps As Long, count As Long, stream As Object, started As Double, frame As Variant
    Dim name As String, candidate As Object, ok As Boolean, value As Double, errorText As String, symmetry As Object
    started = Timer
    If selectedView = "" And prepare Then
        If SixViewMode() And Not profileRoleDrawing Then
            AssignSixViewDimensions d
        ElseIf Not profileRoleDrawing Then
            Set view = AsDrawing(d).GetFirstView.GetNextView
            If Not view Is Nothing Then
                Set symmetry = NativeSymmetryEvidence(view)
                DeduplicateNativeDimensions d, symmetry
            End If
        End If
        PrepareNativeCalloutText d
    End If
    sheetNames = AsDrawing(d).GetSheetNames
    For Each sheetName In sheetNames
        BeginLayoutBudget CStr(sheetName)
        AsDrawing(d).ActivateSheet CStr(sheetName)
        AsDrawing(d).ActivateView ""
        AsDrawing(d).GetCurrentSheet.SheetFormatVisible = True
        AsDrawing(d).EditTemplate
        AsDrawing(d).EditSheet
        d.ForceRebuild3 False
        Set views = New Collection
        Set records = New Collection
        Set displays = CreateObject("Scripting.Dictionary")
        Set owners = CreateObject("Scripting.Dictionary")
        Set expected = CreateObject("Scripting.Dictionary")
        Set view = AsDrawing(d).GetFirstView.GetNextView
        Do While Not view Is Nothing
            views.Add view
            If prepare Then
                MemoryPrepareViewCaption view
            End If
            Set display = view.GetFirstDisplayDimension5
            Do While Not display Is Nothing
                If selectedView = "" Or selectedView = view.Name Then
                    name = display.GetNameForSelection
                    displays.Add name, display
                    owners.Add name, view
                    expected.Add name, display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue
                    If prepare Then
                        MemoryPrepare display
                    End If
                End If
                Set display = display.GetNext5
            Loop
            Set view = view.GetNextView
        Loop
        If displays.Count = 0 Then
            GoTo NextMemorySheet
        End If
        ' Normalize once and read the resulting native graphics once per annotation.
        d.ClearSelection2 True
        For Each item In displays.Keys
            displays(item).GetAnnotation.Select3 True, Nothing
        Next item
        d.BreakDimensionAlignment
        d.ClearSelection2 True
        d.GraphicsRedraw2
        CacheProfileContours views
        Set workspace = SheetWorkspace(d)
        frame = workspace("frame")
        Set obstacles = workspace("obstacles")
        For Each item In displays.Keys
            Set view = owners(item)
            Set record = CreateObject("Scripting.Dictionary")
            record.Add "name", CStr(item)
            record.Add "sectorState", 0
            record.Add "region", MemoryRecordRegion(view, frame)
            record.Add "outline", MemoryViewOutline(view)
            record.Add "material", LayoutMaterial(displays(item), view)
            record.Add "seed", MemoryCapture(displays(item), view)
            record.Add "candidates", New Collection
            Set ownObstacles = New Collection
            For Each obstacle In obstacles
                ownObstacles.Add obstacle
            Next obstacle
            For Each other In views
                If other.Name <> view.Name Then
                    ownObstacles.Add MemoryViewOutline(other)
                End If
            Next other
            If Not planarDetailLabelBoxes Is Nothing Then
                If planarDetailLabelBoxes.Exists(view.Name) Then
                    For Each obstacle In planarDetailLabelBoxes(view.Name)
                        ownObstacles.Add obstacle
                    Next obstacle
                End If
            End If
            record.Add "obstacles", ownObstacles
            records.Add record
        Next item
        MemoryAppendReferences records, views, frame, obstacles, CStr(sheetName)
        Set recordMap = CreateObject("Scripting.Dictionary")
        For Each record In records
            recordMap.Add record("name"), record
        Next record
        For pass = 0 To MEMORY_LAYOUT_REPAIR_LIMIT
            Set chosen = New Collection
            Set pending = New Collection
            For Each record In records
                Set ownObstacles = record("obstacles")
                If record.Exists("nativeSampled") Then
                    MemoryNativeCandidates d, displays(record("name")), owners(record("name")), record
                Else
                    MemoryBuildCandidates record, ownObstacles
                    Set snapshot = record("seed")
                    If record("candidates").Count = 0 And MemoryCanSampleNatively(snapshot) Then
                        record.Add "nativeSampled", True
                        MemoryNativeCandidates d, displays(record("name")), owners(record("name")), record
                    End If
                End If
                Log "MEMORY CANDIDATES " & record("name") & "; count=" & record("candidates").Count
                If record("candidates").Count = 0 Then
                    Log "MEMORY REJECT " & record("name") & "; " & Join(planarLayoutRejections.Keys, ", ")
                    MemoryWriteSeed record
                    Set view = owners(record("name"))
                    profilePageFailureView = view.Name
                    If Not planarDimensionRequests Is Nothing Then
                        RequestPlanarSplit view.Name
                    End If
                    Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "No in-memory placement fits the template workspace: " & record("name")
                End If
                pending.Add record
            Next record
            planarStrictCrossings = True
            planarSearchLimit = MEMORY_LAYOUT_SEARCH_LIMIT
            Set planarLayoutCompatibility = CreateObject("Scripting.Dictionary")
            steps = 0
            BeginStrictSearch
            If MemorySolveCandidates(pending, chosen, steps) = False Then
                EndStrictSearch
                Log "MEMORY strict witness crossing search exhausted; trying the permitted witness crossings on " & sheetName
                planarStrictCrossings = False
                planarSearchLimit = MEMORY_LAYOUT_RELAXED_SEARCH_LIMIT
                Set planarLayoutCompatibility = CreateObject("Scripting.Dictionary")
                Set chosen = New Collection
                steps = 0
                If MemorySolveCandidates(pending, chosen, steps) Then
                    GoTo MemoryAssignmentReady
                End If
                Log "MEMORY assignment exhausted; steps=" & steps
                MemoryAuditIncompatiblePairs records
                For Each record In records
                    MemoryWriteSeed record
                Next record
                Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "No collision-free in-memory assignment fits sheet " & sheetName
            End If
MemoryAssignmentReady:
            EndStrictSearch
            ImprovePlanarAssignment records, chosen
            For Each candidate In chosen
                MemoryAuditGeometry candidate, "expected|" & CStr(pass)
                If displays.Exists(candidate("name")) Then
                    MemoryApply d, displays(candidate("name")), candidate, recordMap(candidate("name"))
                Else
                    MemoryApplyReference candidate
                End If
            Next candidate
            ' One redraw for the entire assignment, never one per candidate.
            d.GraphicsRedraw2
            Set actual = New Collection
            ok = True
            For Each record In records
                name = record("name")
                If record.Exists("reference") Then
                    Set snapshot = MemoryCaptureReference(record("reference"))
                Else
                    Set display = displays(name)
                    value = display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue
                    If Abs(value - CDbl(expected(name))) > LEADER_VALUE_TOLERANCE_M Then
                        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "A model measurement changed during memory layout: " & name
                    End If
                    Set snapshot = MemoryCapture(display, owners(name))
                    snapshot("sectorState") = CLng(record("sectorState"))
                End If
                MemoryAuditGeometry snapshot, "actual|" & CStr(pass)
                Set planarLayoutRejections = CreateObject("Scripting.Dictionary")
                If MemoryCandidateFits(snapshot, record("region"), record("obstacles"), record("material"), record("outline")) = False Then
                    ok = False
                    Log "MEMORY NATIVE mismatch " & name & "; pass=" & pass & "; " & Join(planarLayoutRejections.Keys, ", ")
                    Set record("seed") = snapshot
                    MemoryWriteSeed record
                    If MemoryCanSampleNatively(snapshot) Then
                        record("nativeSampled") = True
                    End If
                End If
                For Each other In actual
                    If PlanarCandidatesCompatible(snapshot, other) = False Then
                        ok = False
                        Log "MEMORY NATIVE collision " & name & " | " & other("name")
                        If MemoryCanSampleNatively(snapshot) Then
                            record("nativeSampled") = True
                        End If
                        If MemoryCanSampleNatively(other) Then
                            recordMap(other("name"))("nativeSampled") = True
                        End If
                    End If
                Next other
                actual.Add snapshot
                Set record("seed") = snapshot
            Next record
            Log "MEMORY LAYOUT sheet=" & sheetName & "; dimensions=" & displays.Count & "; reference letters=" & (records.Count - displays.Count) & "; search steps=" & steps & "; repair=" & pass & "; verified=" & ok
            If ok Then
                Exit For
            End If
        Next pass
        If ok = False Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Native graphics differ from the in-memory assignment after bounded repair: " & sheetName
        End If
        RenderReferenceLeaders d, actual
        Set stream = fs.OpenTextFile(fs.BuildPath(runDir, MEMORY_LAYOUT_AUDIT_FILE), FILE_FOR_APPENDING, True, FILE_UNICODE)
        For Each candidate In actual
            stream.WriteLine CStr(jobNumber) & vbTab & sheetName & vbTab & candidate("owner") & vbTab & candidate("name") & vbTab & candidate("type") & vbTab & candidate("value") & vbTab & Join(candidate("bounds"), vbTab)
        Next candidate
        stream.Close
        count = count + displays.Count
        If prepare And Not profileRoleDrawing Then
            If symmetry Is Nothing Then
                Set symmetry = NativeSymmetryEvidence(views(1))
            End If
            For Each item In views
                AddNativeSymmetryCenterline d, item, symmetry
            Next item
        End If
NextMemorySheet:
        EndLayoutBudget
    Next sheetName
    If selectedView = "" And prepare And SixViewMode() And Not profileRoleDrawing Then
        RenderHolePatternCircles d
        AuditFinalSixViews d
    End If
    Log "MEMORY COMPLETE dimensions=" & count & "; seconds=" & Format$(PlanarElapsedSeconds(started), "0.0") & "; native Auto Arrange calls=0"
End Sub

' Optional offline/native comparison for diagnosis; it never moves an annotation.
Private Sub MemoryAuditGeometry(ByVal candidate As Object, ByVal phase As String)
    Dim stream As Object, line As Variant, polygon As Variant
    If Setting("DimensionLayoutAudit", "0") <> "1" Then
        Exit Sub
    End If
    Set stream = fs.OpenTextFile(fs.BuildPath(runDir, MEMORY_PREDICTION_AUDIT_FILE), FILE_FOR_APPENDING, True, FILE_UNICODE)
    stream.WriteLine phase & vbTab & candidate("name") & vbTab & "position" & vbTab & PlanarAuditCoordinates(candidate("position")) & vbTab & candidate("centerText") & vbTab & candidate("offsetText")
    stream.WriteLine phase & vbTab & candidate("name") & vbTab & "control" & vbTab & PlanarAuditCoordinates(candidate("controlPoint"))
    For Each line In candidate("lines")
        stream.WriteLine phase & vbTab & candidate("name") & vbTab & "line" & vbTab & Join(line, vbTab)
    Next line
    For Each polygon In candidate("text")
        stream.WriteLine phase & vbTab & candidate("name") & vbTab & "text" & vbTab & Join(LayoutInkBoundsOfPolygon(polygon), vbTab)
    Next polygon
    stream.Close
End Sub

' Each child receives already filtered domains. Rechecking every original
' candidate against the complete chosen prefix makes dense sheets needlessly
' expensive even though no further CAD operations are involved.
Private Function MemorySolveCandidates(ByVal remaining As Collection, ByVal chosen As Collection, ByRef steps As Long) As Boolean
    Dim record As Object, best As Object, candidate As Object, alternative As Object
    Dim nextRemaining As Collection, available As Collection, child As Object, possible As Boolean
    CheckExecutionBudget
    If MemoryStrictSearchExpired() Then
        steps = planarSearchLimit + 1
        Exit Function
    End If
    steps = steps + 1
    If steps > planarSearchLimit Then
        Exit Function
    End If
    If remaining.Count = 0 Then
        MemorySolveCandidates = True
        Exit Function
    End If
    For Each record In remaining
        If record("candidates").Count = 0 Then
            Exit Function
        End If
        If best Is Nothing Then
            Set best = record
        ElseIf record("candidates").Count < best("candidates").Count Then
            Set best = record
        End If
    Next record
    For Each candidate In best("candidates")
        ' Unwind every parent when a child exhausts the shared budget. Without
        ' this guard parents keep filtering all remaining candidate domains.
        If steps > planarSearchLimit Then
            Exit Function
        End If
        Set nextRemaining = New Collection
        possible = True
        For Each record In remaining
            If record("name") <> best("name") Then
                Set available = New Collection
                For Each alternative In record("candidates")
                    If PlanarCandidatesCompatibleCached(candidate, alternative) Then
                        available.Add alternative
                    End If
                    If MemoryStrictSearchExpired() Then
                        steps = planarSearchLimit + 1
                        Exit Function
                    End If
                Next alternative
                If available.Count = 0 Then
                    possible = False
                    Exit For
                End If
                Set child = CreateObject("Scripting.Dictionary")
                child.Add "name", record("name")
                child.Add "candidates", available
                nextRemaining.Add child
            End If
        Next record
        If possible Then
            chosen.Add candidate
            If MemorySolveCandidates(nextRemaining, chosen, steps) Then
                MemorySolveCandidates = True
                Exit Function
            End If
            chosen.Remove chosen.Count
        End If
    Next candidate
End Function

' Diagnose an impossible pair without changing any native annotation.
Private Sub MemoryAuditIncompatiblePairs(ByVal records As Collection)
    Dim firstIndex As Long, lastIndex As Long, first As Object, last As Object
    Dim a As Object, b As Object, compatible As Boolean
    For firstIndex = 1 To records.Count - 1
        Set first = records(firstIndex)
        For lastIndex = firstIndex + 1 To records.Count
            Set last = records(lastIndex)
            compatible = False
            For Each a In first("candidates")
                For Each b In last("candidates")
                    If PlanarCandidatesCompatibleCached(a, b) Then
                        compatible = True
                        Exit For
                    End If
                Next b
                If compatible Then
                    Exit For
                End If
            Next a
            If compatible = False Then
                Log "MEMORY incompatible pair " & first("name") & " | " & last("name")
            End If
        Next lastIndex
    Next firstIndex
End Sub

Private Sub MemoryWriteSeed(ByVal record As Object)
    Dim stream As Object, seed As Object, key As Variant, line As Variant, polygon As Variant, box As Variant
    Set seed = record("seed")
    Set stream = fs.OpenTextFile(fs.BuildPath(runDir, MEMORY_LAYOUT_AUDIT_FILE), FILE_FOR_APPENDING, True, FILE_UNICODE)
    stream.WriteLine "SEED" & vbTab & seed("name") & vbTab & seed("type") & vbTab & PlanarAuditCoordinates(seed("position"))
    stream.WriteLine "region" & vbTab & PlanarAuditCoordinates(record("region"))
    stream.WriteLine "outline" & vbTab & PlanarAuditCoordinates(record("outline"))
    For Each key In Array("bounds", "baseline", "first", "last", "direction")
        If seed.Exists(key) Then
            stream.WriteLine key & vbTab & PlanarAuditCoordinates(seed(key))
        End If
    Next key
    For Each line In seed("lines")
        stream.WriteLine "line" & vbTab & Join(line, vbTab)
    Next line
    For Each polygon In seed("text")
        stream.WriteLine "text" & vbTab & Join(LayoutInkBoundsOfPolygon(polygon), vbTab)
    Next polygon
    stream.Close
End Sub
