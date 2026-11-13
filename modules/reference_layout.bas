Attribute VB_Name = "reference_layout"
Option Explicit

' Reference letters participate in the same solve as dimensions. Fixing a
' letter before routing a witness line can make an otherwise free sheet fail.
Private Sub RegisterMemoryReference(ByVal d As SldWorks.ModelDoc2, ByVal label As String, ByVal x As Double, ByVal y As Double, ByVal owner As String, ByVal sourceBounds As Variant)
    Dim note As SldWorks.Note, annotation As SldWorks.Annotation, format As SldWorks.TextFormat, record As Object
    AsDrawing(d).ActivateView ""
    d.ClearSelection2 True
    Set note = d.InsertNote(label)
    If note Is Nothing Then
        Err.Raise vbObjectError + ERR_INSERT_NOTE, , "Cannot insert a source-area reference letter."
    End If
    Set annotation = note.GetAnnotation
    Set format = annotation.GetTextFormat(SW_ANNOTATION_TEXT_FORMAT_FIRST)
    format.TypeFaceName = Setting("Font", DEFAULT_FONT)
    format.CharHeight = MEMORY_REFERENCE_TEXT_HEIGHT_M
    format.Italic = False
    format.Escapement = 0
    annotation.SetTextFormat SW_ANNOTATION_TEXT_FORMAT_FIRST, False, format
    annotation.Color = DIMENSION_COLOR_BLACK
    annotation.LayerOverride = annotation.LayerOverride Or SW_LAYER_OVERRIDE_COLOR
    annotation.SetPosition2 x, y, 0
    note.SetName MEMORY_REFERENCE_NOTE_PREFIX & label
    If memoryReferenceLabels Is Nothing Then
        Set memoryReferenceLabels = CreateObject("Scripting.Dictionary")
    End If
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "name", MEMORY_REFERENCE_PREFIX & AsDrawing(d).GetCurrentSheet.GetName & "|" & label
    record.Add "note", note
    record.Add "noteName", note.GetName
    record.Add "label", label
    record.Add "sheet", AsDrawing(d).GetCurrentSheet.GetName
    record.Add "owner", owner
    record.Add "source", sourceBounds
    memoryReferenceLabels.Add record("name"), record
End Sub

Private Sub MemoryApplyReference(ByVal candidate As Object)
    Dim reference As Object, position As Variant
    Set reference = memoryReferenceLabels(candidate("name"))
    position = candidate("position")
    reference("note").GetAnnotation.SetPosition2 position(0), position(1), position(2)
End Sub

Private Function IsMemoryReferenceNote(ByVal note As SldWorks.Note, ByVal sheetName As String) As Boolean
    Dim key As Variant, record As Object
    If memoryReferenceLabels Is Nothing Then
        Exit Function
    End If
    For Each key In memoryReferenceLabels.Keys
        Set record = memoryReferenceLabels(key)
        If record("sheet") = sheetName And record("noteName") = note.GetName And record("label") = note.GetText Then
            IsMemoryReferenceNote = True
            Exit Function
        End If
    Next key
End Function

Private Function MemoryCaptureReference(ByVal reference As Object) As Object
    Dim result As Object, extent As Variant, polygons As New Collection
    Set result = CreateObject("Scripting.Dictionary")
    extent = reference("note").GetExtent
    polygons.Add MemoryBoxPolygon(Array(extent(BBOX_X_MIN_INDEX), extent(BBOX_Y_MIN_INDEX), extent(BBOX_X_MAX_INDEX), extent(BBOX_Y_MAX_INDEX)))
    result.Add "name", reference("name")
    result.Add "owner", reference("owner")
    result.Add "type", MEMORY_REFERENCE_KIND
    result.Add "value", 0
    result.Add "position", reference("note").GetAnnotation.GetPosition
    result.Add "controlPoint", result("position")
    result.Add "anchor", result("position")
    result.Add "text", polygons
    result.Add "lines", ReferenceLeaderLines(reference("source"), LayoutInkBounds(polygons))
    result.Add "bases", New Collection
    result.Add "heads", New Collection
    result.Add "centerText", False
    result.Add "offsetText", False
    result.Add "arrowSide", 0
    result.Add "cost", 0
    MemoryRebound result
    Set MemoryCaptureReference = result
End Function

Private Sub MemoryAppendReferences(ByVal records As Collection, ByVal views As Collection, ByVal frame As Variant, ByVal occupied As Collection, ByVal sheetName As String)
    Dim key As Variant, reference As Object, record As Object, item As Variant, obstacle As Variant
    Dim view As SldWorks.View, other As Variant, ownObstacles As Collection, material As Variant
    If memoryReferenceLabels Is Nothing Then
        Exit Sub
    End If
    For Each key In memoryReferenceLabels.Keys
        Set reference = memoryReferenceLabels(key)
        If reference("sheet") = sheetName Then
            Set view = Nothing
            For Each item In views
                If item.Name = reference("owner") Then
                    Set view = item
                    Exit For
                End If
            Next item
            If view Is Nothing Then
                Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "A reference letter lost its source view."
            End If
            Set record = CreateObject("Scripting.Dictionary")
            record.Add "name", CStr(key)
            record.Add "reference", reference
            record.Add "region", frame
            record.Add "outline", MemoryViewOutline(view)
            material = Empty
            If Not profileSheetContours Is Nothing Then
                If profileSheetContours.Exists(view.Name) Then
                    material = profileSheetContours(view.Name)
                End If
            End If
            record.Add "material", material
            record.Add "seed", MemoryCaptureReference(reference)
            record.Add "candidates", New Collection
            Set ownObstacles = New Collection
            For Each obstacle In occupied
                ownObstacles.Add obstacle
            Next obstacle
            For Each other In views
                If other.Name <> view.Name Then
                    ownObstacles.Add MemoryViewOutline(other)
                End If
            Next other
            record.Add "obstacles", ownObstacles
            records.Add record
        End If
    Next key
End Sub

Private Sub MemoryReferenceCandidates(ByVal record As Object, ByVal obstacles As Collection)
    Dim source As Variant, box As Variant, seed As Object, candidate As Object
    Dim cx As Double, cy As Double, rx As Double, ry As Double, distance As Double, angle As Double, direction As Long
    source = record("reference")("source")
    cx = (source(0) + source(2)) / 2: cy = (source(1) + source(3)) / 2
    rx = (source(2) - source(0)) / 2: ry = (source(3) - source(1)) / 2
    Set seed = record("seed")
    For distance = MEMORY_REFERENCE_CAPTION_GAP_M To PLANAR_DETAIL_LABEL_REACH_M Step TEXT_HEIGHT_SMALL_M
        For direction = 0 To MEMORY_LAYOUT_DIRECTIONS - 1
            angle = direction * FULL_REVOLUTION_RAD / MEMORY_LAYOUT_DIRECTIONS
            Set candidate = ReferenceLabelCandidate(seed, record("reference"), cx + (rx + distance) * Cos(angle), cy + (ry + distance) * Sin(angle))
            box = LayoutInkBounds(candidate("text"))
            If NativeBoxesOverlap(box, source, LEADER_LABEL_CLEARANCE_M) = False Then
                candidate("cost") = distance
                MemoryAddCandidate candidate, record, obstacles
            End If
        Next direction
        If record("candidates").Count >= MEMORY_CANDIDATE_LIMIT Then
            Exit For
        End If
    Next distance
End Sub

' Rebuild the leader from the actual oval to the shelf below the printed letter.
Private Function ReferenceLeaderLines(ByVal source As Variant, ByVal text As Variant) As Collection
    Dim result As New Collection, cx As Double, cy As Double, rx As Double, ry As Double
    Dim x As Double, y As Double, dx As Double, dy As Double, factor As Double, left As Double, right As Double
    cx = (source(0) + source(2)) / 2: cy = (source(1) + source(3)) / 2
    rx = (source(2) - source(0)) / 2: ry = (source(3) - source(1)) / 2
    left = text(0) - MEMORY_REFERENCE_SHELF_GAP_M
    right = text(2) + MEMORY_REFERENCE_SHELF_GAP_M
    y = text(1) - MEMORY_REFERENCE_SHELF_GAP_M
    x = left
    If (left + right) / 2 < cx Then
        x = right
    End If
    dx = x - cx: dy = y - cy
    factor = 1 / Sqr((dx / rx) ^ 2 + (dy / ry) ^ 2)
    result.Add Array(cx + dx * factor, cy + dy * factor, x, y)
    result.Add Array(left, y, right, y)
    Set ReferenceLeaderLines = result
End Function

Private Function ReferenceLabelCandidate(ByVal seed As Object, ByVal reference As Object, ByVal x As Double, ByVal y As Double) As Object
    Dim clean As Object, key As Variant, result As Object
    Set clean = CreateObject("Scripting.Dictionary")
    For Each key In seed.Keys
        clean.Add key, seed(key)
    Next key
    Set clean("lines") = New Collection
    Set clean("bases") = New Collection
    Set result = MemoryLeaderCandidate(clean, x, y)
    Set result("lines") = ReferenceLeaderLines(reference("source"), LayoutInkBounds(result("text")))
    result("cost") = PlanarGraphicLength(result("lines"))
    MemoryRebound result
    Set ReferenceLabelCandidate = result
End Function

' Render only the verified complete assignment, once per sheet. The solver
' includes these same segments in every candidate before any CAD creation.
Private Sub RenderReferenceLeaders(ByVal d As SldWorks.ModelDoc2, ByVal actual As Collection)
    Dim candidate As Object, line As Variant, sheet As SldWorks.Sheet, properties As Variant
    Dim sheetScaleValue As Double, segment As SldWorks.SketchSegment, oldFocus As Boolean, stream As Object
    Dim manager As SldWorks.SketchManager, oldAddToDB As Boolean
    If memoryReferenceLabels Is Nothing Then
        Exit Sub
    End If
    Set sheet = AsDrawing(d).GetCurrentSheet
    properties = sheet.GetProperties
    sheetScaleValue = properties(SHEET_SCALE_NUMERATOR_PROPERTY) / properties(SHEET_SCALE_DENOMINATOR_PROPERTY)
    AsDrawing(d).ActivateView ""
    d.ClearSelection2 True
    oldFocus = sheet.FocusLocked
    sheet.FocusLocked = True
    Set manager = d.SketchManager
    oldAddToDB = manager.AddToDB
    manager.AddToDB = True
    Set stream = fs.OpenTextFile(fs.BuildPath(runDir, MEMORY_REFERENCE_GRAPHICS_FILE), FILE_FOR_APPENDING, True, FILE_UNICODE)
    For Each candidate In actual
        If candidate("type") = MEMORY_REFERENCE_KIND Then
            For Each line In candidate("lines")
                d.ClearSelection2 True
                Set segment = manager.CreateLine(line(0) / sheetScaleValue, line(1) / sheetScaleValue, 0, line(2) / sheetScaleValue, line(3) / sheetScaleValue, 0)
                If segment Is Nothing Then
                    manager.AddToDB = oldAddToDB
                    sheet.FocusLocked = oldFocus
                    stream.Close
                    Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot draw a verified source reference leader."
                End If
                StyleReferenceSegment segment
                stream.WriteLine candidate("name") & vbTab & sheet.GetName & vbTab & Join(line, vbTab)
            Next line
        End If
    Next candidate
    stream.Close
    manager.AddToDB = oldAddToDB
    sheet.FocusLocked = oldFocus
    d.ClearSelection2 True
End Sub
