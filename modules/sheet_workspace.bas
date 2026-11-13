Attribute VB_Name = "sheet_workspace"
Option Explicit

' Rebuild the map for the active sheet, including continuation-sheet formats.
' Do not infer the position of a stamp from the paper name or orientation.
Private Function SheetWorkspace(ByVal d As SldWorks.ModelDoc2) As Object
    Dim result As Object, sheet As SldWorks.Sheet, sketch As SldWorks.Sketch
    Dim properties As Variant, frame As Variant, raw As New Collection, obstacles As Collection, box As Variant
    Dim width As Double, height As Double, table As SldWorks.TableAnnotation
    Dim view As SldWorks.View, notes As Variant, note As Variant, extent As Variant
    Dim index As Long, coordinates As Variant
    Dim stream As Object
    Set sheet = AsDrawing(d).GetCurrentSheet
    properties = sheet.GetProperties
    width = properties(SHEET_WIDTH_PROPERTY)
    height = properties(SHEET_HEIGHT_PROPERTY)
    Set sketch = sheet.GetTemplateSketch
    If Not sketch Is Nothing Then
        coordinates = sketch.GetLines2(SW_EXCLUDE_CROSSHATCH)
        If Not IsEmpty(coordinates) Then
            For index = LBound(coordinates) To UBound(coordinates) Step MEMORY_FORMAT_RECORD_STRIDE
                ' A sheet-format sketch stores its line coordinates in sheet
                ' space. Its model transform is not a model-view transform.
                raw.Add Array(coordinates(index + MEMORY_FORMAT_START_INDEX), coordinates(index + MEMORY_FORMAT_START_INDEX + 1), coordinates(index + MEMORY_FORMAT_END_INDEX), coordinates(index + MEMORY_FORMAT_END_INDEX + 1))
            Next index
        End If
    End If
    frame = WorkspaceFrame(raw, width, height)
    Set obstacles = WorkspaceFormatBlocks(raw, frame)
    ' The sheet view owns sheet notes, format notes and title-block fields.
    Set view = AsDrawing(d).GetFirstView
    Do While Not view Is Nothing
        notes = view.GetNotes
        If Not IsEmpty(notes) Then
            For Each note In notes
                If IsMemoryReferenceNote(note, sheet.GetName) Then
                    GoTo NextWorkspaceNote
                End If
                extent = note.GetExtent
                If Not IsEmpty(extent) Then
                    box = Array(extent(BBOX_X_MIN_INDEX), extent(BBOX_Y_MIN_INDEX), extent(BBOX_X_MAX_INDEX), extent(BBOX_Y_MAX_INDEX))
                    If box(2) > box(0) And box(3) > box(1) Then
                        obstacles.Add box
                    End If
                End If
NextWorkspaceNote:
            Next note
        End If
        Set table = view.GetFirstTableAnnotation
        Do While Not table Is Nothing
            extent = WorkspaceTableBounds(table)
            If Not IsEmpty(extent) Then
                obstacles.Add extent
            End If
            Set table = table.GetNext
        Loop
        Set view = view.GetNextView
    Loop
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "frame", frame
    result.Add "obstacles", obstacles
    result.Add "sheet", sheet.GetName
    Set stream = fs.OpenTextFile(fs.BuildPath(runDir, MEMORY_WORKSPACE_AUDIT_FILE), FILE_FOR_APPENDING, True, FILE_UNICODE)
    stream.WriteLine CStr(jobNumber) & vbTab & sheet.GetName & vbTab & "frame" & vbTab & Join(frame, vbTab)
    For Each box In obstacles
        stream.WriteLine CStr(jobNumber) & vbTab & sheet.GetName & vbTab & "occupied" & vbTab & Join(box, vbTab)
    Next box
    stream.Close
    Set SheetWorkspace = result
End Function

Private Function WorkspaceTableBounds(ByVal table As SldWorks.TableAnnotation) As Variant
    Dim position As Variant, width As Double, height As Double, index As Long
    Dim x As Double, y As Double
    position = table.GetAnnotation.GetPosition
    If IsEmpty(position) Then
        Exit Function
    End If
    For index = 0 To table.ColumnCount - 1
        If table.ColumnHidden(index) = False Then
            width = width + table.GetColumnWidth(index)
        End If
    Next index
    For index = 0 To table.RowCount - 1
        If table.RowHidden(index) = False Then
            height = height + table.GetRowHeight(index)
        End If
    Next index
    x = position(0): y = position(1)
    If table.AnchorType = WORKSPACE_TABLE_TOP_RIGHT Or table.AnchorType = WORKSPACE_TABLE_BOTTOM_RIGHT Then
        x = x - width
    End If
    If table.AnchorType = WORKSPACE_TABLE_TOP_LEFT Or table.AnchorType = WORKSPACE_TABLE_TOP_RIGHT Then
        y = y - height
    End If
    WorkspaceTableBounds = Array(x, y, x + width, y + height)
End Function

' Select the nearest complete border on each side, ignoring short stamp cells.
Private Function WorkspaceFrame(ByVal lines As Collection, ByVal width As Double, ByVal height As Double) As Variant
    Dim result As Variant, line As Variant, x As Double, y As Double
    Dim left As Double, right As Double, bottom As Double, top As Double
    left = width: right = 0: bottom = height: top = 0
    For Each line In lines
        If Abs(line(0) - line(2)) <= PROFILE_MATERIAL_TOLERANCE_M And Abs(line(3) - line(1)) >= height * MEMORY_FRAME_MIN_SPAN_RATIO Then
            x = line(0)
            If x > PROFILE_MATERIAL_TOLERANCE_M And x < width * MEMORY_FRAME_EDGE_RATIO Then
                left = MinD(left, x)
            ElseIf x < width - PROFILE_MATERIAL_TOLERANCE_M And x > width * (1 - MEMORY_FRAME_EDGE_RATIO) Then
                right = MaxD(right, x)
            End If
        End If
        If Abs(line(1) - line(3)) <= PROFILE_MATERIAL_TOLERANCE_M And Abs(line(2) - line(0)) >= width * MEMORY_FRAME_MIN_SPAN_RATIO Then
            y = line(1)
            If y > PROFILE_MATERIAL_TOLERANCE_M And y < height * MEMORY_FRAME_EDGE_RATIO Then
                bottom = MinD(bottom, y)
            ElseIf y < height - PROFILE_MATERIAL_TOLERANCE_M And y > height * (1 - MEMORY_FRAME_EDGE_RATIO) Then
                top = MaxD(top, y)
            End If
        End If
    Next line
    If left = width Then
        left = MEMORY_FRAME_FALLBACK_M
    End If
    If right = 0 Then
        right = width - MEMORY_FRAME_FALLBACK_M
    End If
    If bottom = height Then
        bottom = MEMORY_FRAME_FALLBACK_M
    End If
    If top = 0 Then
        top = height - MEMORY_FRAME_FALLBACK_M
    End If
    result = Array(left + MEMORY_FORMAT_LINE_CLEARANCE_M, bottom + MEMORY_FORMAT_LINE_CLEARANCE_M, right - MEMORY_FORMAT_LINE_CLEARANCE_M, top - MEMORY_FORMAT_LINE_CLEARANCE_M)
    WorkspaceFrame = result
End Function

' Connected internal format lines form occupied blocks. Removing the border
' before grouping prevents it from connecting every stamp into one full-page box.
Private Function WorkspaceFormatBlocks(ByVal lines As Collection, ByVal frame As Variant) As Collection
    Dim result As New Collection, line As Variant, box As Variant, other As Variant
    Dim index As Long, merged As Boolean, border As Boolean
    For Each line In lines
        box = WorkspaceInternalBox(line, frame)
        If IsEmpty(box) Then
            GoTo NextWorkspaceLine
        End If
        border = False
        If Abs(line(0) - line(2)) <= PROFILE_MATERIAL_TOLERANCE_M Then
            border = Abs(line(0) - frame(0)) <= MEMORY_FORMAT_LINE_CLEARANCE_M + PROFILE_MATERIAL_TOLERANCE_M Or Abs(line(0) - frame(2)) <= MEMORY_FORMAT_LINE_CLEARANCE_M + PROFILE_MATERIAL_TOLERANCE_M
        ElseIf Abs(line(1) - line(3)) <= PROFILE_MATERIAL_TOLERANCE_M Then
            border = Abs(line(1) - frame(1)) <= MEMORY_FORMAT_LINE_CLEARANCE_M + PROFILE_MATERIAL_TOLERANCE_M Or Abs(line(1) - frame(3)) <= MEMORY_FORMAT_LINE_CLEARANCE_M + PROFILE_MATERIAL_TOLERANCE_M
        End If
        If Not border Then
            Do
                merged = False
                For index = result.Count To 1 Step -1
                    other = result(index)
                    If NativeBoxesOverlap(box, other, MEMORY_FORMAT_JOIN_M) Then
                        box = Array(MinD(box(0), other(0)), MinD(box(1), other(1)), MaxD(box(2), other(2)), MaxD(box(3), other(3)))
                        result.Remove index
                        merged = True
                    End If
                Next index
            Loop While merged
            result.Add box
        End If
NextWorkspaceLine:
    Next line
    Set WorkspaceFormatBlocks = result
End Function

' Binding-strip cells outside the inner frame must not connect a side stamp
' to the main title block and reserve a bounding rectangle across free paper.
Private Function WorkspaceInternalBox(ByVal line As Variant, ByVal frame As Variant) As Variant
    Dim box As Variant
    box = Array(MinD(line(0), line(2)), MinD(line(1), line(3)), MaxD(line(0), line(2)), MaxD(line(1), line(3)))
    If box(2) < frame(0) Or box(0) > frame(2) Or box(3) < frame(1) Or box(1) > frame(3) Then
        Exit Function
    End If
    WorkspaceInternalBox = Array(MaxD(frame(0), box(0)), MaxD(frame(1), box(1)), MinD(frame(2), box(2)), MinD(frame(3), box(3)))
End Function
