Attribute VB_Name = "annotations"
Option Explicit
Private Sub AddText(ByVal d As SldWorks.ModelDoc2, ByVal text As String, ByVal x As Double, ByVal y As Double, ByVal height As Double, ByVal angle As Double)
    Dim note As Object, ann As Object, tf As Object
    If Len(text) = 0 Then
        Exit Sub
    End If
    d.ClearSelection2 True
    Set note = d.InsertNote(text)
    If note Is Nothing Then
        Err.Raise vbObjectError + ERR_INSERT_NOTE, , "Cannot insert note."
    End If
    Set ann = note.GetAnnotation
    Set tf = ann.GetTextFormat(SW_ANNOTATION_TEXT_FORMAT_FIRST)
    tf.TypeFaceName = Setting("Font", DEFAULT_FONT): tf.CharHeight = height
    tf.Italic = False: tf.Escapement = angle
    ann.SetTextFormat SW_ANNOTATION_TEXT_FORMAT_FIRST, False, tf
    ann.SetPosition2 x, y, 0
End Sub

Private Sub AddTechnicalNotes(ByVal d As SldWorks.ModelDoc2, ByVal doc As SldWorks.ModelDoc2, ByVal cfg As String, ByVal b As Object, ByVal qty As Long, ByVal assembly As Boolean)
    Dim text As String, notes As String, lines As Variant, firstSheet As String, chunk As String
    Dim i As Long, page As Long, perPage As Long, last As Long, row As Long
    text = "Кількість у виробі: " & qty & "."
    If Not b Is Nothing Then
        text = text & vbCrLf & "Тіло: " & b.Name
    End If
    notes = PropertyValue(doc, cfg, "Технические требования|Технічні вимоги|TechnicalRequirements")
    If notes <> "" Then
        text = text & vbCrLf & notes
    End If
    text = WrapText(text, CLng((sheetWidth - TECHNICAL_TEXT_HORIZONTAL_RESERVED_M) / TECHNICAL_TEXT_GLYPH_WIDTH_M))
    lines = Split(text, vbCrLf)
    If UBound(lines) < TECHNICAL_FIRST_SHEET_MAX_LINES Then
        AddText d, text, CONTENT_LEFT_M, TECHNICAL_FIRST_SHEET_Y_M, TECHNICAL_TEXT_HEIGHT_M, 0
        Exit Sub
    End If
    firstSheet = AsDrawing(d).GetCurrentSheet.GetName
    perPage = Int((sheetHeight - TECHNICAL_PAGE_RESERVED_HEIGHT_M) / TECHNICAL_LINE_HEIGHT_M)
    For i = 0 To UBound(lines) Step perPage
        page = page + 1: chunk = ""
        NewFormattedSheet d, "Технічні вимоги " & page
        last = CLng(MinD(UBound(lines), i + perPage - 1))
        For row = i To last
            If chunk <> "" Then
                chunk = chunk & vbCrLf
            End If
            chunk = chunk & CStr(lines(row))
        Next row
        AddText d, "ТЕХНІЧНІ ВИМОГИ", CONTENT_LEFT_M, sheetHeight - SECTION_HEADER_TOP_MARGIN_M, TECHNICAL_HEADER_HEIGHT_M, 0
        AddText d, chunk, CONTENT_LEFT_M, sheetHeight - TECHNICAL_TEXT_TOP_MARGIN_M, TECHNICAL_TEXT_HEIGHT_M, 0
    Next i
    AsDrawing(d).ActivateSheet firstSheet
End Sub

Private Sub AddAssemblyList(ByVal d As SldWorks.ModelDoc2, ByVal doc As SldWorks.ModelDoc2, ByVal cfg As String, ByVal code As String, ByVal name As String, ByVal mass As String)
    Dim v As Object, bom As Object, tf As Object, table As SldWorks.TableAnnotation
    Dim page As Long, firstRow As Long, lastRow As Long, row As Long, col As Long
    Dim height As Double, rowHeight As Double, width As Double, available As Double
    firstRow = 1: width = sheetWidth - BOM_HORIZONTAL_RESERVED_M: available = sheetHeight - BOM_VERTICAL_RESERVED_M
    Do
        page = page + 1
        NewFormattedSheet d, "Склад " & page
        Set v = MakeView(d, doc, cfg, Nothing, "*Isometric", CONTENT_LEFT_M, BOM_SOURCE_VIEW_Y_M)
        If v Is Nothing Then
            Err.Raise vbObjectError + ERR_NATIVE_BOM, , "Cannot create native BOM source view."
        End If
        v.UseSheetScale = False: v.ScaleDecimal = BOM_HIDDEN_VIEW_SCALE
        Set bom = v.InsertBomTable4(False, CONTENT_LEFT_M, sheetHeight - BOM_TOP_MARGIN_M, SW_BOM_ANCHOR_TOP_LEFT, SW_BOM_TOP_LEVEL_ONLY, cfg, "", False, 0, False)
        If bom Is Nothing Then
            Err.Raise vbObjectError + ERR_NATIVE_BOM, , "Native BOM insertion failed."
        End If
        Set table = bom
        Set tf = table.GetTextFormat
        tf.CharHeight = TEXT_HEIGHT_SMALL_M: tf.TypeFaceName = Setting("Font", DEFAULT_FONT)
        table.SetTextFormat False, tf
        For col = 0 To table.ColumnCount - 1
            If col = 0 Then
                table.SetColumnWidth col, BOM_POSITION_COLUMN_WIDTH_M, 0
            ElseIf col = table.ColumnCount - 1 Then
                table.SetColumnWidth col, BOM_QUANTITY_COLUMN_WIDTH_M, 0
            Else
                table.SetColumnWidth col, (width - BOM_FIXED_COLUMNS_WIDTH_M) / (table.ColumnCount - BOM_FIXED_COLUMN_COUNT), 0
            End If
        Next col
        table.Text2(0, 0, False) = "Поз."
        If table.ColumnCount = BOM_COMBINED_COLUMN_COUNT Then
            table.Text2(0, 1, False) = "Позначення / найменування"
        End If
        If table.ColumnCount = BOM_SPLIT_COLUMN_COUNT Then
            table.Text2(0, 1, False) = "Позначення"
            table.Text2(0, 2, False) = "Найменування"
        End If
        table.Text2(0, table.ColumnCount - 1, False) = "Кільк."
        For row = 0 To table.RowCount - 1
            table.SetRowHeight row, IIf(row = 0, BOM_HEADER_ROW_HEIGHT_M, BOM_DATA_ROW_HEIGHT_M), 0
        Next row
        d.ForceRebuild3 False
        height = table.GetRowHeight(0): lastRow = firstRow - 1
        For row = firstRow To table.RowCount - 1
            rowHeight = table.GetRowHeight(row)
            If height + rowHeight > available And lastRow >= firstRow Then
                Exit For
            End If
            If height + rowHeight > available Then
                Err.Raise vbObjectError + ERR_NATIVE_BOM, , "A BOM row is taller than the selected sheet."
            End If
            height = height + rowHeight: lastRow = row
        Next row
        For row = 1 To table.RowCount - 1
            table.RowHidden(row) = (row < firstRow Or row > lastRow)
        Next row
        v.SetVisible False, False
        AsDrawing(d).HiddenViewsVisible = False
        AddText d, "Склад виробу (таблиця компонентів)", CONTENT_LEFT_M, sheetHeight - SECTION_HEADER_TOP_MARGIN_M, TEXT_HEIGHT_STANDARD_M, 0
        Log "BOM page=" & page & " rows=" & firstRow & ".." & lastRow
        firstRow = lastRow + 1
    Loop While firstRow < table.RowCount
End Sub

