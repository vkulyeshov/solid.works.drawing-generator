Attribute VB_Name = "sheets"
Option Explicit
Private Sub ConfigureDrawing(ByVal d As SldWorks.ModelDoc2)
    Dim tf As Object, i As Long
    d.Extension.SetUserPreferenceInteger SW_PREF_DIMENSION_STANDARD, 0, SW_DRAFTING_STANDARD_GOST ' GOST
    d.Extension.SetUserPreferenceInteger SW_PREF_UNIT_SYSTEM, 0, SW_UNIT_SYSTEM_LEGACY_DEFAULT ' Keep the existing preset; linear units are set explicitly below.
    d.Extension.SetUserPreferenceInteger SW_PREF_LINEAR_UNITS, 0, SW_LENGTH_MILLIMETRES ' millimetres
    d.Extension.SetUserPreferenceInteger SW_PREF_LINEAR_PRECISION, 0, LINEAR_DECIMAL_PLACES
    d.Extension.SetUserPreferenceInteger SW_PREF_ANGULAR_PRECISION, 0, ANGULAR_DECIMAL_PLACES
    d.Extension.SetUserPreferenceDouble SW_PREF_ARROW_HEAD_LENGTH, SW_PREFERENCE_NO_OPTION, GOST_ARROW_HEAD_LENGTH_M
    d.Extension.SetUserPreferenceDouble SW_PREF_ARROW_HEAD_WIDTH, SW_PREFERENCE_NO_OPTION, GOST_ARROW_HEAD_WIDTH_M
    d.Extension.SetUserPreferenceDouble SW_PREF_ARROW_LENGTH, SW_PREFERENCE_NO_OPTION, GOST_ARROW_LENGTH_M
    d.Extension.SetUserPreferenceDouble SW_PREF_WITNESS_EXTENSION, SW_PREFERENCE_NO_OPTION, GOST_WITNESS_EXTENSION_M
    d.Extension.SetUserPreferenceDouble SW_PREF_WITNESS_GAP, SW_PREFERENCE_NO_OPTION, GOST_WITNESS_GAP_M
    d.Extension.SetUserPreferenceDouble SW_PREF_DIMENSION_ROW_OFFSET, SW_PREFERENCE_NO_OPTION, COORDINATE_LANE_PITCH_M
    d.Extension.SetUserPreferenceDouble SW_PREF_DIMENSION_OBJECT_OFFSET, SW_PREFERENCE_NO_OPTION, GOST_OBJECT_OFFSET_M
    d.Extension.SetUserPreferenceInteger SW_PREF_DIMENSION_ARROW_STYLE, SW_PREFERENCE_NO_OPTION, SW_ARROW_CLOSED
    d.Extension.SetUserPreferenceInteger SW_PREF_DIMENSION_ARROW_POSITION, SW_PREFERENCE_NO_OPTION, SW_DIMENSION_ARROWS_SMART
    d.Extension.SetUserPreferenceInteger SW_PREF_DIMENSION_TRAILING_ZEROES, SW_PREFERENCE_NO_OPTION, SW_REMOVE_TRAILING_ZEROES
    d.Extension.SetUserPreferenceToggle SW_PREF_SCALE_ARROWS_WITH_TEXT, SW_PREFERENCE_NO_OPTION, False
    Dim toggleId As Variant
    For Each toggleId In Array(SW_HIDE_AXES, SW_HIDE_PLANES, SW_HIDE_ORIGINS, SW_HIDE_TEMPORARY_AXES, SW_HIDE_COORDINATE_SYSTEMS, SW_HIDE_REFERENCE_POINTS, SW_HIDE_CURVES, SW_HIDE_SKETCHES, SW_HIDE_REFERENCE_POINTS_2)
        d.SetUserPreferenceToggle CLng(toggleId), False
    Next toggleId
    For i = SW_TEXT_FORMAT_NOTE To SW_TEXT_FORMAT_DIMENSION
        Set tf = d.GetUserPreferenceTextFormat(i)
        tf.TypeFaceName = Setting("Font", DEFAULT_FONT)
        tf.CharHeight = TEXT_HEIGHT_STANDARD_M: tf.Italic = False
        d.SetUserPreferenceTextFormat i, tf
    Next i
End Sub

Private Sub ResolveDrawingTemplate()
    Dim base As String, folder As String
    base = Setting("LauncherSettingsDirectory", sw.GetCurrentMacroPathFolder)
    folder = Setting("TemplatesDirectory", Setting("LauncherScriptDirectory", sw.GetCurrentMacroPathFolder) & "\templates")
    If InStr(folder, ":") = 0 And Left$(folder, 2) <> "\\" Then
        folder = fs.BuildPath(base, folder)
    End If
    paperFormat = UCase$(Setting("PaperFormat", PAPER_A3))
    paperOrientation = LCase$(Setting("PaperOrientation", ORIENTATION_LANDSCAPE))
    If paperFormat <> PAPER_A4 And paperFormat <> PAPER_A3 Then
        Err.Raise vbObjectError + ERR_DRAWING_SETTINGS, , "PaperFormat must be A4 or A3."
    End If
    If paperOrientation <> ORIENTATION_LANDSCAPE And paperOrientation <> ORIENTATION_PORTRAIT Then
        Err.Raise vbObjectError + ERR_DRAWING_SETTINGS, , "PaperOrientation must be landscape or portrait."
    End If
    templatePath = Setting("DrawingTemplate", "")
    If templatePath = "" Then
        templatePath = fs.BuildPath(folder, paperFormat & "." & paperOrientation & DRAWING_TEMPLATE_EXTENSION)
    End If
    If InStr(templatePath, ":") = 0 And Left$(templatePath, 2) <> "\\" Then
        templatePath = fs.BuildPath(base, templatePath)
    End If
    templatePath = fs.GetAbsolutePathName(templatePath)
    If LCase$(fs.GetExtensionName(templatePath)) <> DRAWING_TEMPLATE_FILE_TYPE Then
        Err.Raise vbObjectError + ERR_DRAWING_SETTINGS, , "DrawingTemplate must be a .drwdot file."
    End If
    If Not fs.FileExists(templatePath) Then
        Err.Raise vbObjectError + ERR_DRAWING_SETTINGS, , "Drawing template not found: " & templatePath
    End If
End Sub

Private Sub ValidateDrawingTemplate()
    Dim candidate As SldWorks.ModelDoc2, errorNumber As Long, errorText As String
    On Error GoTo Failed
    Set candidate = sw.NewDocument(templatePath, 0, 0, 0)
    If candidate Is Nothing Then
        Err.Raise vbObjectError + ERR_DRAWING_TEMPLATE, , "Cannot open selected drawing template."
    End If
    InitializePaper candidate, False
    sw.CloseDoc candidate.GetTitle
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    If Not candidate Is Nothing Then
        sw.CloseDoc candidate.GetTitle
    End If
    On Error GoTo 0
    Err.Raise errorNumber, , errorText
End Sub

Private Sub InitializePaper(ByVal d As SldWorks.ModelDoc2, Optional ByVal saveFormat As Boolean = True)
    Dim sheet As SldWorks.Sheet, properties As Variant, shortSide As Double, longSide As Double
    Dim active As SldWorks.ModelDoc2, error As Long, annotations As Variant, noteCount As Long, formatResult As Long
    PrepareDocumentWindow d
    Set active = sw.ActivateDoc3(d.GetTitle, False, SW_REBUILD_USER_DECISION, error)
    If AsDrawing(d).GetSheetCount <> 1 Then
        Err.Raise vbObjectError + ERR_DRAWING_TEMPLATE, , "Drawing template must contain one sheet."
    End If
    Set sheet = AsDrawing(d).GetCurrentSheet
    properties = sheet.GetProperties
    sheetWidth = CDbl(properties(SHEET_WIDTH_PROPERTY)): sheetHeight = CDbl(properties(SHEET_HEIGHT_PROPERTY))
    shortSide = MinD(sheetWidth, sheetHeight): longSide = -MinD(-sheetWidth, -sheetHeight)
    If paperFormat = PAPER_A4 Then
        If Abs(shortSide - A4_SHORT_SIDE_M) > SHEET_SIZE_TOLERANCE_M Or Abs(longSide - A4_LONG_SIDE_AND_A3_SHORT_SIDE_M) > SHEET_SIZE_TOLERANCE_M Then
            Err.Raise vbObjectError + ERR_DRAWING_TEMPLATE, , "Selected template size does not match A4 (210 x 297 mm)."
        End If
    Else
        If Abs(shortSide - A4_LONG_SIDE_AND_A3_SHORT_SIDE_M) > SHEET_SIZE_TOLERANCE_M Or Abs(longSide - A3_LONG_SIDE_M) > SHEET_SIZE_TOLERANCE_M Then
            Err.Raise vbObjectError + ERR_DRAWING_TEMPLATE, , "Selected template size does not match A3 (297 x 420 mm)."
        End If
    End If
    layoutSx = (sheetWidth - LAYOUT_HORIZONTAL_RESERVED_M) / REFERENCE_USABLE_WIDTH_M
    layoutSy = (sheetHeight - LAYOUT_VERTICAL_RESERVED_M) / REFERENCE_USABLE_HEIGHT_M
    ' Export the EMBEDDED format. Renamed/copied .drwdot files need no old external .slddrt.
    If saveFormat Then
        AsDrawing(d).EditTemplate
        AsDrawing(d).EditSheet
        d.ForceRebuild3 False
        annotations = AsDrawing(d).GetFirstView.GetAnnotations
        If Not IsEmpty(annotations) Then
            noteCount = UBound(annotations) + 1
        End If
        Log "TEMPLATE path=" & templatePath & "; notes=" & noteCount
        formatResult = sheet.SaveFormat(runtimeFormatPath)
        Log "TEMPLATE SaveFormat result=" & formatResult
        If formatResult = 0 Then
            Err.Raise vbObjectError + ERR_DRAWING_TEMPLATE, , "Cannot save embedded drawing sheet format."
        End If
        sheet.SetTemplateName runtimeFormatPath
    End If
    Log "PAPER " & paperFormat & " " & Format$(sheetWidth * MILLIMETRES_PER_METRE, "0") & "x" & Format$(sheetHeight * MILLIMETRES_PER_METRE, "0") & " mm | " & templatePath
End Sub

Private Function LX(ByVal reference As Double) As Double
    LX = LAYOUT_ORIGIN_X_M + (reference - LAYOUT_ORIGIN_X_M) * layoutSx
End Function

Private Function LY(ByVal reference As Double) As Double
    LY = LAYOUT_ORIGIN_Y_M + (reference - LAYOUT_ORIGIN_Y_M) * layoutSy
End Function

Private Function FitWidth(ByVal reference As Double) As Double
    FitWidth = reference * layoutSx - SMALL_SHEET_WIDTH_CLEARANCE_M * (1 - MinD(1, layoutSx))
End Function

Private Function FitHeight(ByVal reference As Double) As Double
    FitHeight = reference * layoutSy - SMALL_SHEET_HEIGHT_CLEARANCE_M * (1 - MinD(1, layoutSy))
End Function

Private Sub NewFormattedSheet(ByVal d As SldWorks.ModelDoc2, ByVal name As String)
    Dim sheet As SldWorks.Sheet
    If AsDrawing(d).NewSheet4(name, SW_CUSTOM_PAPER, SW_CUSTOM_SHEET_TEMPLATE, 1, 1, True, runtimeFormatPath, sheetWidth, sheetHeight, "", 0, 0, 0, 0, 0, 0) = False Then
        Err.Raise vbObjectError + ERR_DETAIL_SHEET, , "Cannot create formatted sheet: " & name
    End If
    AsDrawing(d).ActivateSheet name
    If AsDrawing(d).SetupSheet5(name, SW_CUSTOM_PAPER, SW_CUSTOM_SHEET_TEMPLATE, 1, 1, True, runtimeFormatPath, sheetWidth, sheetHeight, "", True) = False Then
        Err.Raise vbObjectError + ERR_DETAIL_SHEET, , "Cannot apply sheet format: " & name
    End If
    Set sheet = AsDrawing(d).GetCurrentSheet
    sheet.SetTemplateName runtimeFormatPath
    If sheet.ReloadTemplate(False) <> 0 Then
        Err.Raise vbObjectError + ERR_DETAIL_SHEET, , "Cannot reload sheet format: " & name
    End If
    d.ForceRebuild3 False
End Sub

Private Sub SetPaperScale(ByVal sheet As Object, ByVal scaleValue As Double)
    If scaleValue >= 1 Then
        sheet.SetScale scaleValue, 1, False, False
    Else
        sheet.SetScale 1, 1 / scaleValue, False, False
    End If
End Sub

Private Sub SetDrawingProperty(ByVal d As SldWorks.ModelDoc2, ByVal name As String, ByVal value As String)
    d.Extension.CustomPropertyManager("").Add3 name, SW_CUSTOM_PROPERTY_TEXT, value, SW_PROPERTY_REPLACE_VALUE
End Sub

Private Sub FillTemplateProperties(ByVal d As SldWorks.ModelDoc2, ByVal code As String, ByVal name As String, ByVal material As String, ByVal mass As String)
    Dim key As Variant
    SetDrawingProperty d, "Designation", code
    SetDrawingProperty d, "Description", WrapText(name, DESCRIPTION_WRAP_CHARACTERS)
    SetDrawingProperty d, "Material", WrapText(material, MATERIAL_WRAP_CHARACTERS)
    SetDrawingProperty d, "Mass", mass
    If Setting("Organization", "") <> "" Then
        SetDrawingProperty d, "Organization", Setting("Organization", "")
    End If
    For Each key In Array("Liter1", "Liter2", "Liter3", "DrawnBy", "DrawnDate", "CheckedBy", "CheckedDate", "TechnicalControlBy", "TechnicalControlDate", "NormControlBy", "NormControlDate", "ApprovedBy", "ApprovedDate")
        If settings.Exists(CStr(key)) Then
            SetDrawingProperty d, CStr(key), Setting(CStr(key), "")
        End If
    Next key
    d.ForceRebuild3 False
End Sub

Private Sub RefreshTemplateFields(ByVal d As SldWorks.ModelDoc2)
    Dim sheets As Variant, name As Variant, view As SldWorks.View, annotations As Variant, item As Variant
    Dim ann As SldWorks.Annotation, note As SldWorks.Note, tf As SldWorks.TextFormat, box As Variant
    Dim w As Double, h As Double, actualW As Double, actualH As Double, attempt As Long
    sheets = AsDrawing(d).GetSheetNames
    For Each name In sheets
        AsDrawing(d).ActivateSheet CStr(name)
        Set view = AsDrawing(d).GetFirstView
        annotations = view.GetAnnotations
        If Not IsEmpty(annotations) Then
            For Each item In annotations
                Set ann = item
                If ann.GetType = SW_ANNOTATION_NOTE Then
                    Set note = ann.GetSpecificAnnotation
                    If note.GetName = "DSTU_SheetNumber" Then
                        If UBound(sheets) = 0 Then
                            note.SetText "Аркуш "
                        Else
                            note.SetText "Аркуш $PRP:""SW-Current Sheet"""
                        End If
                    End If
                    w = 0: h = 0
                    Select Case note.GetName
                        Case "DSTU_Designation": w = TITLE_DESIGNATION_WIDTH_M: h = TITLE_TEXT_CELL_HEIGHT_M
                        Case "DSTU_Designation26": w = TITLE_MAIN_CELL_WIDTH_M: h = TITLE_AUX_CODE_HEIGHT_M
                        Case "DSTU_Description": w = TITLE_MAIN_CELL_WIDTH_M: h = TITLE_DESCRIPTION_HEIGHT_M
                        Case "DSTU_Material": w = TITLE_MAIN_CELL_WIDTH_M: h = TITLE_TEXT_CELL_HEIGHT_M
                        Case "DSTU_Organization": w = TITLE_ORGANIZATION_WIDTH_M: h = TITLE_TEXT_CELL_HEIGHT_M
                    End Select
                    If w > 0 And Len(note.GetText) > 0 Then
                        For attempt = 0 To TITLE_LAST_FIT_PASS
                            box = note.GetExtent
                            actualW = box(BBOX_X_MAX_INDEX) - box(BBOX_X_MIN_INDEX): actualH = box(BBOX_Y_MAX_INDEX) - box(BBOX_Y_MIN_INDEX)
                            If Abs(Sin(note.Angle)) > VERTICAL_NOTE_SINE_THRESHOLD Then
                                actualW = box(BBOX_Y_MAX_INDEX) - box(BBOX_Y_MIN_INDEX)
                                actualH = box(BBOX_X_MAX_INDEX) - box(BBOX_X_MIN_INDEX)
                            End If
                            If actualW <= w And actualH <= h Then
                                Exit For
                            End If
                            Set tf = ann.GetTextFormat(SW_ANNOTATION_TEXT_FORMAT_FIRST)
                            If tf.CharHeight > TITLE_MEDIUM_FONT_THRESHOLD_M Then
                                tf.CharHeight = TEXT_HEIGHT_STANDARD_M
                            ElseIf tf.CharHeight > TITLE_SMALL_FONT_THRESHOLD_M Then
                                tf.CharHeight = TEXT_HEIGHT_SMALL_M
                            Else
                                tf.WidthFactor = tf.WidthFactor * MinD(1, w / actualW)
                            End If
                            ann.SetTextFormat SW_ANNOTATION_TEXT_FORMAT_FIRST, False, tf
                        Next attempt
                    End If
                End If
            Next item
        End If
    Next name
End Sub

