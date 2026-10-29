Attribute VB_Name = "jobs"
Option Explicit
Private Sub ProcessJob(ByVal j As Object)
    Dim doc As SldWorks.ModelDoc2, d As SldWorks.ModelDoc2, mainView As Object, bottomView As Object, sideView As Object, isoView As Object, flatView As SldWorks.View
    Dim b As Object, cfg As String, oldCfg As String, name As String, code As String, material As String
    Dim outName As String, best As String, dims As Long, rc As Long, er As Long, wr As Long, success As Boolean
    Dim sheet As Object, viewScaleValue As Double, flatScale As Double, bodies As Variant, oneBody(0) As Object
    Dim mark As String, bodyName As String, mass As String, isAssembly As Boolean, multi As Boolean, active As Object
    Dim useNativeDimensions As Boolean
    Dim firstSheetName As String
    On Error GoTo Failed
    Set dimensionKeys = CreateObject("Scripting.Dictionary")
    outName = CStr(j("out"))
    Log "JOB " & outName
    Dim previousDrawing As SldWorks.ModelDoc2
    Set previousDrawing = sw.GetOpenDocumentByName(runDir & DRAWINGS_RELATIVE_DIRECTORY & outName & DRAWING_EXTENSION)
    If Not previousDrawing Is Nothing Then
        sw.CloseDoc previousDrawing.GetTitle
    End If
    Set previousDrawing = Nothing
    cfg = CStr(j("cfg")): bodyName = CStr(j("body"))
    Set doc = sw.OpenDoc6(CStr(j("path")), CLng(j("type")), SW_OPEN_SILENT, cfg, er, wr)
    If doc Is Nothing Then
        Err.Raise vbObjectError + ERR_OPEN_JOB_MODEL, , "Cannot open copied model for drawing: " & er
    End If
    PrepareDocumentWindow doc
    oldCfg = doc.ConfigurationManager.ActiveConfiguration.Name
    doc.ShowConfiguration2 cfg
    isAssembly = (doc.GetType = SW_DOC_ASSEMBLY)
    useNativeDimensions = Setting("NativeDimensions", "0") = "1" And Not isAssembly And Not CBool(j("sheet"))
    If isAssembly Then
        AsAssembly(doc).ResolveAllLightWeightComponents False
    End If
    name = PropertyValue(doc, cfg, "Наименование|Найменування|Description|SW-Title")
    If name = "" Then
        name = OriginalBase(doc.GetPathName)
    End If
    code = PropertyValue(doc, cfg, "Обозначение|Позначення|Designation|PartNumber|Number")
    If code = "" Then
        code = OriginalBase(doc.GetPathName)
    End If
    multi = False
    If Not isAssembly Then
        bodies = AsPart(doc).GetBodies2(SW_SOLID_BODY, True)
        multi = UBound(bodies) > LBound(bodies)
        Set b = FindBody(doc, bodyName)
        If b Is Nothing Then
            Err.Raise vbObjectError + ERR_BODY_MISSING, , "Body not found: " & bodyName
        End If
        If multi Then
            name = name & " / " & IIf(CStr(j("cut")) = "", bodyName, CStr(j("cut")))
            code = code & "." & Format$(BodyOrdinal(doc, bodyName), JOB_NUMBER_FORMAT)
        End If
    Else
        code = code & " СБ"
    End If
    material = PropertyValue(doc, cfg, "Материал|Матеріал|Material")
    If material = "" And Not isAssembly Then
        Dim matDb As String
        material = AsPart(doc).GetMaterialPropertyName2(cfg, matDb)
    End If
    mass = MassText(doc, b)
    JobStage j, "Creating drawing"
    Set d = sw.NewDocument(templatePath, 0, 0, 0)
    If d Is Nothing Then
        Err.Raise vbObjectError + ERR_NEW_DRAWING, , "Cannot create drawing."
    End If
    HideBackgroundFrame
    Set currentDrawing = d
    Set sheet = AsDrawing(d).GetCurrentSheet
    firstSheetName = sheet.GetName
    InitializePaper d
    FillTemplateProperties d, code, name, material, mass
    ConfigureDrawing d
    best = BestOrientation(doc, b)
    Set mainView = MakeView(d, doc, cfg, b, best, LX(ORTHOGRAPHIC_COLUMN_X_M), LY(ORTHOGRAPHIC_ROW_Y_M))
    If mainView Is Nothing Then
        Err.Raise vbObjectError + ERR_MAIN_VIEW, , "Main view failed."
    End If
    mainView.UseSheetScale = SW_USE_INDEPENDENT_SCALE
    mainView.ScaleDecimal = 1
    d.ForceRebuild3 False
    Dim box As Variant
    box = mainView.GetOutline
    If Not isAssembly Then
        If CBool(j("sheet")) And box(OUTLINE_TOP_INDEX) - box(1) > box(2) - box(0) Then
            mainView.Angle = PI / 2
        End If
    End If
    If Not CBool(j("sheet")) Then
        UprightView mainView
    End If
    d.ForceRebuild3 False
    viewScaleValue = FitScale(mainView, FitWidth(MAIN_BOX_WIDTH_M), FitHeight(ORTHOGRAPHIC_BOX_HEIGHT_M))
    SetPaperScale sheet, viewScaleValue
    mainView.UseSheetScale = SW_USE_SHEET_SCALE
    CenterView mainView, LX(ORTHOGRAPHIC_COLUMN_X_M), LY(ORTHOGRAPHIC_ROW_Y_M)
    d.ClearSelection2 True
    AsDrawing(d).ActivateView ""
    d.Extension.SelectByID2 mainView.Name, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0
    Set bottomView = AsDrawing(d).CreateUnfoldedViewAt3(LX(ORTHOGRAPHIC_COLUMN_X_M), LY(AUXILIARY_ROW_Y_M), 0, False)
    If Not bottomView Is Nothing Then
        ApplyBody bottomView, b
        bottomView.AlignWithView SW_NO_VIEW_ALIGNMENT, Nothing
        bottomView.UseParentScale = True
    End If
    d.ClearSelection2 True
    d.Extension.SelectByID2 mainView.Name, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0
    Set sideView = AsDrawing(d).CreateUnfoldedViewAt3(LX(AUXILIARY_COLUMN_X_M), LY(ORTHOGRAPHIC_ROW_Y_M), 0, False)
    If Not sideView Is Nothing Then
        ApplyBody sideView, b
        sideView.AlignWithView SW_NO_VIEW_ALIGNMENT, Nothing
        sideView.UseParentScale = True
    End If
    If Not CBool(j("sheet")) Then
        If Not bottomView Is Nothing Then
            UprightView bottomView
        End If
        If Not sideView Is Nothing Then
            UprightView sideView
        End If
    End If
    d.ForceRebuild3 False
    ' Reduce all orthographic views together if a projected view is too large.
    Dim reduction As Double
    reduction = 1
    If Not bottomView Is Nothing Then
        reduction = MinD(reduction, FitFactor(bottomView, FitWidth(MAIN_BOX_WIDTH_M), FitHeight(BOTTOM_BOX_HEIGHT_M)))
    End If
    If Not sideView Is Nothing Then
        reduction = MinD(reduction, FitFactor(sideView, FitWidth(SIDE_BOX_WIDTH_M), FitHeight(ORTHOGRAPHIC_BOX_HEIGHT_M)))
    End If
    If reduction < SCALE_REDUCTION_THRESHOLD Then
        viewScaleValue = StandardScale(viewScaleValue * reduction)
        SetPaperScale sheet, viewScaleValue
        d.ForceRebuild3 False
    End If
    CenterView mainView, LX(ORTHOGRAPHIC_COLUMN_X_M), LY(ORTHOGRAPHIC_ROW_Y_M)
    If Not bottomView Is Nothing Then
        CenterView bottomView, LX(ORTHOGRAPHIC_COLUMN_X_M), LY(AUXILIARY_ROW_Y_M)
    End If
    If Not sideView Is Nothing Then
        CenterView sideView, LX(AUXILIARY_COLUMN_X_M), LY(ORTHOGRAPHIC_ROW_Y_M)
    End If
    If CBool(j("sheet")) Then
        Set flatView = MakeFlatView(d, doc, cfg, b)
        If Not flatView Is Nothing Then
            flatView.UseSheetScale = False
            flatView.ScaleDecimal = INITIAL_FLAT_VIEW_SCALE
        Else
            Warn "Flat pattern not created for " & bodyName
        End If
    Else
        Set isoView = MakeView(d, doc, cfg, b, "*Isometric", LX(AUXILIARY_COLUMN_X_M), LY(AUXILIARY_ROW_Y_M))
        If Not isoView Is Nothing Then
            isoView.UseSheetScale = SW_USE_INDEPENDENT_SCALE: isoView.ScaleDecimal = 1
            d.ForceRebuild3 False
            isoView.ScaleDecimal = FitScale(isoView, FitWidth(AUXILIARY_BOX_WIDTH_M), FitHeight(ISOMETRIC_BOX_HEIGHT_M))
            CenterView isoView, LX(AUXILIARY_COLUMN_X_M), LY(AUXILIARY_ROW_Y_M)
        End If
    End If
    d.ForceRebuild3 False
    ' Flat view insertion can change the sheet scale; restore the calculated layout.
    SetPaperScale sheet, viewScaleValue
    mainView.UseSheetScale = SW_USE_SHEET_SCALE
    If Not bottomView Is Nothing Then
        bottomView.UseParentScale = True
    End If
    If Not sideView Is Nothing Then
        sideView.UseParentScale = True
    End If
    d.ForceRebuild3 False
    CenterView mainView, LX(ORTHOGRAPHIC_COLUMN_X_M), LY(ORTHOGRAPHIC_ROW_Y_M)
    If Not bottomView Is Nothing Then
        CenterView bottomView, LX(ORTHOGRAPHIC_COLUMN_X_M), LY(AUXILIARY_ROW_Y_M)
    End If
    If Not sideView Is Nothing Then
        CenterView sideView, LX(AUXILIARY_COLUMN_X_M), LY(ORTHOGRAPHIC_ROW_Y_M)
    End If
    d.GraphicsRedraw2
    box = mainView.Position
    Log "VIEW before dimensions scale=" & mainView.ScaleDecimal & " position=" & box(0)
    JobStage j, "Dimensioning folded views"
    If useNativeDimensions Then
        dims = NativeDimensionView(d, mainView)
        If Not bottomView Is Nothing Then
            rc = NativeDimensionView(d, bottomView)
        End If
        If Not sideView Is Nothing Then
            rc = NativeDimensionView(d, sideView)
        End If
    Else
        If Not isAssembly Then
            ImportModelItems d
        End If
        dims = OverallDimensions(d, mainView)
        If Not bottomView Is Nothing Then
            rc = OverallDimensions(d, bottomView)
        End If
        If Not sideView Is Nothing Then
            rc = OverallDimensions(d, sideView)
        End If
        If Not isAssembly Then
            ' Keep graphical feature dimensions; suppress only matching model-coordinate intervals.
            If Not CBool(j("sheet")) Then
                DimensionFeatures d, mainView, False
                If Not bottomView Is Nothing Then
                    DimensionFeatures d, bottomView, False
                End If
                If Not sideView Is Nothing Then
                    DimensionFeatures d, sideView, False
                End If
            ElseIf flatView Is Nothing Then
                rc = DiameterDimensions(d, mainView): rc = RadiusDimensions(d, mainView)
                If Not bottomView Is Nothing Then
                    rc = DiameterDimensions(d, bottomView)
                End If
                If Not sideView Is Nothing Then
                    rc = DiameterDimensions(d, sideView)
                End If
            End If
        End If
    End If
    box = mainView.Position
    Log "VIEW after dimensions scale=" & mainView.ScaleDecimal & " position=" & box(0)
    AsDrawing(d).ActivateView "": d.ClearSelection2 True
    RefreshTemplateFields d
    If Not flatView Is Nothing Then
        AddText d, "Розміри розгортки - на окремому аркуші.", CONTENT_LEFT_M, FLAT_NOTICE_Y_M, TEXT_HEIGHT_SMALL_M, 0
    End If
    If Not isoView Is Nothing Then
        AddText d, "Ізометрія (" & ScaleText(isoView.ScaleDecimal) & ")", LX(AUXILIARY_CAPTION_X_M), AUXILIARY_CAPTION_Y_M, TEXT_HEIGHT_SMALL_M, 0
    End If
    AddTechnicalNotes d, doc, cfg, b, CLng(j("qty")), isAssembly
    If isAssembly Then
        AddAssemblyList d, doc, cfg, code, name, mass
    End If
    NormalizeDimensions d
    AsDrawing(d).ActivateSheet sheet.GetName
    d.ActiveView.EnableGraphicsUpdate = True
    d.ForceRebuild3 False
    d.ViewZoomtofit2
    ' Persist derived flat configurations in the snapshot only.
    doc.ShowConfiguration2 oldCfg
    If doc.Save3(SW_SAVE_SILENT, er, wr) = False Then
        Err.Raise vbObjectError + ERR_SAVE_MODEL, , "Cannot save snapshot model: " & er
    End If
    success = d.Extension.SaveAs(runDir & DRAWINGS_RELATIVE_DIRECTORY & outName & DRAWING_EXTENSION, SW_SAVE_CURRENT_VERSION, SW_SAVE_SILENT, Nothing, er, wr)
    If success = False Then
        Err.Raise vbObjectError + ERR_SAVE_DRAWING, , "Drawing save failed: " & er
    End If
    ' Reopen to refresh cached drawing display data before PDF export.
    dimensionKeys.RemoveAll
    Set coordinateSelection = Nothing
    Set coordinateSourceView = Nothing
    Set coordinateAssignedRequests = Nothing
    coordinateRegion = Empty
    sw.CloseDoc d.GetTitle
    Set d = sw.OpenDoc6(runDir & DRAWINGS_RELATIVE_DIRECTORY & outName & DRAWING_EXTENSION, SW_DOC_DRAWING, SW_OPEN_SILENT, "", er, wr)
    If d Is Nothing Then
        Err.Raise vbObjectError + ERR_REOPEN_DRAWING, , "Cannot reopen saved drawing: " & er
    End If
    Set currentDrawing = d
    If backgroundOwned Then
        PrepareDocumentWindow d
    End If
    d.ForceRebuild3 False
    Set flatView = AsDrawing(d).GetFirstView
    Set flatView = flatView.GetNextView
    Do While Not flatView Is Nothing
        If flatView.IsFlatPatternView Then
            JobStage j, "Building and dimensioning flat pattern"
            flatScale = LayoutFlatView(d, doc, flatView)
            Set active = sw.ActivateDoc3(d.GetTitle, False, SW_REBUILD_USER_DECISION, er)
            AddText d, "Розгортка (" & ScaleText(flatScale) & ")", LX(AUXILIARY_CAPTION_X_M), AUXILIARY_CAPTION_Y_M, TEXT_HEIGHT_SMALL_M, 0
            AddFlatDimensionSheet d, doc, flatView, code, name, material, mass
            AsDrawing(d).ActivateSheet firstSheetName
            Exit Do
        End If
        Set flatView = flatView.GetNextView
    Loop
    d.ForceRebuild3 False
    d.GraphicsRedraw2
    RefreshTemplateFields d
    JobStage j, "Saving rebuilt drawing"
    NormalizeDimensions d
    If useNativeDimensions Then
        JobStage j, "Arranging native dimensions and inclined leaders"
        FinishNativeDimensionLayout d
    End If
    WriteDimensionCoverage d, outName
    AsDrawing(d).ActivateSheet firstSheetName
    Log "SAVE rebuilt drawing " & outName
    If d.Save3(SW_SAVE_SILENT, er, wr) = False Then
        Err.Raise vbObjectError + ERR_SAVE_DRAWING, , "Rebuilt drawing save failed: " & er
    End If
    Dim pdf As Object
    Set pdf = sw.GetExportFileData(SW_EXPORT_PDF)
    pdf.ViewPdfAfterSaving = False
    pdf.SetSheets SW_EXPORT_ALL_SHEETS, Empty
    JobStage j, "Exporting PDF"
    Log "EXPORT PDF " & outName
    success = d.Extension.SaveAs(runDir & PDF_RELATIVE_DIRECTORY & outName & PDF_EXTENSION, SW_SAVE_CURRENT_VERSION, SW_SAVE_SILENT, pdf, er, wr)
    If success = False Then
        Err.Raise vbObjectError + ERR_SAVE_PDF, , "PDF save failed: " & er
    End If
    AppendManifest jobNumber & vbTab & OriginalBase(doc.GetPathName) & vbTab & cfg & vbTab & bodyName & vbTab & j("qty") & vbTab & outName & DRAWING_EXTENSION & vbTab & "OK; REVIEW"
    j("status") = STATE_DONE: j("stage") = "Saved"
    j("drawingBytes") = CStr(fs.GetFile(runDir & DRAWINGS_RELATIVE_DIRECTORY & outName & DRAWING_EXTENSION).Size)
    j("pdfBytes") = CStr(fs.GetFile(runDir & PDF_RELATIVE_DIRECTORY & outName & PDF_EXTENSION).Size)
    RecountJobResults
    SaveCheckpoint
    Log "OK " & outName & " dims(main)=" & dims
    dimensionKeys.RemoveAll
    Set coordinateSelection = Nothing
    Set coordinateSourceView = Nothing
    Set coordinateAssignedRequests = Nothing
    coordinateRegion = Empty
    sw.CloseDoc d.GetTitle
    Set currentDrawing = Nothing
    Exit Sub
Failed:
    Dim msg As String
    msg = Err.Number & " " & Err.Description
    batchFailures = batchFailures + 1
    On Error Resume Next
    j("status") = STATE_FAILED: j("stage") = msg
    RecountJobResults
    SaveCheckpoint
    Log "FAIL " & outName & " | " & msg
    AppendManifest jobNumber & vbTab & OriginalBase(CStr(j("path"))) & vbTab & cfg & vbTab & bodyName & vbTab & j("qty") & vbTab & outName & vbTab & "FAILED " & msg
    On Error Resume Next
    If Not doc Is Nothing Then
        doc.ShowConfiguration2 oldCfg
    End If
    If Not d Is Nothing Then
        dimensionKeys.RemoveAll
    Set coordinateSelection = Nothing
    Set coordinateSourceView = Nothing
    Set coordinateAssignedRequests = Nothing
    coordinateRegion = Empty
        sw.CloseDoc d.GetTitle
    End If
    Set currentDrawing = Nothing
End Sub

