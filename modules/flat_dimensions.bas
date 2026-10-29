Attribute VB_Name = "flat_dimensions"
Option Explicit
Private Sub AddFlatDimensionSheet(ByVal d As SldWorks.ModelDoc2, ByVal doc As SldWorks.ModelDoc2, ByVal sourceView As SldWorks.View, ByVal code As String, ByVal name As String, ByVal material As String, ByVal mass As String)
    Dim v As SldWorks.View, detailScale As Double, sheet As SldWorks.Sheet, rc As Long, bounds As Variant, position As Variant
    If Setting("AutoDimensions", "1") <> "1" Then
        Exit Sub
    End If
    NewFormattedSheet d, "Розміри розгортки"
    Set v = AsDrawing(d).CreateDrawViewFromModelView3(doc.GetPathName, sourceView.GetOrientationName, LX(DETAIL_INITIAL_X_M), LY(DETAIL_INITIAL_Y_M), 0)
    If v Is Nothing Then
        Err.Raise vbObjectError + ERR_DETAIL_SHEET, , "Cannot create flat dimension view."
    End If
    v.ReferencedConfiguration = sourceView.ReferencedConfiguration
    v.UseSheetScale = False: v.ScaleDecimal = 1: v.Angle = sourceView.Angle
    d.ForceRebuild3 False
    bounds = ProjectedFlatBounds(v)
    v.ScaleDecimal = StandardScale(VIEW_FIT_SAFETY_FACTOR * MinD((sheetWidth - DETAIL_HORIZONTAL_RESERVED_M) / (bounds(OUTLINE_RIGHT_INDEX) - bounds(OUTLINE_LEFT_INDEX)), (sheetHeight - DETAIL_VERTICAL_RESERVED_M) / (bounds(OUTLINE_TOP_INDEX) - bounds(OUTLINE_BOTTOM_INDEX))))
    v.SetDisplayMode3 False, SW_DISPLAY_HIDDEN_LINES_REMOVED, False, False
    If FitFlatDimensionLayout(d, v) Then
        rc = OverallDimensions(d, v, True)
        DimensionFeatures d, v, True
    Else
        AddSameSheetDimensionDetails d, v
    End If
    Set sheet = AsDrawing(d).GetCurrentSheet
    detailScale = v.ScaleDecimal: SetPaperScale sheet, detailScale
    d.ForceRebuild3 False
    AsDrawing(d).ActivateView "": d.ClearSelection2 True
    RefreshTemplateFields d
    AddText d, "РОЗГОРТКА - КООРДИНАТНІ РОЗМІРИ ТА ОТВОРИ", CONTENT_LEFT_M, sheetHeight - SECTION_HEADER_TOP_MARGIN_M, TEXT_HEIGHT_SMALL_M, 0
End Sub

