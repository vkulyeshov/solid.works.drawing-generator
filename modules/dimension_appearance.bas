Attribute VB_Name = "dimension_appearance"
Option Explicit

Private Function DrawingDimensionColor() As Long
    Select Case LCase$(Trim$(Setting("DimensionColor", "gray")))
        Case "gray"
            DrawingDimensionColor = RGB(DIMENSION_COLOR_GRAY_CHANNEL, DIMENSION_COLOR_GRAY_CHANNEL, DIMENSION_COLOR_GRAY_CHANNEL)
        Case "black"
            DrawingDimensionColor = DIMENSION_COLOR_BLACK
        Case Else
            Err.Raise vbObjectError + ERR_DRAWING_SETTINGS, , "DimensionColor must be gray or black."
    End Select
End Function

Private Sub ConfigureDrawingLineWeights(ByVal d As SldWorks.ModelDoc2)
    ' GOST 2.303: continuous narrow dimension lines are between s/3 and s/2.
    With d.Extension
        .SetUserPreferenceDouble SW_PREF_VISIBLE_EDGE_CUSTOM_THICKNESS, SW_PREFERENCE_NO_OPTION, DRAWING_CONTOUR_WIDTH_M
        .SetUserPreferenceInteger SW_PREF_VISIBLE_EDGE_THICKNESS, SW_PREFERENCE_NO_OPTION, SW_LINE_WEIGHT_CUSTOM
        .SetUserPreferenceDouble SW_PREF_PRINT_DIMENSION_WEIGHT, SW_PREFERENCE_NO_OPTION, DRAWING_DIMENSION_WIDTH_M
    End With
End Sub

Private Sub ApplyDrawingDimensionAppearance(ByVal d As SldWorks.ModelDoc2)
    Dim sheetName As Variant, view As SldWorks.View, display As SldWorks.DisplayDimension
    Dim annotation As SldWorks.Annotation, colorValue As Long, count As Long
    colorValue = DrawingDimensionColor()
    ConfigureDrawingLineWeights d
    For Each sheetName In AsDrawing(d).GetSheetNames
        AsDrawing(d).ActivateSheet CStr(sheetName)
        Set view = AsDrawing(d).GetFirstView.GetNextView
        Do While Not view Is Nothing
            Set display = view.GetFirstDisplayDimension5
            Do While Not display Is Nothing
                display.ExtensionLineUseDocumentDisplay = False
                display.SetLineFontDimensionThickness False, SW_LINE_WEIGHT_DIMENSION
                display.SetLineFontExtensionThickness False, SW_LINE_WEIGHT_DIMENSION
                Set annotation = display.GetAnnotation
                annotation.UseDocDispLeader = False
                annotation.LeaderThickness = SW_LINE_WEIGHT_DIMENSION
                annotation.Color = colorValue
                annotation.Width = SW_LINE_WEIGHT_DIMENSION
                annotation.LayerOverride = annotation.LayerOverride Or SW_LAYER_OVERRIDE_COLOR Or SW_LAYER_OVERRIDE_WIDTH
                count = count + 1
                Set display = display.GetNext5
            Loop
            Set view = view.GetNextView
        Loop
    Next sheetName
    d.GraphicsRedraw2
    Log "DIMENSION APPEARANCE count=" & count & "; color=" & Setting("DimensionColor", "gray") & "; contour mm=" & DRAWING_CONTOUR_WIDTH_M * MILLIMETRES_PER_METRE & "; dimension mm=" & DRAWING_DIMENSION_WIDTH_M * MILLIMETRES_PER_METRE
End Sub
