Attribute VB_Name = "native_dimensions"
Option Explicit
Private Function NativeDimensionView(ByVal d As SldWorks.ModelDoc2, ByVal v As SldWorks.View) As Long
    Dim status As Long, display As SldWorks.DisplayDimension
    If Setting("AutoDimensions", "1") <> "1" Then
        Exit Function
    End If
    d.ClearSelection2 True
    AsDrawing(d).ActivateView v.Name
    If Not d.Extension.SelectByID2(v.Name, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0) Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot select a view for native dimensioning: " & v.Name
    End If
    status = AsDrawing(d).AutoDimension(SW_AUTODIM_ALL_ENTITIES, SW_AUTODIM_CHAIN, SW_AUTODIM_ABOVE, SW_AUTODIM_CHAIN, SW_AUTODIM_LEFT)
    d.ClearSelection2 True
    Set display = v.GetFirstDisplayDimension5
    Do While Not display Is Nothing
        NativeDimensionView = NativeDimensionView + 1
        Set display = display.GetNext5
    Loop
    Log "NATIVE DIMENSIONS " & v.Name & "; status=" & status & "; dimensions=" & NativeDimensionView
    If NativeDimensionView = 0 Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Native dimensioning produced no dimensions: " & v.Name
    End If
End Function

