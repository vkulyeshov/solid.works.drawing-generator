Attribute VB_Name = "planar_text"
Option Explicit

Private Function PlanarMeasureTextWidth(ByVal display As SldWorks.DisplayDimension, ByVal text As String, ByVal height As Double, Optional ByVal printed As Boolean = False, Optional ByVal faceOverride As String = "") As Double
    Dim format As SldWorks.TextFormat, extent As DrawingTextExtent, key As String, weight As Long, face As String, measured As Double
    Dim metrics As DrawingFontMetrics, measurementHeight As Double
#If VBA7 Then
    Dim device As LongPtr, font As LongPtr, previous As LongPtr
#Else
    Dim device As Long, font As Long, previous As Long
#End If
    If Len(text) = 0 Then
        Exit Function
    End If
    On Error GoTo Fallback
    Set format = display.GetAnnotation.GetTextFormat(SW_ANNOTATION_TEXT_FORMAT_FIRST)
    face = format.TypeFaceName
    If Len(faceOverride) > 0 Then
        face = faceOverride
    End If
    key = face & vbTab & CStr(format.Bold) & vbTab & CStr(format.Italic) & vbTab & CStr(printed) & vbTab & text
    If planarTextMetricCache Is Nothing Then
        Set planarTextMetricCache = CreateObject("Scripting.Dictionary")
    End If
    If planarTextMetricCache.Exists(key) Then
        PlanarMeasureTextWidth = CDbl(planarTextMetricCache(key)) * height
        Exit Function
    End If
    weight = GDI_NORMAL_FONT_WEIGHT
    If format.Bold Then
        weight = GDI_BOLD_FONT_WEIGHT
    End If
    device = CreateCompatibleDC(GDI_DEFAULT_OPTION)
    If device = 0 Then
        GoTo Fallback
    End If
    font = CreateFontW(-PLANAR_FONT_MEASUREMENT_PIXELS, GDI_DEFAULT_OPTION, GDI_DEFAULT_OPTION, GDI_DEFAULT_OPTION, weight, Abs(CLng(format.Italic)), GDI_DEFAULT_OPTION, GDI_DEFAULT_OPTION, GDI_DEFAULT_CHARACTER_SET, GDI_DEFAULT_OPTION, GDI_DEFAULT_OPTION, GDI_DEFAULT_OPTION, GDI_DEFAULT_OPTION, StrPtr(face))
    If font = 0 Then
        GoTo Fallback
    End If
    previous = SelectObject(device, font)
    If GetTextExtentPoint32W(device, StrPtr(text), Len(text), extent) = 0 Then
        GoTo Fallback
    End If
    measurementHeight = PLANAR_FONT_MEASUREMENT_PIXELS
    If printed Then
        If GetTextMetricsW(device, metrics) <> 0 And metrics.Ascent > 0 Then
            ' SolidWorks specifies the ascent, whereas a negative GDI font
            ' height specifies the em. They differ even in the GOST font.
            measurementHeight = metrics.Ascent
        End If
    End If
    measured = extent.Width / measurementHeight
    If printed And Len(text) > LAYOUT_METRIC_AUDIT_LENGTH And Setting("DimensionLayoutAudit", "0") = "1" Then
        Log "FONT METRIC face=" & face & "; h=" & height & "; format h=" & format.CharHeight & "; pixels=" & extent.Width & "; ascent=" & measurementHeight & "; ratio=" & measured & "; text=" & text
    End If
    Call SelectObject(device, previous)
    Call DeleteObject(font)
    Call DeleteDC(device)
    planarTextMetricCache.Add key, measured
    PlanarMeasureTextWidth = measured * height
    Exit Function
Fallback:
    If printed And Setting("DimensionLayoutAudit", "0") = "1" Then
        Log "FONT FALLBACK error=" & Err.Number & "; " & Err.Description & "; face=" & face & "; h=" & height & "; text=" & text
    End If
    On Error Resume Next
    If previous <> 0 And device <> 0 Then
        Call SelectObject(device, previous)
    End If
    If font <> 0 Then
        Call DeleteObject(font)
    End If
    If device <> 0 Then
        Call DeleteDC(device)
    End If
    PlanarMeasureTextWidth = Len(text) * height * LEADER_GLYPH_WIDTH_RATIO
End Function
