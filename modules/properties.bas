Attribute VB_Name = "properties"
Option Explicit
Private Function PropertyValue(ByVal doc As SldWorks.ModelDoc2, ByVal cfg As String, ByVal aliases As String) As String
    Dim a As Variant, n As Variant, pm As Object, raw As String, value As String
    Dim resolved As Boolean, linked As Boolean, which As Long
    a = Split(aliases, "|")
    For which = 0 To 1
        If which = 0 Then
            Set pm = doc.Extension.CustomPropertyManager(cfg)
        Else
            Set pm = doc.Extension.CustomPropertyManager("")
        End If
        For Each n In a
            raw = "": value = ""
            pm.Get6 CStr(n), False, raw, value, resolved, linked
            If Len(Trim$(value)) > 0 Then
                PropertyValue = value
                Exit Function
            End If
        Next n
    Next which
End Function

Private Function MassText(ByVal doc As SldWorks.ModelDoc2, ByVal body As Object) As String
    Dim mp As SldWorks.MassProperty2, a(0) As Object
    On Error GoTo Unknown
    Set mp = doc.Extension.CreateMassProperty2
    mp.UseSystemUnits = True
    mp.IncludeHiddenBodiesOrComponents = False
    If Not body Is Nothing Then
        Set a(0) = body
        mp.SelectedItems = a
    End If
    mp.Recalculate
    MassText = Format$(mp.Mass, "0.000")
    Exit Function
Unknown:
    MassText = ""
    Warn "Mass unavailable for " & doc.GetTitle
End Function

Private Function OriginalBase(ByVal p As String) As String
    OriginalBase = fs.GetBaseName(p)
    If Left$(OriginalBase, Len(prefix)) = prefix Then
        OriginalBase = Mid$(OriginalBase, Len(prefix) + 1)
    End If
End Function

Private Function SafeFile(ByVal text As String) As String
    Dim bad As Variant, ch As Variant
    bad = Array("<", ">", ":", """", "/", "\", "|", "?", "*", vbTab, vbCr, vbLf)
    For Each ch In bad: text = Replace(text, CStr(ch), "_"): Next ch
    text = Trim$(text)
    Do While Len(text) > 0
        If Right$(text, 1) <> "." And Right$(text, 1) <> " " Then
            Exit Do
        End If
        text = Left$(text, Len(text) - 1)
    Loop
    If Len(text) = 0 Then
        text = "unnamed"
    End If
    SafeFile = Left$(text, MAX_GENERATED_STEM_LENGTH)
End Function

Private Function ValidArtifactStem(ByVal text As String) As Boolean
    ' Validate the persisted basename without sanitizing, truncating or renaming it.
    ' SafeFile limits individual name segments; their combined basename can be longer.
    Dim i As Long, ch As String, code As Long, deviceName As String
    If Len(text) = 0 Or Len(text) + Len(DRAWING_EXTENSION) > MAX_WINDOWS_FILENAME_LENGTH Then
        Exit Function
    End If
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1): code = AscW(ch)
        If code >= 0 And code < ASCII_CONTROL_LIMIT Then
            Exit Function
        End If
        If InStr(1, "<>:""/\|?*", ch, vbBinaryCompare) > 0 Then
            Exit Function
        End If
    Next i
    ' The actual filenames include an extension, so a stem may end in a dot/space.
    deviceName = UCase$(Trim$(Split(text, ".")(0)))
    Select Case deviceName
        Case "", "CON", "PRN", "AUX", "NUL", "CONIN$", "CONOUT$": Exit Function
    End Select
    If Len(deviceName) = RESERVED_DEVICE_NAME_LENGTH Then
        If Left$(deviceName, RESERVED_DEVICE_PREFIX_LENGTH) = "COM" Or Left$(deviceName, RESERVED_DEVICE_PREFIX_LENGTH) = "LPT" Then
            If InStr(1, "123456789" & ChrW$(WINDOWS_DEVICE_SUPERSCRIPT_ONE) & ChrW$(WINDOWS_DEVICE_SUPERSCRIPT_TWO) & ChrW$(WINDOWS_DEVICE_SUPERSCRIPT_THREE), Right$(deviceName, 1), vbBinaryCompare) > 0 Then
                Exit Function
            End If
        End If
    End If
    ValidArtifactStem = True
End Function

Private Function ScaleText(ByVal viewScaleValue As Double) As String
    If viewScaleValue >= 1 Then
        ScaleText = CStr(viewScaleValue) & ":1"
    Else
        ScaleText = "1:" & CStr(1 / viewScaleValue)
    End If
End Function

Private Function TextFit(ByVal text As String, ByVal width As Double, ByVal height As Double) As Double
    If Len(text) = 0 Then
        TextFit = height
        Exit Function
    End If
    TextFit = MinD(height, width / (Len(text) * TEXT_ESTIMATED_GLYPH_WIDTH_RATIO))
    If TextFit < TEXT_HEIGHT_SMALL_M Then
        TextFit = TEXT_HEIGHT_SMALL_M
    End If
End Function

Private Function WrapText(ByVal text As String, ByVal maxChars As Long) As String
    Dim lines As Variant, line As Variant, s As String, cut As Long, result As String
    lines = Split(Replace(text, vbCr, ""), vbLf)
    For Each line In lines
        s = CStr(line)
        Do While Len(s) > maxChars
            cut = InStrRev(Left$(s, maxChars), " ")
            If cut < maxChars \ 2 Then
                cut = maxChars
            End If
            result = result & Left$(s, cut) & vbCrLf
            s = LTrim$(Mid$(s, cut + 1))
        Loop
        result = result & s & vbCrLf
    Next line
    If Len(result) > 1 Then
        result = Left$(result, Len(result) - 2)
    End If
    WrapText = result
End Function

Private Function MinD(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then
        MinD = a
    Else
        MinD = b
    End If
End Function

