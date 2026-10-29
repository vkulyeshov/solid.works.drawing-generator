Attribute VB_Name = "runtime"
Option Explicit
Private Sub ReadSettings(ByVal p As String)
    Dim f As Object, line As String, i As Long
    If Not fs.FileExists(p) Then
        Exit Sub
    End If
    Set f = fs.OpenTextFile(p, FILE_FOR_READING, False, FILE_UNICODE)
    Do Until f.AtEndOfStream
        line = Trim$(f.ReadLine)
        If Len(line) > 0 And Left$(line, 1) <> ";" And Left$(line, 1) <> "[" Then
            i = InStr(line, "=")
            If i > 1 Then
                settings(Trim$(Left$(line, i - 1))) = Trim$(Mid$(line, i + 1))
            End If
        End If
    Loop
    f.Close
End Sub

Private Function Setting(ByVal key As String, ByVal fallback As String) As String
    Setting = fallback
    If settings.Exists(key) Then
        Setting = CStr(settings(key))
    End If
End Function



Private Sub RestorePreferences()
    On Error Resume Next
    If preferenceChanged Then
        sw.SetUserPreferenceToggle SW_PREF_INPUT_DIMENSION, inputDimPreference
        sw.SetUserPreferenceToggle SW_PREF_PDF_EMBED_FONTS, oldEmbedFonts
        sw.SetUserPreferenceToggle SW_PREF_PDF_IN_COLOR, oldPdfColor
        sw.SetUserPreferenceToggle SW_PREF_PDF_HIGH_QUALITY, oldPdfQuality
    End If
    preferenceChanged = False
End Sub

Private Sub CloseSnapshots()
    Dim docs As Variant, item As Variant, doc As SldWorks.ModelDoc2, pass As Long
    On Error Resume Next
    For pass = SW_DOC_ASSEMBLY To SW_DOC_PART Step -1
        docs = sw.GetDocuments
        If IsEmpty(docs) Then
            Exit Sub
        End If
        For Each item In docs
            Set doc = item
            If InStr(1, doc.GetPathName, runDir & MODELS_RELATIVE_DIRECTORY, vbTextCompare) = 1 Then
                If doc.GetType = pass Then
                    sw.CloseDoc doc.GetTitle
                End If
            End If
        Next item
    Next pass
End Sub

Private Sub Log(ByVal text As String)
    HideBackgroundFrame
    If Len(runDir) > 0 Then
        Set logFile = fs.OpenTextFile(runDir & RUN_LOG_RELATIVE_PATH, FILE_FOR_APPENDING, True, FILE_UNICODE)
        logFile.WriteLine Format$(Now, "hh:nn:ss") & " " & text
        logFile.Close: Set logFile = Nothing
    End If
    ReportProgress STATE_RUNNING, text
    Debug.Print text
End Sub

Private Sub ReportProgress(ByVal state As String, ByVal message As String)
    Dim stream As Object
    If Len(progressPath) = 0 Then
        Exit Sub
    End If
    ' Close after every record: the VBS monitor reads while RunMacro2 is blocked.
    On Error Resume Next
    message = Replace(Replace(Replace(message, vbTab, " "), vbCr, " "), vbLf, " ")
    Set stream = fs.OpenTextFile(progressPath, FILE_FOR_APPENDING, True, FILE_UNICODE)
    stream.WriteLine Format$(Now, "hh:nn:ss") & vbTab & state & vbTab & (completed + failures) & vbTab & plannedJobs & vbTab & completed & vbTab & failures & vbTab & warningCount & vbTab & runDir & vbTab & message
    stream.Close
End Sub

Private Sub Warn(ByVal text As String)
    warningCount = warningCount + 1
    Log "WARN " & text
End Sub

Private Function AsDrawing(ByVal doc As SldWorks.ModelDoc2) As SldWorks.DrawingDoc
    Set AsDrawing = doc
End Function
Private Function AsPart(ByVal doc As SldWorks.ModelDoc2) As SldWorks.PartDoc
    Set AsPart = doc
End Function
Private Function AsAssembly(ByVal doc As SldWorks.ModelDoc2) As SldWorks.AssemblyDoc
    Set AsAssembly = doc
End Function


