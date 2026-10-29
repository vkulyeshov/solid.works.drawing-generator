Attribute VB_Name = "checkpoint"
Option Explicit
Private Function FileStamp(ByVal path As String) As String
    Dim f As Object
    Set f = fs.GetFile(path)
    FileStamp = CStr(f.Size) & "|" & Format$(f.DateLastModified, FILE_STAMP_FORMAT)
End Function

Private Function RenderSettingsSignature() As String
    Dim key As Variant
    RenderSettingsSignature = RENDER_SIGNATURE_PREFIX & paperFormat & "|Template=" & templatePath & "|TemplateStamp=" & FileStamp(templatePath)
    For Each key In Array("Font", "Organization", "AutoDimensions", "NativeDimensions", "MaxEdgesForAutoDimensions", "ImportModelDimensions", "UpAxis", "StrictDimensionLayout", "DimensionSpacing", "Liter1", "Liter2", "Liter3", "DrawnBy", "DrawnDate", "CheckedBy", "CheckedDate", "TechnicalControlBy", "TechnicalControlDate", "NormControlBy", "NormControlDate", "ApprovedBy", "ApprovedDate")
        RenderSettingsSignature = RenderSettingsSignature & "|" & CStr(key) & "=" & Setting(CStr(key), "")
    Next key
End Function

Private Sub SelectPlan()
    Dim selected As Object, key As Variant, j As Object, n As Long
    Set selected = CreateObject("Scripting.Dictionary"): selected.CompareMode = 1
    For Each key In jobs.Keys
        Set j = jobs(key)
        If ShouldRunJob(j) Then
            n = n + 1
            j.Add "number", n: j.Add "status", STATE_PENDING: j.Add "stage", "Queued"
            j.Add "drawingBytes", "0": j.Add "pdfBytes", "0"
            j.Add "attempts", 0
            j.Add "out", Format$(n, JOB_NUMBER_FORMAT) & "_" & SafeFile(OriginalBase(CStr(j("path"))) & "_" & CStr(j("cfg")))
            If CBool(j("multi")) Then
                j("out") = j("out") & "_" & SafeFile(CStr(j("cut")) & "_" & CStr(j("body")))
            End If
            selected.Add CStr(n), j
            If CLng(Setting("MaxJobs", "0")) > 0 Then
                If n >= CLng(Setting("MaxJobs", "0")) Then
                    Exit For
                End If
            End If
        End If
    Next key
    Set jobs = selected
End Sub

Private Function Cell(ByVal value As Variant) As String
    Cell = Replace(Replace(Replace(CStr(value), vbTab, " "), vbCr, " "), vbLf, " ")
End Function

Private Sub SaveCheckpoint()
    Dim stream As Object, tempPath As String, key As Variant, j As Object, line As String, field As Variant
    tempPath = checkpointPath & ATOMIC_TEMP_SUFFIX
    Set stream = fs.CreateTextFile(tempPath, True, True)
    stream.WriteLine CHECKPOINT_HEADER_TAG & vbTab & CHECKPOINT_VERSION & vbTab & sourcePath & vbTab & sourceStamp & vbTab & snapshotPath & vbTab & snapshotConfig & vbTab & prefix & vbTab & renderSignature
    For Each key In jobs.Keys
        Set j = jobs(key): line = CHECKPOINT_JOB_TAG
        For Each field In Array("number", "path", "cfg", "body", "cut", "qty", "sheet", "multi", "type", "status", "out", "stage", "drawingBytes", "pdfBytes", "attempts")
            line = line & vbTab & Cell(j(CStr(field)))
        Next field
        stream.WriteLine line
    Next key
    stream.Close
    ' Atomic replacement on the same volume. Old checkpoint survives an interrupted write.
    If MoveFileExW(StrPtr(tempPath), StrPtr(checkpointPath), ATOMIC_REPLACE_WRITE_THROUGH) = 0 Then
        Err.Raise vbObjectError + ERR_GENERATION_DATA, , "Cannot commit generation-state.tsv"
    End If
End Sub

Private Sub LoadCheckpoint()
    Dim stream As Object, fields As Variant, names As Variant, j As Object, i As Long
    Dim legacy As Boolean, count As Double
    Set stream = fs.OpenTextFile(checkpointPath, FILE_FOR_READING, False, FILE_UNICODE)
    If stream.AtEndOfStream Then
        Err.Raise vbObjectError + ERR_CHECKPOINT_INVALID, , "Empty checkpoint. Use /reset to generate again."
    End If
    fields = Split(stream.ReadLine, vbTab)
    If UBound(fields) <> CHECKPOINT_HEADER_LAST_FIELD Then
        Err.Raise vbObjectError + ERR_CHECKPOINT_INVALID, , "Invalid checkpoint header."
    End If
    If fields(HEADER_TAG_FIELD) <> CHECKPOINT_HEADER_TAG Then
        Err.Raise vbObjectError + ERR_CHECKPOINT_INVALID, , "Unsupported checkpoint version."
    End If
    legacy = (fields(HEADER_VERSION_FIELD) = LEGACY_CHECKPOINT_VERSION)
    If Not legacy And fields(HEADER_VERSION_FIELD) <> CHECKPOINT_VERSION Then
        Err.Raise vbObjectError + ERR_CHECKPOINT_INVALID, , "Unsupported checkpoint version."
    End If
    If StrComp(fields(HEADER_SOURCE_PATH), sourcePath, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + ERR_CHECKPOINT_INVALID, , "Checkpoint belongs to another input. Choose another output folder."
    End If
    If fields(HEADER_SOURCE_STAMP) <> sourceStamp Then
        Err.Raise vbObjectError + ERR_CHECKPOINT_INVALID, , "Input file changed. Use /reset for a new snapshot."
    End If
    If fields(HEADER_RENDER_SIGNATURE) <> renderSignature Then
        Err.Raise vbObjectError + ERR_CHECKPOINT_INVALID, , "Drawing settings or generator version changed. Use /reset."
    End If
    snapshotPath = fields(HEADER_SNAPSHOT_PATH): snapshotConfig = fields(HEADER_SNAPSHOT_CONFIG): prefix = fields(HEADER_PREFIX)
    names = Array("number", "path", "cfg", "body", "cut", "qty", "sheet", "multi", "type", "status", "out", "stage", "drawingBytes", "pdfBytes")
    Do While Not stream.AtEndOfStream
        fields = Split(stream.ReadLine, vbTab)
        If (legacy And UBound(fields) <> LEGACY_CHECKPOINT_JOB_LAST_FIELD) Or (Not legacy And UBound(fields) <> CHECKPOINT_JOB_LAST_FIELD) Then
            Err.Raise vbObjectError + ERR_CHECKPOINT_INVALID, , "Invalid checkpoint job."
        End If
        Set j = CreateObject("Scripting.Dictionary")
        For i = 0 To UBound(names): j.Add CStr(names(i)), fields(i + 1): Next i
        count = 0
        If legacy Then
            If j("status") = STATE_DONE Or j("status") = STATE_FAILED Or j("status") = STATE_RUNNING Then
                count = 1
            End If
        Else
            If Not IsNumeric(fields(CHECKPOINT_ATTEMPTS_FIELD)) Then
                Err.Raise vbObjectError + ERR_CHECKPOINT_INVALID, , "Invalid checkpoint attempt count."
            End If
            count = CDbl(fields(CHECKPOINT_ATTEMPTS_FIELD))
            If count < 0 Or count <> Fix(count) Then
                Err.Raise vbObjectError + ERR_CHECKPOINT_INVALID, , "Invalid checkpoint attempt count."
            End If
        End If
        j.Add "attempts", CLng(count)
        j("sheet") = CBool(j("sheet")): j("multi") = CBool(j("multi"))
        j("path") = fs.GetAbsolutePathName(CStr(j("path")))
        If Not ValidArtifactStem(CStr(j("out"))) Then
            Err.Raise vbObjectError + ERR_CHECKPOINT_INVALID, , "Invalid checkpoint artifact name for job " & CStr(j("number")) & ": " & CStr(j("out"))
        End If
        If InStr(1, CStr(j("path")), runDir & MODELS_RELATIVE_DIRECTORY, vbTextCompare) <> 1 Then
            Err.Raise vbObjectError + ERR_CHECKPOINT_INVALID, , "Snapshot path is outside output/Models."
        End If
        If Not fs.FileExists(CStr(j("path"))) Then
            Err.Raise vbObjectError + ERR_CHECKPOINT_INVALID, , "Snapshot model missing. Use /reset."
        End If
        If j("status") = STATE_DONE Then
            If ArtifactsValid(j) Then
                completed = completed + 1
            Else
                j("status") = STATE_PENDING: j("stage") = "Artifact missing or changed; regenerate"
                j("attempts") = 0
            End If
        ElseIf j("status") = STATE_RUNNING Then
            j("status") = STATE_FAILED
            j("stage") = "Interrupted at: " & CStr(j("stage"))
            If CLng(j("attempts")) < maxJobAttempts Then
                interruptedJobNumber = CLng(j("number"))
            End If
        ElseIf j("status") <> STATE_PENDING And j("status") <> STATE_FAILED Then
            Err.Raise vbObjectError + ERR_CHECKPOINT_INVALID, , "Invalid checkpoint job status."
        End If
        jobs.Add CStr(j("number")), j
    Loop
    stream.Close
    If jobs.Count = 0 Then
        Err.Raise vbObjectError + ERR_CHECKPOINT_INVALID, , "Checkpoint has no jobs."
    End If
    RecountJobResults
End Sub

Private Sub RecountJobResults()
    Dim key As Variant, job As Object
    completed = 0
    failures = 0
    For Each key In jobs.Keys
        Set job = jobs(key)
        If job("status") = STATE_DONE Then
            completed = completed + 1
        ElseIf job("status") = STATE_FAILED Then
            failures = failures + 1
        End If
    Next key
End Sub

Private Function HasPendingJobs() As Boolean
    Dim key As Variant, job As Object
    For Each key In jobs.Keys
        Set job = jobs(key)
        If job("status") = STATE_PENDING And CLng(job("attempts")) < maxJobAttempts Then
            HasPendingJobs = True
            Exit Function
        End If
    Next key
End Function

Private Function HasRunnableJobs() As Boolean
    Dim key As Variant, job As Object
    For Each key In jobs.Keys
        Set job = jobs(key)
        If job("status") <> STATE_DONE And CLng(job("attempts")) < maxJobAttempts Then
            HasRunnableJobs = True
            Exit Function
        End If
    Next key
End Function

Private Function JobEligible(ByVal job As Object, ByVal retryBatch As Boolean) As Boolean
    If job("status") = STATE_DONE Or CLng(job("attempts")) >= maxJobAttempts Then
        Exit Function
    End If
    If interruptedJobNumber > 0 Then
        JobEligible = (CLng(job("number")) = interruptedJobNumber)
    ElseIf retryBatch Then
        JobEligible = (job("status") = STATE_FAILED)
    Else
        JobEligible = (job("status") = STATE_PENDING)
    End If
End Function

Private Sub WriteFailedJobs()
    Dim stream As Object, key As Variant, job As Object, path As String, temp As String
    On Error GoTo Failed
    path = runDir & FAILED_JOBS_RELATIVE_PATH
    temp = path & ATOMIC_TEMP_SUFFIX
    Set stream = fs.CreateTextFile(temp, True, True)
    stream.WriteLine "Number" & vbTab & "Model" & vbTab & "Configuration" & vbTab & "Body" & vbTab & "Attempts" & vbTab & "MaxAttempts" & vbTab & "Retryable" & vbTab & "Error"
    For Each key In jobs.Keys
        Set job = jobs(key)
        If job("status") = STATE_FAILED Then
            stream.WriteLine Cell(job("number")) & vbTab & Cell(job("path")) & vbTab & Cell(job("cfg")) & vbTab & Cell(job("body")) & vbTab & job("attempts") & vbTab & maxJobAttempts & vbTab & CStr(CLng(job("attempts")) < maxJobAttempts) & vbTab & Cell(job("stage"))
        End If
    Next key
    stream.Close
    If MoveFileExW(StrPtr(temp), StrPtr(path), ATOMIC_REPLACE_WRITE_THROUGH) = 0 Then
        Err.Raise vbObjectError + ERR_GENERATION_DATA, , "Cannot commit failed-jobs.tsv"
    End If
    Exit Sub
Failed:
    On Error Resume Next
    If Not stream Is Nothing Then
        stream.Close
    End If
    Warn "Cannot update failed-jobs.tsv: " & Err.Description
End Sub

Private Function ArtifactsValid(ByVal j As Object) As Boolean
    Dim native As String, pdf As String
    native = runDir & DRAWINGS_RELATIVE_DIRECTORY & CStr(j("out")) & DRAWING_EXTENSION
    pdf = runDir & PDF_RELATIVE_DIRECTORY & CStr(j("out")) & PDF_EXTENSION
    If Not fs.FileExists(native) Or Not fs.FileExists(pdf) Then
        Exit Function
    End If
    ArtifactsValid = (CStr(fs.GetFile(native).Size) = CStr(j("drawingBytes")) And CStr(fs.GetFile(pdf).Size) = CStr(j("pdfBytes")) And CDbl(j("pdfBytes")) > 0)
End Function

Private Sub JobStage(ByVal j As Object, ByVal stage As String)
    j("stage") = stage
    SaveCheckpoint
    Log "STAGE " & CStr(j("out")) & " | " & stage
End Sub

Private Sub AppendManifest(ByVal line As String)
    Set manifest = fs.OpenTextFile(runDir & MANIFEST_RELATIVE_PATH, FILE_FOR_APPENDING, True, FILE_UNICODE)
    manifest.WriteLine line
    manifest.Close: Set manifest = Nothing
End Sub

