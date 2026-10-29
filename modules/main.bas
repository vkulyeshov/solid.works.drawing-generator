Attribute VB_Name = "main"
Option Explicit
Public Sub main()
    Dim source As SldWorks.ModelDoc2, p As String, cfg As String, key As Variant, job As Object
    Dim errNumber As Long, errText As String
    Dim retryBatch As Boolean, attemptSetting As String
    On Error GoTo Fatal
    Set sw = Application.SldWorks
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set settings = CreateObject("Scripting.Dictionary")
    settings.CompareMode = 1
    progressPath = "": runDir = "": plannedJobs = 0
    backgroundOwned = False: attempts = 0: resumeRun = False
    completed = 0: failures = 0: warningCount = 0: jobNumber = 0: viewCounter = 0
    batchFailures = 0: interruptedJobNumber = 0
    ReadSettings sw.GetCurrentMacroPathFolder & SETTINGS_RELATIVE_PATH
    progressPath = Setting("LauncherProgress", "")
    attemptSetting = Setting("MaxJobAttempts", DEFAULT_MAX_JOB_ATTEMPTS_TEXT)
    If Not IsNumeric(attemptSetting) Then
        Err.Raise vbObjectError + ERR_DRAWING_SETTINGS, , "MaxJobAttempts must be a positive integer."
    End If
    maxJobAttempts = CLng(attemptSetting)
    If maxJobAttempts < 1 Or CDbl(attemptSetting) <> maxJobAttempts Then
        Err.Raise vbObjectError + ERR_DRAWING_SETTINGS, , "MaxJobAttempts must be a positive integer."
    End If
    ReportProgress STATE_RUNNING, "VBA started"
    If Setting("LauncherBackground", "0") = "1" Then
        If CStr(sw.GetProcessID) = Setting("LauncherForbiddenPID", "0") Then
            Err.Raise vbObjectError + ERR_APPLICATION_STATE, , "Background startup reused the user's SolidWorks process."
        End If
        backgroundOwned = True
        ' Retain the background application when the last visible drawing closes.
        sw.UserControlBackground = True
        Dim hiddenFrame As SldWorks.Frame, inputType As Long, inputError As Long, inputWarning As Long
        Set hiddenFrame = sw.Frame: hiddenFrame.KeepInvisible = True
        backgroundLeft = sw.FrameLeft: backgroundTop = sw.FrameTop
        backgroundWidth = sw.FrameWidth: backgroundHeight = sw.FrameHeight
        HideBackgroundFrame
        inputType = SW_DOC_PART
        If LCase$(fs.GetExtensionName(Setting("LauncherInput", ""))) = ASSEMBLY_EXTENSION Then
            inputType = SW_DOC_ASSEMBLY
        End If
        sw.SetCurrentWorkingDirectory fs.GetParentFolderName(Setting("LauncherInput", ""))
        Set original = sw.OpenDoc6(Setting("LauncherInput", ""), inputType, SW_OPEN_SILENT_READ_ONLY, "", inputError, inputWarning)
        If original Is Nothing Then
            Err.Raise vbObjectError + ERR_OPEN_INPUT_OR_SNAPSHOT, , "Cannot open background input: " & inputError
        End If
        PrepareDocumentWindow original
        Set source = sw.ActivateDoc3(fs.GetFileName(original.GetPathName), False, SW_REBUILD_USER_DECISION, inputError)
        ReportProgress STATE_RUNNING, "Hidden SolidWorks PID=" & sw.GetProcessID
    Else
        Set original = sw.ActiveDoc
    End If
    If original Is Nothing Then
        Err.Raise vbObjectError + ERR_SOURCE_NOT_OPEN, , "Open a saved part or assembly first."
    End If
    If original.GetType <> SW_DOC_PART And original.GetType <> SW_DOC_ASSEMBLY Then
        Err.Raise vbObjectError + ERR_SOURCE_TYPE, , "Activate a part or assembly, not a drawing."
    End If
    p = original.GetPathName
    If Len(p) = 0 Then
        Err.Raise vbObjectError + ERR_SOURCE_NOT_SAVED, , "Save the source model before running."
    End If
    If Len(Setting("LauncherInput", "")) > 0 Then
        If StrComp(p, Setting("LauncherInput", ""), vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + ERR_APPLICATION_STATE, , "The requested input model is not active."
        End If
    End If
    sourceDir = fs.GetParentFolderName(p)
    ResolveDrawingTemplate
    ValidateDrawingTemplate
    runDir = Setting("LauncherOutput", "")
    If Len(runDir) = 0 Then
        runDir = sourceDir & OUTPUT_FOLDER_RELATIVE_PREFIX & Format$(Now, TIMESTAMP_FORMAT)
        If fs.FolderExists(runDir) Then
            Err.Raise vbObjectError + ERR_OUTPUT_DIRECTORY, , "Output folder already exists. Try again in a second."
        End If
    Else
        runDir = fs.GetAbsolutePathName(runDir)
    End If
    EnsureFolder runDir
    EnsureFolder runDir & MODELS_RELATIVE_PATH
    EnsureFolder runDir & DRAWINGS_RELATIVE_PATH
    EnsureFolder runDir & PDF_RELATIVE_PATH
    runtimeFormatPath = runDir & SHEET_FORMAT_RELATIVE_PATH
    sourcePath = original.GetPathName
    sourceStamp = FileStamp(sourcePath)
    checkpointPath = runDir & CHECKPOINT_RELATIVE_PATH
    renderSignature = RenderSettingsSignature()
    Set jobs = CreateObject("Scripting.Dictionary"): jobs.CompareMode = 1
    resumeRun = fs.FileExists(checkpointPath)
    Dim er As Long, wr As Long
    If resumeRun Then
        LoadCheckpoint
        SaveCheckpoint
        Log "RESUME existing plan; completed=" & completed & " total=" & jobs.Count
    Else
        ' A new snapshot must reflect a saved source. Resume uses only saved Models/.
        If original.GetSaveFlag Then
            Err.Raise vbObjectError + ERR_UNSAVED_SOURCE, , "Source has unsaved changes. Save or discard them yourself before creating a new snapshot."
        End If
        Set logFile = fs.CreateTextFile(runDir & RUN_LOG_RELATIVE_PATH, True, True)
        logFile.Close: Set logFile = Nothing
        Set manifest = fs.CreateTextFile(runDir & MANIFEST_RELATIVE_PATH, True, True)
        manifest.WriteLine "Number" & vbTab & "Model" & vbTab & "Configuration" & vbTab & "Body" & vbTab & "Quantity" & vbTab & "Drawing" & vbTab & "Status"
        manifest.Close: Set manifest = Nothing
        Log "START SolidWorks " & sw.RevisionNumber & " | " & sourcePath
        cfg = original.ConfigurationManager.ActiveConfiguration.Name
        snapshotConfig = cfg
        snapshotPath = MakeSnapshot(original)
        Set rootModel = sw.OpenDoc6(snapshotPath, original.GetType, SW_OPEN_SILENT, cfg, er, wr)
        If rootModel Is Nothing Then
            Err.Raise vbObjectError + ERR_OPEN_INPUT_OR_SNAPSHOT, , "Cannot open snapshot: " & er
        End If
        Set quantities = CreateObject("Scripting.Dictionary"): quantities.CompareMode = 1
        If rootModel.GetType = SW_DOC_ASSEMBLY Then
            AsAssembly(rootModel).ResolveAllLightWeightComponents False
            Inventory rootModel
        End If
        AddModelJobs rootModel, cfg, 1
        For Each key In quantities.Keys
            Set job = quantities(key): Set source = job("doc")
            AddModelJobs source, CStr(job("cfg")), CLng(job("qty"))
        Next key
        SelectPlan
        SaveCheckpoint
    End If
    plannedJobs = jobs.Count
    Log "JOBS selected=" & plannedJobs
    If plannedJobs = 0 Then
        Err.Raise vbObjectError + ERR_EMPTY_PLAN, , "No drawing jobs match the settings."
    End If
    ' Inventory is now plain data. Release all model COM references before drawing.
    Set source = Nothing: Set rootModel = Nothing: Set quantities = Nothing
    For Each key In jobs.Keys
        Set job = jobs(key)
        If job.Exists("doc") Then
            job.Remove "doc"
        End If
    Next key
    CloseSnapshots
    If backgroundOwned Then
        sw.CloseDoc original.GetTitle
        Set original = Nothing
    End If
    inputDimPreference = sw.GetUserPreferenceToggle(SW_PREF_INPUT_DIMENSION)
    oldEmbedFonts = sw.GetUserPreferenceToggle(SW_PREF_PDF_EMBED_FONTS)
    oldPdfColor = sw.GetUserPreferenceToggle(SW_PREF_PDF_IN_COLOR)
    oldPdfQuality = sw.GetUserPreferenceToggle(SW_PREF_PDF_HIGH_QUALITY)
    sw.SetUserPreferenceToggle SW_PREF_INPUT_DIMENSION, False
    sw.SetUserPreferenceToggle SW_PREF_PDF_EMBED_FONTS, True
    sw.SetUserPreferenceToggle SW_PREF_PDF_IN_COLOR, False
    sw.SetUserPreferenceToggle SW_PREF_PDF_HIGH_QUALITY, True
    preferenceChanged = True
    retryBatch = Not HasPendingJobs()
    WriteFailedJobs
    For Each key In jobs.Keys
        Set job = jobs(key)
        If JobEligible(job, retryBatch) Then
            jobNumber = CLng(job("number"))
            job("attempts") = CLng(job("attempts")) + 1
            job("status") = STATE_RUNNING: job("stage") = "Opening model"
            SaveCheckpoint
            attempts = attempts + 1
            Log "ATTEMPT " & job("out") & " " & job("attempts") & "/" & maxJobAttempts
            ProcessJob job
            interruptedJobNumber = 0
            WriteFailedJobs
            Set flatBody = Nothing
            ' Close copied assemblies before copied parts, after ProcessJob releases locals.
            CloseSnapshots
            If CLng(Setting("JobsPerBatch", DEFAULT_BATCH_SIZE_TEXT)) > 0 Then
                If attempts >= CLng(Setting("JobsPerBatch", DEFAULT_BATCH_SIZE_TEXT)) Then
                    Exit For
                End If
            End If
            If CLng(Setting("MaxFailures", "0")) > 0 Then
                If batchFailures >= CLng(Setting("MaxFailures", "0")) Then
                    Exit For
                End If
            End If
        End If
        DoEvents
    Next key
    RestorePreferences
    CloseSnapshots
    Log "BATCH END drawings=" & completed & " failed=" & failures & " attempts=" & attempts & " warnings=" & warningCount
    If Not logFile Is Nothing Then
        logFile.Close
    End If
    If Not manifest Is Nothing Then
        manifest.Close
    End If
    Set logFile = Nothing: Set manifest = Nothing
    Set currentDrawing = Nothing
    If Not original Is Nothing Then
        Set source = sw.ActivateDoc3(fs.GetFileName(original.GetPathName), False, SW_REBUILD_USER_DECISION, er)
    End If
    If backgroundOwned Then
        HideBackgroundFrame
        If FrameIsOffscreen() = False Then
            Err.Raise vbObjectError + ERR_APPLICATION_STATE, , "Background frame is on screen."
        End If
    End If
    If completed = plannedJobs Then
        ReportProgress STATE_COMPLETE, "Drawings=" & completed & "; warnings=" & warningCount
    ElseIf HasRunnableJobs() Then
        ReportProgress STATE_PAUSED, "Batch saved; drawings=" & completed & "/" & plannedJobs & "; failed=" & failures & "; continuing automatically"
    Else
        ReportProgress STATE_FAILED, "All jobs attempted; drawings=" & completed & "; failed=" & failures & "; selected=" & plannedJobs & "; attempt limits reached. See failed-jobs.tsv."
    End If
    Debug.Print "Drawing pack: " & completed & " drawings; " & failures & " failures. " & runDir
    If Setting("ShowSummary", "0") = "1" Then
        MsgBox "Drawings: " & completed & vbCrLf & "Failures: " & failures & vbCrLf & "Review warnings: " & warningCount & vbCrLf & runDir, vbInformation, "Drawing pack"
    End If
    FinishBackground
    Exit Sub
Fatal:
    errNumber = Err.Number: errText = Err.Description
    On Error Resume Next
    RestorePreferences
    If Not logFile Is Nothing Then
        logFile.WriteLine "FATAL " & errNumber & " " & errText
        logFile.Close
    End If
    If Not manifest Is Nothing Then
        manifest.Close
    End If
    ReportProgress STATE_FAILED, "VBA error " & errNumber & ": " & errText
    If Setting("ShowSummary", "0") = "1" Then
        MsgBox "Macro stopped: " & errText & vbCrLf & runDir, vbExclamation, "Drawing pack"
    End If
    FinishBackground
End Sub
