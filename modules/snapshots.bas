Attribute VB_Name = "snapshots"
Option Explicit
Private Function ShouldRunJob(ByVal job As Object) As Boolean
    ShouldRunJob = Len(Setting("JobFilter", "")) = 0 Or InStr(1, CStr(job("path")) & "|" & CStr(job("body")), Setting("JobFilter", ""), vbTextCompare) > 0
End Function

Private Sub EnsureFolder(ByVal path As String)
    Dim parent As String
    If fs.FolderExists(path) Then
        Exit Sub
    End If
    parent = fs.GetParentFolderName(path)
    If Len(parent) = 0 Or parent = path Then
        Err.Raise vbObjectError + ERR_OUTPUT_DIRECTORY, , "Cannot create output folder: " & path
    End If
    EnsureFolder parent
    fs.CreateFolder path
End Sub

Private Function ConfigureSnapshotTargets(ByVal pg As Object, ByVal names As Variant) As Variant
    Dim targets As Variant, documentStates As Variant, fileIndex As Long
    Dim folder As String, target As String, seen As Object
    folder = fs.GetAbsolutePathName(runDir & MODELS_RELATIVE_PATH)
    ' Let Pack and Go manage imported CAD dependencies as well as native documents.
    ' SetDocumentSaveToNames can reject a list containing a linked STEP document.
    If Not pg.SetSaveToName(True, folder) Then
        Err.Raise vbObjectError + ERR_PACK_TARGETS, , "Pack and Go rejected the snapshot folder: " & folder
    End If
    pg.AddPrefix = prefix
    pg.AddSuffix = ""
    If Not pg.GetDocumentSaveToNames(targets, documentStates) Then
        Err.Raise vbObjectError + ERR_PACK_TARGETS, , "Cannot read the Pack and Go destination names."
    End If
    If Not IsArray(targets) Then
        Err.Raise vbObjectError + ERR_PACK_TARGETS, , "Pack and Go returned no destination array."
    End If
    If LBound(targets) <> LBound(names) Or UBound(targets) <> UBound(names) Then
        Err.Raise vbObjectError + ERR_PACK_TARGETS, , "Pack and Go destination count does not match the dependency list."
    End If
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    For fileIndex = LBound(names) To UBound(names)
        If Len(CStr(targets(fileIndex))) = 0 Then
            Err.Raise vbObjectError + ERR_PACK_TARGETS, , "Pack and Go excluded a required dependency: " & CStr(names(fileIndex))
        End If
        target = fs.GetAbsolutePathName(CStr(targets(fileIndex)))
        If StrComp(fs.GetParentFolderName(target), folder, vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + ERR_PACK_TARGETS, , "Pack and Go destination is outside Models: " & target
        End If
        If StrComp(fs.GetExtensionName(target), fs.GetExtensionName(CStr(names(fileIndex))), vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + ERR_PACK_TARGETS, , "Pack and Go changed a dependency file type: " & target
        End If
        If seen.Exists(target) Then
            Err.Raise vbObjectError + ERR_PACK_TARGETS, , "Duplicate Pack and Go destination: " & target
        End If
        seen.Add target, True
    Next fileIndex
    Log "PACK destination names configured; files=" & seen.Count
    ConfigureSnapshotTargets = targets
End Function

Private Function MakeSnapshot(ByVal doc As SldWorks.ModelDoc2) As String
    Dim extension As SldWorks.ModelDocExtension, pg As SldWorks.PackAndGo, names As Variant, results As Variant, n As Variant, seen As Object
    Dim backupList As Object, backupComplete As Boolean, savedError As Long, savedText As String
    Dim destinationNames As Variant, fileIndex As Long
    On Error GoTo SnapshotFailed
    Log "PACK get object"
    Set extension = doc.Extension
    Set pg = extension.GetPackAndGo()

    pg.IncludeDrawings = False
    pg.IncludeSimulationResults = False
    pg.IncludeToolboxComponents = True
    pg.IncludeSuppressed = False
    pg.FlattenToSingleFolder = True
    Log "PACK get names"
    pg.GetDocumentNames names
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = 1
    For Each n In names
        If Not fs.FileExists(CStr(n)) Then
            Err.Raise vbObjectError + ERR_DEPENDENCY_MISSING, , "Missing reference: " & CStr(n)
        End If
        If seen.Exists(fs.GetFileName(CStr(n))) Then
            Err.Raise vbObjectError + ERR_DEPENDENCY_COLLISION, , "Duplicate dependency filenames. Use unique filenames before Pack and Go: " & CStr(n)
        End If
        seen.Add fs.GetFileName(CStr(n)), True
    Next n
    prefix = SNAPSHOT_PREFIX & Format$(Now, TIMESTAMP_FORMAT) & "_"
    destinationNames = ConfigureSnapshotTargets(pg, names)
    EnsureFolder runDir & SOURCE_BACKUP_RELATIVE_DIRECTORY & prefix
    Set backupList = fs.CreateTextFile(runDir & SOURCE_BACKUP_INDEX_RELATIVE_PATH, True, True)
    backupList.WriteLine "Original" & vbTab & "Backup"
    For Each n In names
        fs.CopyFile CStr(n), runDir & SOURCE_BACKUP_RELATIVE_DIRECTORY & prefix & "\" & fs.GetFileName(CStr(n)), False
        backupList.WriteLine CStr(n) & vbTab & "SourceBackup\" & prefix & "\" & fs.GetFileName(CStr(n))
    Next n
    backupList.Close
    backupComplete = True
    Log "SOURCE exact file backups created before Pack and Go"
    Log "PACK save"
    results = doc.Extension.SavePackAndGo(pg)
    If IsEmpty(results) Then
        Err.Raise vbObjectError + ERR_PACK_RESULT, , "Pack and Go returned no status."
    End If
    For Each n In results
        If CLng(n) <> 0 Then
            Err.Raise vbObjectError + ERR_PACK_STATUS, , "Pack and Go error " & CStr(n)
        End If
    Next n
    ' SavePackAndGo can serialize source documents even when opened read-only.
    ' Restore their exact bytes from the pre-operation backup before committing a plan.
    RestoreSnapshotSources names
    sourceStamp = FileStamp(sourcePath)
    MakeSnapshot = runDir & MODELS_RELATIVE_DIRECTORY & prefix & fs.GetFileName(doc.GetPathName)
    If Not fs.FileExists(MakeSnapshot) Then
        Err.Raise vbObjectError + ERR_SNAPSHOT_MISSING, , "Snapshot assembly is absent."
    End If
    For fileIndex = LBound(destinationNames) To UBound(destinationNames)
        If Not fs.FileExists(destinationNames(fileIndex)) Then
            Err.Raise vbObjectError + ERR_SNAPSHOT_MISSING, , "Snapshot file is absent: " & destinationNames(fileIndex)
        End If
        If fs.GetFile(destinationNames(fileIndex)).Size = 0 Then
            Err.Raise vbObjectError + ERR_SNAPSHOT_MISSING, , "Snapshot file is empty: " & destinationNames(fileIndex)
        End If
    Next fileIndex
    Log "SNAPSHOT " & MakeSnapshot
    backupComplete = False
    CleanupSnapshotBackup
    Exit Function
SnapshotFailed:
    savedError = Err.Number: savedText = Err.Description
    On Error Resume Next
    If Not backupList Is Nothing Then
        backupList.Close
    End If
    If backupComplete Then
        RestoreSnapshotSources names
    End If
    On Error GoTo 0
    Err.Raise savedError, "MakeSnapshot", savedText
End Function

Private Sub RestoreSnapshotSources(ByVal names As Variant)
    Dim n As Variant, exactBackup As String
    For Each n In names
        exactBackup = runDir & SOURCE_BACKUP_RELATIVE_DIRECTORY & prefix & "\" & fs.GetFileName(CStr(n))
        If FileStamp(CStr(n)) <> FileStamp(exactBackup) Then
            fs.CopyFile exactBackup, CStr(n), True
        End If
    Next n
End Sub

Private Sub CleanupSnapshotBackup()
    Dim backupRoot As String, backupFolder As String, indexPath As String, errorText As String
    On Error GoTo CleanupFailed
    backupRoot = fs.GetAbsolutePathName(runDir & SOURCE_BACKUP_RELATIVE_DIRECTORY)
    backupFolder = fs.GetAbsolutePathName(fs.BuildPath(backupRoot, prefix))
    If Len(prefix) = 0 Or StrComp(fs.GetParentFolderName(backupFolder), backupRoot, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + ERR_GENERATION_DATA, , "Invalid snapshot backup cleanup path."
    End If
    ' Remove only this snapshot's backup; retain backups from other failed runs.
    If fs.FolderExists(backupFolder) Then
        fs.DeleteFolder backupFolder, True
    End If
    indexPath = runDir & SOURCE_BACKUP_INDEX_RELATIVE_PATH
    If fs.FileExists(indexPath) Then
        fs.DeleteFile indexPath, True
    End If
    If fs.FolderExists(backupRoot) Then
        If fs.GetFolder(backupRoot).SubFolders.Count = 0 And fs.GetFolder(backupRoot).Files.Count = 0 Then
            fs.DeleteFolder backupRoot, True
        End If
    End If
    Log "SOURCE temporary backup removed after snapshot validation and source restoration"
    Exit Sub
CleanupFailed:
    errorText = Err.Description
    On Error Resume Next
    Warn "Source backup cleanup incomplete; generation will continue: " & errorText
End Sub

