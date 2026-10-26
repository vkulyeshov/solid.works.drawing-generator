Option Explicit

' Named constants: keep API codes, wire values and physical dimensions here.

' Polling and recovery
Const HEARTBEAT_INTERVAL_SECONDS = 15
Const PROGRESS_POLL_MILLISECONDS = 500
Const SHUTDOWN_GRACE_SECONDS = 10

' API values and serialized field indices
Const MAX_POSITIONAL_ARGUMENTS = 3
Const MAX_RETRIES_WITHOUT_PROGRESS = 2
Const EXIT_SUCCESS = 0
Const EXIT_LAUNCHER_FAILURE = 1
Const EXIT_GENERATION_FAILURE = 2
Const PERCENT_COMPLETE = 100
Const PERCENT_PENDING_MAX = 99
Const FILE_FOR_READING = 1
Const FILE_FOR_APPENDING = 8
Const FILE_UNICODE = -1
Const SPECIAL_FOLDER_TEMP = 2
Const ADODB_TEXT = 2
Const BOM_BYTE_FF = 255
Const BOM_BYTE_FE = 254
Const ASCII_QUOTE = 34
Const SHELL_WINDOW_NORMAL = 1
Const SW_DOC_PART = 1
Const SW_DOC_ASSEMBLY = 2
Const PROGRESS_LAST_FIELD = 8
Const PROGRESS_STATE_FIELD = 1
Const PROGRESS_DONE_FIELD = 2
Const PROGRESS_TOTAL_FIELD = 3
Const PROGRESS_OK_FIELD = 4
Const PROGRESS_FAILED_FIELD = 5
Const PROGRESS_WARNINGS_FIELD = 6
Const PROGRESS_OUTPUT_FIELD = 7
Const PROGRESS_MESSAGE_FIELD = 8
Const CHECKPOINT_JOB_LAST_FIELD = 15
Const LEGACY_CHECKPOINT_JOB_LAST_FIELD = 14
Const CHECKPOINT_ATTEMPTS_FIELD = 15
Const CHECKPOINT_STATUS_FIELD = 10
Const CHECKPOINT_STEM_FIELD = 11
Const CHECKPOINT_DRAWING_BYTES_FIELD = 13
Const CHECKPOINT_PDF_BYTES_FIELD = 14

' Files and serialized protocol values
Const SETTINGS_FILENAME = "generator.ini"
Const MACRO_FILENAME = "generator.swp"
Const SOURCE_FILENAME = "generator.bas"
Const BUILDER_RELATIVE_PATH = "tools\rebuild_native.py"
Const SHARED_LOCK_FILENAME = "SolidWorks_AirDuctDrawings.lock"
Const CHECKPOINT_FILENAME = "generation-state.tsv"
Const BACKGROUND_STATE_FILENAME = "background-state.ini"
Const BUILD_REPORT_FILENAME = "generator.build.json"
Const PROGRESS_FILENAME = "progress.tsv"
Const LAUNCHER_LOG_FILENAME = "launcher.log"
Const WORKER_RESULT_FILENAME = "worker-result.ini"
Const CONTEXT_FILENAME = "context.ini"
Const SETTINGS_ARCHIVE_FILENAME = "settings-used.ini"
Const CONTROLLER_PID_FILENAME = "controller.pid"
Const CONSOLE_HOST_FILENAME = "cscript.exe"
Const DRAWINGS_DIRECTORY = "Drawings"
Const PDF_DIRECTORY = "PDF"
Const DRAWING_EXTENSION = ".slddrw"
Const PDF_EXTENSION = ".pdf"
Const ASSEMBLY_EXTENSION = "sldasm"
Const PART_EXTENSION = "sldprt"
Const DRAWING_TEMPLATE_FILE_TYPE = "drwdot"
Const DRAWING_TEMPLATE_EXTENSION = ".drwdot"
Const PAPER_A3 = "A3"
Const PAPER_A4 = "A4"
Const ORIENTATION_LANDSCAPE = "landscape"
Const ORIENTATION_PORTRAIT = "portrait"
Const CHARSET_UTF8 = "utf-8"
Const CHARSET_UTF16_LE = "unicode"
Const CHARSET_UTF16_BE = "unicodeFFFE"
Const CHARSET_BYTE_INSPECTION = "iso-8859-1"

' Progress and checkpoint wire values
Const STATE_STARTING = "STARTING"
Const STATE_RUNNING = "RUNNING"
Const STATE_DONE = "DONE"
Const STATE_FAILED = "FAILED"
Const STATE_COMPLETE = "COMPLETE"
Const STATE_PAUSED = "PAUSED"
Const CHECKPOINT_JOB_TAG = "JOB"

' Stable launcher error offsets
Const ERR_ARGUMENTS = 1
Const ERR_PATH_OR_BUILD = 2
Const ERR_BUSY = 3
Const ERR_INPUT_OPEN = 10
Const ERR_UNSAVED_INPUT = 11
Const ERR_INPUT_ACTIVATION = 12
Const ERR_RUN_MACRO = 13
Const ERR_BACKGROUND_PROCESS = 14
Const ERR_BACKGROUND_WINDOW = 15

' Files and serialized protocol values
Const OUTPUT_FOLDER_PREFIX = "DrawingPack_"
Const TEMP_SESSION_PREFIX = "GeneratorLaunch_"

' API values and serialized field indices
Const EXIT_BATCH_PAUSED = 3
Const EXE_EXTENSION_TAIL_LENGTH = 3
Const SHELL_WINDOW_HIDDEN = 0
Const SHELL_PROCESS_RUNNING = 0

' Run with cscript.exe. The worker owns COM; the controller remains responsive.
' All paths passed to COM are absolute. Source checkouts build BAS with Python 3.
Dim fso, shell, named, sessionDir, progressFile, consoleLog, lockFile, lockOwned
Dim outputDir, lastState, lastDone, lastTotal, lastOK, lastFailed, lastWarnings
Dim lastMessage, recordsSeen, exitCode, startTime, isWorker, inputFile, iniFile
Dim macroFile, scriptDir, runtimeMacro, capturedError, capturedDescription
Dim sourceFile, builderFile, buildFromSource
Dim currentStage
Dim backgroundMode, backgroundVerified, backgroundPID
Dim existingPIDAtStart, existingAliveAtEnd
Dim paperFormat, paperOrientation, drawingTemplate, templatesDirectory
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
Set named = WScript.Arguments.Named
Set consoleLog = Nothing
backgroundMode = False: backgroundVerified = False: backgroundPID = 0
existingPIDAtStart = 0: existingAliveAtEnd = False
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
lockOwned = False
exitCode = EXIT_SUCCESS
isWorker = named.Exists("worker")

If LCase(fso.GetFileName(WScript.FullName)) <> CONSOLE_HOST_FILENAME Then
    Dim forwarded, arg
    forwarded = Q(ConsoleHost()) & " //nologo " & Q(WScript.ScriptFullName)
    For Each arg In WScript.Arguments
        forwarded = forwarded & " " & Q(arg)
    Next
    shell.Run forwarded, SHELL_WINDOW_NORMAL, False
    WScript.Quit EXIT_SUCCESS
End If

On Error Resume Next
If isWorker Then
    sessionDir = fso.GetAbsolutePathName(named("worker"))
    WorkerMain
Else
    ControllerMain
End If
capturedError = Err.Number
capturedDescription = Err.Description & " (stage: " & currentStage & ")"
On Error GoTo 0
If capturedError <> 0 Then
    exitCode = EXIT_LAUNCHER_FAILURE
    If isWorker And Len(sessionDir) > 0 Then
        exitCode = WorkerErrorExitCode(capturedError)
        WriteWorkerResult exitCode, "ERROR " & capturedError & ": " & capturedDescription
    Else
        Say "ERROR " & capturedError & ": " & capturedDescription
    End If
End If
If Not isWorker Then
    On Error Resume Next
    If Not (consoleLog Is Nothing) Then
        consoleLog.Close
    End If
    ArchiveDiagnostics
    On Error GoTo 0
End If
ReleaseLock
WScript.Quit exitCode

Sub ControllerMain()
    Dim key, positional, content, context, child, command
    Dim lastHeartbeat, result, available
    currentStage = "arguments"
    If named.Exists("help") Or named.Exists("?") Then
        Usage
        Exit Sub
    End If
    For Each key In named
        Select Case LCase(key)
            Case "input", "ini", "out", "background", "reset", "format", "orientation", "template", "templates"
            Case Else
                Err.Raise vbObjectError + ERR_ARGUMENTS, , "Unknown option: /" & key & ". Use /help."
        End Select
    Next
    If named.Exists("reset") Then
        If Len(named("reset")) > 0 Then
            Err.Raise vbObjectError + ERR_ARGUMENTS, , "/reset is a flag and takes no value."
        End If
    End If
    backgroundMode = named.Exists("background")
    If backgroundMode Then
        If Len(named("background")) > 0 Then
            Err.Raise vbObjectError + ERR_ARGUMENTS, , "/background is a flag and takes no value."
        End If
    End If
    Set positional = WScript.Arguments.Unnamed
    If positional.Count > MAX_POSITIONAL_ARGUMENTS Then
        Err.Raise vbObjectError + ERR_ARGUMENTS, , "Expected at most three positional arguments. Use /help."
    End If
    inputFile = ArgumentValue("input", 0, "")
    iniFile = ArgumentValue("ini", 1, fso.BuildPath(scriptDir, SETTINGS_FILENAME))
    outputDir = ArgumentValue("out", 2, "")
    If Len(inputFile) = 0 Then
        Usage
        Err.Raise vbObjectError + ERR_ARGUMENTS, , "Input model is required."
    End If
    currentStage = "path validation"
    inputFile = AbsolutePath(inputFile)
    iniFile = AbsolutePath(iniFile)
    macroFile = fso.BuildPath(scriptDir, MACRO_FILENAME)
    sourceFile = fso.BuildPath(scriptDir, SOURCE_FILENAME)
    builderFile = fso.BuildPath(scriptDir, BUILDER_RELATIVE_PATH)
    buildFromSource = fso.FileExists(sourceFile)
    If Not fso.FileExists(inputFile) Then
        Err.Raise vbObjectError + ERR_PATH_OR_BUILD, , "Input file not found: " & inputFile
    End If
    Select Case LCase(fso.GetExtensionName(inputFile))
        Case ASSEMBLY_EXTENSION, PART_EXTENSION
        Case Else
            Err.Raise vbObjectError + ERR_PATH_OR_BUILD, , "Input must be an .SLDASM or .SLDPRT file."
    End Select
    If Not fso.FileExists(iniFile) Then
        Err.Raise vbObjectError + ERR_PATH_OR_BUILD, , "INI file not found: " & iniFile
    End If
    If buildFromSource Then
        If Not fso.FileExists(builderFile) Then
            Err.Raise vbObjectError + ERR_PATH_OR_BUILD, , "Source build tools not found: " & builderFile
        End If
    ElseIf Not fso.FileExists(macroFile) Then
        Err.Raise vbObjectError + ERR_PATH_OR_BUILD, , "Generator source or packaged macro not found beside launcher."
    End If
    If Len(outputDir) > 0 Then
        outputDir = AbsolutePath(outputDir)
        If fso.FileExists(outputDir) Then
            Err.Raise vbObjectError + ERR_PATH_OR_BUILD, , "Output path is a file: " & outputDir
        End If
    End If
    ' Reading first avoids starting SolidWorks when the configuration is unreadable.
    currentStage = "reading INI"
    content = ReadIniText(iniFile)
    ResolveDrawingSettings content
    currentStage = "session setup"
    lockFile = fso.BuildPath(fso.GetSpecialFolder(SPECIAL_FOLDER_TEMP), SHARED_LOCK_FILENAME)
    AcquireLock
    If Len(outputDir) = 0 Then
        outputDir = fso.BuildPath(fso.GetParentFolderName(inputFile), OUTPUT_FOLDER_PREFIX & TimeStamp())
    End If
    EnsureFolder outputDir
    sessionDir = fso.BuildPath(fso.GetSpecialFolder(SPECIAL_FOLDER_TEMP), TEMP_SESSION_PREFIX & fso.GetTempName())
    fso.CreateFolder sessionDir
    WriteText lockFile, sessionDir
    progressFile = fso.BuildPath(sessionDir, PROGRESS_FILENAME)
    WriteText progressFile, ""
    Set consoleLog = fso.CreateTextFile(fso.BuildPath(sessionDir, LAUNCHER_LOG_FILENAME), False, True)
    consoleLog.Close: Set consoleLog = Nothing
    runtimeMacro = fso.BuildPath(sessionDir, MACRO_FILENAME)
    If buildFromSource Then
        BuildNativeMacro
    Else
        fso.CopyFile macroFile, runtimeMacro, False
    End If
    If named.Exists("reset") Then
        If fso.FileExists(fso.BuildPath(outputDir, CHECKPOINT_FILENAME)) Then
            fso.DeleteFile fso.BuildPath(outputDir, CHECKPOINT_FILENAME), True
        End If
    End If
    ' The last occurrence of each key wins in the VBA INI reader.
    content = content & vbCrLf & "; Per-run launcher overrides" & vbCrLf _
        & "LauncherInput=" & inputFile & vbCrLf _
        & "LauncherOutput=" & outputDir & vbCrLf _
        & "LauncherProgress=" & progressFile & vbCrLf _
        & "LauncherSettingsDirectory=" & fso.GetParentFolderName(iniFile) & vbCrLf _
        & "LauncherScriptDirectory=" & scriptDir & vbCrLf _
        & "PaperFormat=" & paperFormat & vbCrLf & "PaperOrientation=" & paperOrientation & vbCrLf _
        & "TemplatesDirectory=" & templatesDirectory & vbCrLf & "DrawingTemplate=" & drawingTemplate & vbCrLf _
        & "LauncherBackground=0" & vbCrLf & "LauncherForbiddenPID=0" & vbCrLf _
        & "ShowSummary=0" & vbCrLf
    WriteText fso.BuildPath(sessionDir, SETTINGS_FILENAME), content
    Set context = fso.CreateTextFile(fso.BuildPath(sessionDir, CONTEXT_FILENAME), False, True)
    context.WriteLine "Input=" & inputFile
    context.WriteLine "Output=" & outputDir
    context.WriteLine "Ini=" & iniFile
    context.WriteLine "Macro=" & runtimeMacro
    context.WriteLine "Lock=" & lockFile
    If backgroundMode Then
        context.WriteLine "Background=1"
    End If
    context.Close
    lastState = STATE_STARTING: lastDone = 0: lastTotal = 0: lastOK = 0
    lastFailed = 0: lastWarnings = 0: recordsSeen = 0
    lastMessage = "Starting worker"
    startTime = Now: lastHeartbeat = Now
    Say "Input: " & inputFile
    Say "INI:   " & iniFile
    Say "Paper: " & paperFormat & "; template: " & drawingTemplate
    If backgroundMode Then
        Say "Mode: background (separate SolidWorks instance)"
    End If
    If Len(outputDir) > 0 Then
        Say "Output: " & outputDir
    End If
    Say "Diagnostics: " & sessionDir
    command = Q(ConsoleHost()) & " //nologo " & Q(WScript.ScriptFullName) & " " & Q("/worker:" & sessionDir)
    currentStage = "worker and progress monitor"
    Dim beforeDone, afterDone, beforeAttempts, afterAttempts, retries, batchNumber
    retries = 0: batchNumber = 0
    Do
        beforeDone = CheckpointDone(outputDir)
        beforeAttempts = CheckpointAttempts(outputDir)
        batchNumber = batchNumber + 1
        Say "Batch " & batchNumber & "; checkpoint completed=" & beforeDone & "; attempts=" & beforeAttempts
        WriteText progressFile, "": recordsSeen = 0: lastState = STATE_STARTING
        If fso.FileExists(fso.BuildPath(sessionDir, WORKER_RESULT_FILENAME)) Then
            fso.DeleteFile fso.BuildPath(sessionDir, WORKER_RESULT_FILENAME), True
        End If
        If fso.FileExists(fso.BuildPath(sessionDir, BACKGROUND_STATE_FILENAME)) Then
            fso.DeleteFile fso.BuildPath(sessionDir, BACKGROUND_STATE_FILENAME), True
        End If
        Set child = shell.Exec(command)
        Do
            PollProgress
            If backgroundMode Then
                FinishLingeringBackground
            End If
            If child.Status <> SHELL_PROCESS_RUNNING Then
                Exit Do
            End If
            If DateDiff("s", lastHeartbeat, Now) >= HEARTBEAT_INTERVAL_SECONDS Then
                Say "[running, " & DateDiff("s", startTime, Now) & "s] " & lastMessage
                lastHeartbeat = Now
            End If
            WScript.Sleep PROGRESS_POLL_MILLISECONDS
        Loop
        PollProgress
        available = child.StdOut.ReadAll
        If Len(Trim(available)) > 0 Then
            Say Trim(available)
        End If
        available = child.StdErr.ReadAll
        If Len(Trim(available)) > 0 Then
            Say Trim(available)
        End If
        Set result = ReadKeyValues(fso.BuildPath(sessionDir, WORKER_RESULT_FILENAME))
        exitCode = EXIT_LAUNCHER_FAILURE
        If result.Exists("ExitCode") Then
            exitCode = CLng(result("ExitCode"))
            Say result("Message")
        Else
            Say "ERROR: worker ended without a result; process exit code=" & child.ExitCode
        End If
        If child.ExitCode <> 0 And exitCode = EXIT_SUCCESS Then
            exitCode = child.ExitCode
        End If
        If lastState <> STATE_COMPLETE And exitCode = EXIT_SUCCESS Then
            exitCode = EXIT_GENERATION_FAILURE
        End If
        afterDone = CheckpointDone(outputDir)
        afterAttempts = CheckpointAttempts(outputDir)
        ArchiveDiagnostics
        If exitCode = EXIT_BATCH_PAUSED Then
            If afterDone <= beforeDone And afterAttempts <= beforeAttempts Then
                exitCode = EXIT_GENERATION_FAILURE
                Say "ERROR: batch made no checkpoint progress."
                Exit Do
            End If
            retries = 0
        ElseIf exitCode = EXIT_LAUNCHER_FAILURE And fso.FileExists(fso.BuildPath(outputDir, CHECKPOINT_FILENAME)) Then
            If afterDone > beforeDone Or afterAttempts > beforeAttempts Then
                retries = 0
            End If
            retries = retries + 1
            If retries > MAX_RETRIES_WITHOUT_PROGRESS Then
                Say "ERROR: repeated crash without progress. Resume with the same /out folder after fixing the model."
                Exit Do
            End If
            Say "Recovering interrupted job from checkpoint; retry " & retries & "/" & MAX_RETRIES_WITHOUT_PROGRESS
        Else
            Exit Do
        End If
    Loop
    If Len(outputDir) > 0 Then
        Say "Artifacts: " & outputDir
    End If
    Say "Exit code: " & exitCode
End Sub

Sub BuildNativeMacro()
    Dim pythonPath, command, child, details, errorNumber, errorText
    currentStage = "building native VBA macro"
    Say "Building generator.bas into the external launcher session"
    pythonPath = shell.Environment("PROCESS")("SW_DRAWING_PYTHON")
    If Len(pythonPath) > 0 Then
        command = Q(pythonPath)
    Else
        command = "py -3"
    End If
    command = command & " -B " & Q(builderFile) & " --source " & Q(sourceFile) & " --output " & Q(runtimeMacro)
    On Error Resume Next
    Set child = shell.Exec(command)
    errorNumber = Err.Number: errorText = Err.Description
    Err.Clear
    On Error GoTo 0
    If errorNumber <> 0 Then
        Err.Raise vbObjectError + ERR_PATH_OR_BUILD, , "Python 3 is required to build generator.bas. Install Python with the Windows py launcher, or set SW_DRAWING_PYTHON to python.exe. " & errorText
    End If
    details = child.StdOut.ReadAll & child.StdErr.ReadAll
    If child.ExitCode <> 0 Then
        Err.Raise vbObjectError + ERR_PATH_OR_BUILD, , "VBA build failed: " & CleanLine(details)
    End If
    If Not fso.FileExists(runtimeMacro) Then
        Err.Raise vbObjectError + ERR_PATH_OR_BUILD, , "VBA build did not produce generator.swp."
    End If
    Say "VBA build completed"
End Sub

Sub ResolveDrawingSettings(ByVal content)
    Dim iniDirectory, configuredTemplate, configuredDirectory
    iniDirectory = fso.GetParentFolderName(iniFile)
    paperFormat = UCase(IniSetting(content, "PaperFormat", PAPER_A3))
    paperOrientation = LCase(IniSetting(content, "PaperOrientation", ORIENTATION_LANDSCAPE))
    If named.Exists("format") Then
        paperFormat = UCase(Trim(named("format")))
    End If
    If named.Exists("orientation") Then
        paperOrientation = LCase(Trim(named("orientation")))
    End If
    If paperFormat <> PAPER_A4 And paperFormat <> PAPER_A3 Then
        Err.Raise vbObjectError + ERR_PATH_OR_BUILD, , "Format must be A4 or A3."
    End If
    If paperOrientation <> ORIENTATION_LANDSCAPE And paperOrientation <> ORIENTATION_PORTRAIT Then
        Err.Raise vbObjectError + ERR_PATH_OR_BUILD, , "Orientation must be landscape or portrait."
    End If
    configuredDirectory = IniSetting(content, "TemplatesDirectory", "")
    templatesDirectory = fso.BuildPath(scriptDir, "templates")
    If Len(configuredDirectory) > 0 Then
        templatesDirectory = SettingsPath(configuredDirectory, iniDirectory)
    End If
    If named.Exists("templates") Then
        If Len(named("templates")) = 0 Then
            Err.Raise vbObjectError + ERR_PATH_OR_BUILD, , "Templates directory is required after /templates:."
        End If
        templatesDirectory = AbsolutePath(named("templates"))
    End If
    configuredTemplate = IniSetting(content, "DrawingTemplate", "")
    drawingTemplate = ""
    If Len(configuredTemplate) > 0 Then
        drawingTemplate = SettingsPath(configuredTemplate, iniDirectory)
    End If
    If named.Exists("template") Then
        If Len(named("template")) = 0 Then
            Err.Raise vbObjectError + ERR_PATH_OR_BUILD, , "Template file is required after /template:."
        End If
        drawingTemplate = AbsolutePath(named("template"))
    End If
    If drawingTemplate = "" Then
        drawingTemplate = fso.BuildPath(templatesDirectory, paperFormat & "." & paperOrientation & DRAWING_TEMPLATE_EXTENSION)
    End If
    If LCase(fso.GetExtensionName(drawingTemplate)) <> DRAWING_TEMPLATE_FILE_TYPE Then
        Err.Raise vbObjectError + ERR_PATH_OR_BUILD, , "Drawing template must be a .drwdot file."
    End If
    If Not fso.FileExists(drawingTemplate) Then
        Err.Raise vbObjectError + ERR_PATH_OR_BUILD, , "Drawing template not found: " & drawingTemplate
    End If
End Sub

Function SettingsPath(ByVal value, ByVal directory)
    If InStr(value, ":") = 0 And Left(value, 2) <> "\\" Then
        value = fso.BuildPath(directory, value)
    End If
    SettingsPath = AbsolutePath(value)
End Function

Function IniSetting(ByVal content, ByVal key, ByVal fallback)
    Dim line, equals
    IniSetting = fallback
    For Each line In Split(Replace(content, vbCr, ""), vbLf)
        line = Trim(line)
        If Len(line) > 0 Then
            If Left(line, 1) <> ";" And Left(line, 1) <> "[" Then
                equals = InStr(line, "=")
                If equals > 1 Then
                    If StrComp(Trim(Left(line, equals - 1)), key, vbTextCompare) = 0 Then
                        IniSetting = Trim(Mid(line, equals + 1))
                    End If
                End If
            End If
        End If
    Next
End Function

Sub FinishLingeringBackground()
    Dim statePath, state, service, processes, process, pid
    If lastState <> STATE_COMPLETE And lastState <> STATE_PAUSED And lastState <> STATE_FAILED Then
        Exit Sub
    End If
    statePath = fso.BuildPath(sessionDir, BACKGROUND_STATE_FILENAME)
    If Not fso.FileExists(statePath) Then
        Exit Sub
    End If
    If DateDiff("s", fso.GetFile(statePath).DateLastModified, Now) < SHUTDOWN_GRACE_SECONDS Then
        Exit Sub
    End If
    Set state = ReadKeyValues(statePath)
    If Not state.Exists("ReadyToExit") Or Not state.Exists("PID") Then
        Exit Sub
    End If
    If state("ReadyToExit") <> "1" Then
        Exit Sub
    End If
    pid = CLng(state("PID"))
    On Error Resume Next
    Set service = GetObject("winmgmts:\\.\root\cimv2")
    Set processes = service.ExecQuery("Select * from Win32_Process Where Name='SLDWORKS.exe' And ProcessId=" & pid)
    For Each process In processes
        If Not IsNull(process.CommandLine) Then
            If InStr(1, process.CommandLine, runtimeMacro, vbTextCompare) > 0 Then
                ' The macro has committed artifacts, closed documents and restored its window.
                ' Only this launcher-owned OLE server is eligible; user sessions never match.
                process.Terminate 0
                Say "Closed completed background OLE server PID=" & pid
            End If
        End If
    Next
    Err.Clear
    On Error GoTo 0
End Sub

Sub WorkerMain()
    Dim context, sw, doc, active, spec, docType, ok
    currentStage = "worker context"
    Set context = ReadKeyValues(fso.BuildPath(sessionDir, CONTEXT_FILENAME))
    RecordControllerPID
    inputFile = context("Input")
    If context.Exists("Output") Then
        outputDir = context("Output")
    End If
    iniFile = context("Ini")
    runtimeMacro = context("Macro")
    lockFile = context("Lock")
    lockOwned = False ' Only the controller owns the global lock across batches.
    progressFile = fso.BuildPath(sessionDir, PROGRESS_FILENAME)
    backgroundMode = context.Exists("Background")
    currentStage = "connecting to SolidWorks"
    Set sw = Nothing
    If backgroundMode Then
        RunBackgroundWorker
        Exit Sub
    End If
    WorkerProgress "Connecting to SolidWorks"
    On Error Resume Next
    Set sw = GetObject(, "SldWorks.Application")
    Err.Clear
    On Error GoTo 0
    If sw Is Nothing Then
        WorkerProgress "Starting SolidWorks"
        Set sw = CreateObject("SldWorks.Application")
    End If
    sw.Visible = True
    sw.UserControl = True
    WorkerProgress "Opening input model"
    currentStage = "opening input model"
    sw.SetCurrentWorkingDirectory fso.GetParentFolderName(inputFile)
    currentStage = "GetOpenDocumentByName"
    Set doc = sw.GetOpenDocumentByName(inputFile)
    If doc Is Nothing Then
        docType = SW_DOC_PART
        If LCase(fso.GetExtensionName(inputFile)) = ASSEMBLY_EXTENSION Then
            docType = SW_DOC_ASSEMBLY
        End If
        currentStage = "OpenDoc7"
        ' Specification properties avoid typed ByRef Long outputs unsupported by VBScript.
        Set spec = sw.GetOpenDocSpec(inputFile)
        spec.DocumentType = CLng(docType)
        spec.Silent = True: spec.ReadOnly = True
        Set doc = sw.OpenDoc7(spec)
        If doc Is Nothing Then
            Err.Raise vbObjectError + ERR_INPUT_OPEN, , "OpenDoc7 failed: error=" & spec.Error & ", warnings=" & spec.Warning
        End If
    End If
    currentStage = "checking active model"
    If doc.GetSaveFlag Then
        If fso.FileExists(fso.BuildPath(outputDir, CHECKPOINT_FILENAME)) Then
            WorkerProgress "Input has unsaved changes; resuming saved snapshot without saving input"
        Else
            Err.Raise vbObjectError + ERR_UNSAVED_INPUT, , "Input has unsaved changes. Save or discard them in SolidWorks before creating a new snapshot."
        End If
    End If
    currentStage = "making model visible"
    doc.Visible = True
    currentStage = "ActivateDoc"
    Set active = sw.ActivateDoc(fso.GetFileName(inputFile))
    Set active = sw.ActiveDoc
    If active Is Nothing Then
        Err.Raise vbObjectError + ERR_INPUT_ACTIVATION, , "Input model did not become active."
    End If
    If StrComp(active.GetPathName, inputFile, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + ERR_INPUT_ACTIVATION, , "Another model is active; generation was not started."
    End If
    WorkerProgress "Running VBA drawing generator"
    currentStage = "RunMacro"
    ' RunMacro has no typed ByRef outputs. VBA reports detailed success/failure itself.
    ok = sw.RunMacro(runtimeMacro, "Macro11", "main")
    If ok = False Then
        Err.Raise vbObjectError + ERR_RUN_MACRO, , "RunMacro failed. See SolidWorks and launcher diagnostics."
    End If
    CompleteFromProgress
End Sub

Function WorkerErrorExitCode(ByVal errorNumber)
    ' Explicit model/macro validation failures cannot recover through a crash retry.
    WorkerErrorExitCode = EXIT_LAUNCHER_FAILURE
    Select Case errorNumber
        Case vbObjectError + ERR_INPUT_OPEN, vbObjectError + ERR_UNSAVED_INPUT, vbObjectError + ERR_INPUT_ACTIVATION, vbObjectError + ERR_RUN_MACRO, vbObjectError + ERR_BACKGROUND_WINDOW
            WorkerErrorExitCode = EXIT_GENERATION_FAILURE
    End Select
End Function

Sub RunBackgroundWorker()
    Dim existing, existingPID, clsid, server, exePath, runtimeIni, content, processCode, state
    currentStage = "background startup"
    Set existing = Nothing: existingPID = 0
    On Error Resume Next
    Set existing = GetObject(, "SldWorks.Application")
    If Not (existing Is Nothing) Then
        existingPID = existing.GetProcessID
    End If
    Err.Clear
    On Error GoTo 0
    existingPIDAtStart = existingPID
    clsid = shell.RegRead("HKCR\SldWorks.Application\CLSID\")
    server = Trim(shell.RegRead("HKCR\CLSID\" & clsid & "\LocalServer32\"))
    If Left(server, 1) = Chr(ASCII_QUOTE) Then
        exePath = Mid(server, 2, InStr(2, server, Chr(ASCII_QUOTE)) - 2)
    Else
        exePath = Left(server, InStr(1, server, ".exe", vbTextCompare) + EXE_EXTENSION_TAIL_LENGTH)
    End If
    If Not fso.FileExists(exePath) Then
        Err.Raise vbObjectError + ERR_BACKGROUND_PROCESS, , "Registered SolidWorks executable was not found: " & exePath
    End If
    runtimeIni = fso.BuildPath(sessionDir, SETTINGS_FILENAME)
    content = ReadUnicode(runtimeIni) & vbCrLf & "LauncherBackground=1" & vbCrLf & "LauncherForbiddenPID=" & existingPID & vbCrLf
    WriteText runtimeIni, content
    WorkerProgress "Starting SolidWorks /b /m in a separate process"
    ' /b is the supported background startup switch. Window style 0 also hides startup UI.
    processCode = shell.Run(Q(exePath) & " /b /m " & Q(runtimeMacro), SHELL_WINDOW_HIDDEN, True)
    If Not (existing Is Nothing) Then
        On Error Resume Next
        existingAliveAtEnd = (existing.GetProcessID = existingPID)
        Err.Clear
        On Error GoTo 0
    End If
    Set state = ReadKeyValues(fso.BuildPath(sessionDir, BACKGROUND_STATE_FILENAME))
    If state.Exists("PID") Then
        backgroundPID = CLng(state("PID"))
    End If
    If state.Exists("Hidden") Then
        backgroundVerified = (LCase(state("Hidden")) = "true")
    End If
    If state.Exists("Offscreen") Then
        backgroundVerified = backgroundVerified Or (LCase(state("Offscreen")) = "true")
    End If
    If backgroundPID = 0 Or backgroundPID = existingPID Then
        Err.Raise vbObjectError + ERR_BACKGROUND_PROCESS, , "SolidWorks ended before shutdown confirmation; process exit code=" & processCode & ". Resume from generation-state.tsv."
    End If
    If backgroundVerified = False Then
        Err.Raise vbObjectError + ERR_BACKGROUND_WINDOW, , "The macro did not confirm a background window outside the screen."
    End If
    CompleteFromProgress
    If processCode <> 0 And exitCode = EXIT_SUCCESS Then
        exitCode = EXIT_LAUNCHER_FAILURE
        WriteWorkerResult exitCode, "Background SolidWorks exited with code " & processCode
    End If
End Sub

Sub CompleteFromProgress()
    Dim events, rows, fields, line, terminal, message
    terminal = "": message = "VBA did not report a completion status."
    events = ReadUnicode(progressFile)
    rows = Split(events, vbCrLf)
    For Each line In rows
        fields = Split(line, vbTab)
        If UBound(fields) >= PROGRESS_LAST_FIELD Then
            If fields(PROGRESS_STATE_FIELD) = STATE_COMPLETE Or fields(PROGRESS_STATE_FIELD) = STATE_FAILED Or fields(PROGRESS_STATE_FIELD) = STATE_PAUSED Then
                terminal = fields(PROGRESS_STATE_FIELD): message = fields(PROGRESS_MESSAGE_FIELD)
            End If
        End If
    Next
    exitCode = EXIT_GENERATION_FAILURE
    If terminal = STATE_COMPLETE Then
        exitCode = EXIT_SUCCESS
    End If
    If terminal = STATE_PAUSED Then
        exitCode = EXIT_BATCH_PAUSED
    End If
    WriteWorkerResult exitCode, message
End Sub

Sub PollProgress()
    Dim text, lines, i, fields, percentage
    On Error Resume Next
    text = ReadUnicode(progressFile)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    lines = Split(text, vbCrLf)
    ' An incomplete last record is retried on the next poll.
    For i = recordsSeen To UBound(lines) - 1
        fields = Split(lines(i), vbTab)
        recordsSeen = i + 1
        If UBound(fields) >= PROGRESS_LAST_FIELD Then
            lastState = fields(PROGRESS_STATE_FIELD)
            If CLng(fields(PROGRESS_TOTAL_FIELD)) > 0 Or lastTotal = 0 Then
                lastDone = CLng(fields(PROGRESS_DONE_FIELD)): lastTotal = CLng(fields(PROGRESS_TOTAL_FIELD))
                lastOK = CLng(fields(PROGRESS_OK_FIELD)): lastFailed = CLng(fields(PROGRESS_FAILED_FIELD)): lastWarnings = CLng(fields(PROGRESS_WARNINGS_FIELD))
            End If
            If Len(fields(PROGRESS_OUTPUT_FIELD)) > 0 Then
                outputDir = fields(PROGRESS_OUTPUT_FIELD)
            End If
            lastMessage = fields(PROGRESS_MESSAGE_FIELD)
            percentage = "--"
            If lastTotal > 0 Then
                percentage = Fix(PERCENT_COMPLETE * lastDone / lastTotal)
                If lastState <> STATE_COMPLETE And percentage >= PERCENT_COMPLETE Then
                    percentage = PERCENT_PENDING_MAX
                End If
                If lastState = STATE_COMPLETE Then
                    percentage = PERCENT_COMPLETE
                End If
            End If
            Say "[" & percentage & "% " & lastDone & "/" & lastTotal & "] " & lastMessage
        End If
    Next
End Sub

Sub WorkerProgress(ByVal message)
    Dim stream, done, total
    done = 0: total = 0
    If Len(outputDir) > 0 Then
        done = CheckpointHandled(outputDir): total = CheckpointTotal(outputDir)
    End If
    Set stream = fso.OpenTextFile(progressFile, FILE_FOR_APPENDING, True, FILE_UNICODE)
    stream.WriteLine Time & vbTab & STATE_STARTING & vbTab & done & vbTab & total & vbTab & done & vbTab & "0" & vbTab & "0" & vbTab & outputDir & vbTab & message
    stream.Close
End Sub

Sub WriteWorkerResult(ByVal code, ByVal message)
    On Error Resume Next
    WriteText fso.BuildPath(sessionDir, WORKER_RESULT_FILENAME), "ExitCode=" & code & vbCrLf & "Message=" & CleanLine(message) & vbCrLf _
        & "Background=" & CStr(backgroundMode) & vbCrLf & "BackgroundPID=" & backgroundPID & vbCrLf & "BackgroundHiddenAtEnd=" & CStr(backgroundVerified) & vbCrLf
    Dim details
    Set details = fso.OpenTextFile(fso.BuildPath(sessionDir, WORKER_RESULT_FILENAME), FILE_FOR_APPENDING, True, FILE_UNICODE)
    details.WriteLine "ExistingPID=" & existingPIDAtStart
    details.WriteLine "ExistingProcessStillAlive=" & Abs(CInt(existingAliveAtEnd))
    details.Close
    On Error GoTo 0
End Sub

Sub AcquireLock()
    Dim stream, errorNumber
    RecoverStaleLock
    On Error Resume Next
    Set stream = fso.CreateTextFile(lockFile, False, True)
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    If errorNumber <> 0 Then
        Err.Raise vbObjectError + ERR_BUSY, , "Another launcher is running, or a previous run was interrupted. Lock: " & lockFile & ". Check SolidWorks before removing a stale lock."
    End If
    stream.Close
    lockOwned = True
End Sub

Sub RecordControllerPID()
    ' The unique /worker:session command identifies this worker without touching other jobs.
    Dim service, processes, process
    On Error Resume Next
    Set service = GetObject("winmgmts:\\.\root\cimv2")
    Set processes = service.ExecQuery("Select ParentProcessId, CommandLine from Win32_Process Where Name='cscript.exe'")
    For Each process In processes
        If Not IsNull(process.CommandLine) Then
            If InStr(1, process.CommandLine, "/worker:" & sessionDir, vbTextCompare) > 0 Then
                WriteText fso.BuildPath(sessionDir, CONTROLLER_PID_FILENAME), CStr(process.ParentProcessId)
                Exit For
            End If
        End If
    Next
    Err.Clear
    On Error GoTo 0
End Sub

Sub RecoverStaleLock()
    Dim previousSession, ownerFile, ownerPID, service, processes, process, alive
    If Not fso.FileExists(lockFile) Then
        Exit Sub
    End If
    On Error Resume Next
    previousSession = ReadUnicode(lockFile)
    ownerFile = fso.BuildPath(previousSession, CONTROLLER_PID_FILENAME)
    If Not fso.FileExists(ownerFile) Then
        Exit Sub
    End If
    ownerPID = Trim(ReadUnicode(ownerFile))
    If Not IsNumeric(ownerPID) Then
        Exit Sub
    End If
    Set service = GetObject("winmgmts:\\.\root\cimv2")
    Set processes = service.ExecQuery("Select CommandLine from Win32_Process Where ProcessId=" & CLng(ownerPID))
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    alive = False
    For Each process In processes
        If Not IsNull(process.CommandLine) Then
            If InStr(1, process.CommandLine, fso.GetFileName(WScript.ScriptFullName), vbTextCompare) > 0 Then
                alive = True
            End If
        End If
    Next
    ' Do not break the lock while the former worker/native macro may still be running.
    Set processes = service.ExecQuery("Select CommandLine from Win32_Process Where Name='cscript.exe' Or Name='SLDWORKS.exe'")
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    For Each process In processes
        If Not IsNull(process.CommandLine) Then
            If InStr(1, process.CommandLine, previousSession, vbTextCompare) > 0 Then
                alive = True
            End If
        End If
    Next
    If Not alive Then
        If ReadUnicode(lockFile) = previousSession Then
            fso.DeleteFile lockFile, True
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Sub ReleaseLock()
    On Error Resume Next
    If lockOwned And Len(lockFile) > 0 Then
        If fso.FileExists(lockFile) Then
            If ReadUnicode(lockFile) = sessionDir Then
                fso.DeleteFile lockFile, True
            End If
        End If
    End If
    On Error GoTo 0
End Sub

Sub ArchiveDiagnostics()
    Dim filename
    If Len(sessionDir) = 0 Or Len(outputDir) = 0 Then
        Exit Sub
    End If
    If Not fso.FolderExists(outputDir) Then
        Exit Sub
    End If
    For Each filename In Array(LAUNCHER_LOG_FILENAME, PROGRESS_FILENAME, WORKER_RESULT_FILENAME, BACKGROUND_STATE_FILENAME, BUILD_REPORT_FILENAME)
        If fso.FileExists(fso.BuildPath(sessionDir, filename)) Then
            fso.CopyFile fso.BuildPath(sessionDir, filename), fso.BuildPath(outputDir, filename), True
        End If
    Next
    If Len(iniFile) > 0 Then
        fso.CopyFile iniFile, fso.BuildPath(outputDir, SETTINGS_ARCHIVE_FILENAME), True
    End If
End Sub

Function ArgumentValue(ByVal key, ByVal index, ByVal fallback)
    ArgumentValue = fallback
    If WScript.Arguments.Unnamed.Count > index Then
        ArgumentValue = WScript.Arguments.Unnamed(index)
    End If
    If named.Exists(key) Then
        If WScript.Arguments.Unnamed.Count > index Then
            Err.Raise vbObjectError + ERR_ARGUMENTS, , "Do not specify " & key & " both by name and by position."
        End If
        ArgumentValue = named(key)
    End If
End Function

Function AbsolutePath(ByVal path)
    If InStr(path, vbCr) > 0 Or InStr(path, vbLf) > 0 Or InStr(path, vbTab) > 0 Or InStr(path, Chr(ASCII_QUOTE)) > 0 Then
        Err.Raise vbObjectError + ERR_ARGUMENTS, , "Invalid characters in path."
    End If
    AbsolutePath = fso.GetAbsolutePathName(path)
End Function

Function Q(ByVal value)
    If InStr(value, Chr(ASCII_QUOTE)) > 0 Then
        Err.Raise vbObjectError + ERR_ARGUMENTS, , "A path contains a quotation mark."
    End If
    Q = Chr(ASCII_QUOTE) & value & Chr(ASCII_QUOTE)
End Function

Function ConsoleHost()
    ConsoleHost = fso.BuildPath(fso.GetParentFolderName(WScript.FullName), CONSOLE_HOST_FILENAME)
End Function

Function ReadUnicode(ByVal path)
    Dim stream
    Set stream = fso.OpenTextFile(path, FILE_FOR_READING, False, FILE_UNICODE)
    ReadUnicode = stream.ReadAll
    stream.Close
End Function

Function ReadIniText(ByVal path)
    Dim stream, leading, charset
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = ADODB_TEXT: stream.Charset = CHARSET_BYTE_INSPECTION
    stream.Open: stream.LoadFromFile path
    leading = stream.ReadText(2)
    charset = CHARSET_UTF8
    If Len(leading) >= 2 Then
        If AscW(Left(leading, 1)) = BOM_BYTE_FF And AscW(Mid(leading, 2, 1)) = BOM_BYTE_FE Then
            charset = CHARSET_UTF16_LE
        End If
        If AscW(Left(leading, 1)) = BOM_BYTE_FE And AscW(Mid(leading, 2, 1)) = BOM_BYTE_FF Then
            charset = CHARSET_UTF16_BE
        End If
    End If
    stream.Position = 0: stream.Charset = charset
    ReadIniText = stream.ReadText
    stream.Close
End Function

Function ReadKeyValues(ByVal path)
    Dim dict, line, i
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    If fso.FileExists(path) Then
        For Each line In Split(ReadUnicode(path), vbCrLf)
            i = InStr(line, "=")
            If i > 1 Then
                dict(Left(line, i - 1)) = Mid(line, i + 1)
            End If
        Next
    End If
    Set ReadKeyValues = dict
End Function

Sub WriteText(ByVal path, ByVal content)
    Dim stream
    Set stream = fso.CreateTextFile(path, True, True)
    stream.Write content
    stream.Close
End Sub

Function CleanLine(ByVal value)
    CleanLine = Replace(Replace(Replace(value, vbCr, " "), vbLf, " "), vbTab, " ")
End Function

Sub Say(ByVal text)
    WScript.Echo text
    If Len(sessionDir) > 0 Then
        Set consoleLog = fso.OpenTextFile(fso.BuildPath(sessionDir, LAUNCHER_LOG_FILENAME), FILE_FOR_APPENDING, True, FILE_UNICODE)
        consoleLog.WriteLine Time & " " & text
        consoleLog.Close: Set consoleLog = Nothing
    End If
End Sub

Sub Usage()
    WScript.Echo "SolidWorks drawing generator launcher"
    WScript.Echo "cscript //nologo launcher.vbs MODEL [INI] [OUTPUT]"
    WScript.Echo "cscript //nologo launcher.vbs /input:MODEL [/ini:INI] [/out:OUTPUT]"
    WScript.Echo "Drawing options: /format:A4|A3 /orientation:landscape|portrait /template:FILE /templates:DIRECTORY"
    WScript.Echo "Defaults: A3 landscape, templates\A3.landscape.drwdot beside this script. Explicit template controls orientation."
    WScript.Echo "Add /background to keep a separate SolidWorks window outside the screen and exit when finished."
    WScript.Echo "Quote arguments containing spaces. Default INI is beside this script."
    WScript.Echo "Default output: DrawingPack_YYYYMMDD_HHMMSS beside the input model."
    WScript.Echo "Source checkout: generator.bas is built outside Git using Python 3 (py -3 or SW_DRAWING_PYTHON)."
    WScript.Echo "Prebuilt runtime kits run generator.swp without Python."
    WScript.Echo "Existing output resumes generation-state.tsv. Delete that file, or add /reset, to regenerate all jobs."
    WScript.Echo "JobsPerBatch in INI limits one batch (default 8). Background batches restart SolidWorks automatically."
    WScript.Echo "Failed jobs do not stop later batches. MaxJobAttempts in INI limits persistent attempts per job (default 3)."
    WScript.Echo "Permanent failures are listed in failed-jobs.tsv; a finished queue with failures exits with code 2."
    WScript.Echo "Exit code 0 means complete success; 1 launcher/crash; 2 model validation, macro or drawing failure."
End Sub

Function TimeStamp()
    Dim d
    d = Now
    TimeStamp = Year(d) & Right("0" & Month(d), 2) & Right("0" & Day(d), 2) & "_" & Right("0" & Hour(d), 2) & Right("0" & Minute(d), 2) & Right("0" & Second(d), 2)
End Function

Sub EnsureFolder(ByVal path)
    If fso.FolderExists(path) Then
        Exit Sub
    End If
    EnsureFolder fso.GetParentFolderName(path)
    fso.CreateFolder path
End Sub

Function CheckpointDone(ByVal path)
    Dim stateFile, line, cells
    CheckpointDone = 0
    stateFile = fso.BuildPath(path, CHECKPOINT_FILENAME)
    If Not fso.FileExists(stateFile) Then
        Exit Function
    End If
    For Each line In Split(ReadUnicode(stateFile), vbCrLf)
        cells = Split(line, vbTab)
        If IsCheckpointJobRow(cells) Then
            If cells(0) = CHECKPOINT_JOB_TAG And cells(CHECKPOINT_STATUS_FIELD) = STATE_DONE Then
                If CheckpointArtifactsValid(path, cells) Then
                    CheckpointDone = CheckpointDone + 1
                End If
            End If
        End If
    Next
End Function

Function CheckpointArtifactsValid(ByVal path, ByVal cells)
    Dim nativePath, pdfPath
    CheckpointArtifactsValid = False
    If fso.GetFileName(cells(CHECKPOINT_STEM_FIELD)) <> cells(CHECKPOINT_STEM_FIELD) Then
        Exit Function
    End If
    nativePath = fso.BuildPath(fso.BuildPath(path, DRAWINGS_DIRECTORY), cells(CHECKPOINT_STEM_FIELD) & DRAWING_EXTENSION)
    pdfPath = fso.BuildPath(fso.BuildPath(path, PDF_DIRECTORY), cells(CHECKPOINT_STEM_FIELD) & PDF_EXTENSION)
    If Not fso.FileExists(nativePath) Or Not fso.FileExists(pdfPath) Then
        Exit Function
    End If
    CheckpointArtifactsValid = CStr(fso.GetFile(nativePath).Size) = cells(CHECKPOINT_DRAWING_BYTES_FIELD) And CStr(fso.GetFile(pdfPath).Size) = cells(CHECKPOINT_PDF_BYTES_FIELD) And CDbl(cells(CHECKPOINT_PDF_BYTES_FIELD)) > 0
End Function

Function CheckpointTotal(ByVal path)
    Dim stateFile, line, cells
    CheckpointTotal = 0
    stateFile = fso.BuildPath(path, CHECKPOINT_FILENAME)
    If Not fso.FileExists(stateFile) Then
        Exit Function
    End If
    For Each line In Split(ReadUnicode(stateFile), vbCrLf)
        cells = Split(line, vbTab)
        If IsCheckpointJobRow(cells) Then
            If cells(0) = CHECKPOINT_JOB_TAG Then
                CheckpointTotal = CheckpointTotal + 1
            End If
        End If
    Next
End Function

Function IsCheckpointJobRow(ByVal cells)
    Dim lastField
    lastField = UBound(cells)
    IsCheckpointJobRow = False
    If lastField = CHECKPOINT_JOB_LAST_FIELD Or lastField = LEGACY_CHECKPOINT_JOB_LAST_FIELD Then
        IsCheckpointJobRow = (cells(0) = CHECKPOINT_JOB_TAG)
    End If
End Function

Function CheckpointJobAttempts(ByVal cells)
    CheckpointJobAttempts = 0
    If UBound(cells) = CHECKPOINT_JOB_LAST_FIELD Then
        If Not IsNumeric(cells(CHECKPOINT_ATTEMPTS_FIELD)) Then
            Err.Raise vbObjectError + ERR_PATH_OR_BUILD, , "Invalid checkpoint attempt count."
        End If
        If CDbl(cells(CHECKPOINT_ATTEMPTS_FIELD)) < 0 Or CDbl(cells(CHECKPOINT_ATTEMPTS_FIELD)) <> Fix(CDbl(cells(CHECKPOINT_ATTEMPTS_FIELD))) Then
            Err.Raise vbObjectError + ERR_PATH_OR_BUILD, , "Invalid checkpoint attempt count."
        End If
        CheckpointJobAttempts = CLng(cells(CHECKPOINT_ATTEMPTS_FIELD))
    ElseIf cells(CHECKPOINT_STATUS_FIELD) = STATE_DONE Or cells(CHECKPOINT_STATUS_FIELD) = STATE_FAILED Or cells(CHECKPOINT_STATUS_FIELD) = STATE_RUNNING Then
        CheckpointJobAttempts = 1
    End If
End Function

Function CheckpointAttempts(ByVal path)
    Dim stateFile, line, cells, valid
    CheckpointAttempts = 0
    stateFile = fso.BuildPath(path, CHECKPOINT_FILENAME)
    If Not fso.FileExists(stateFile) Then
        Exit Function
    End If
    For Each line In Split(ReadUnicode(stateFile), vbCrLf)
        cells = Split(line, vbTab)
        If IsCheckpointJobRow(cells) Then
            valid = True
            If cells(CHECKPOINT_STATUS_FIELD) = STATE_DONE Then
                valid = CheckpointArtifactsValid(path, cells)
            End If
            If valid Then
                CheckpointAttempts = CheckpointAttempts + CheckpointJobAttempts(cells)
            End If
        End If
    Next
End Function

Function CheckpointHandled(ByVal path)
    Dim stateFile, line, cells, handled
    CheckpointHandled = 0
    stateFile = fso.BuildPath(path, CHECKPOINT_FILENAME)
    If Not fso.FileExists(stateFile) Then
        Exit Function
    End If
    For Each line In Split(ReadUnicode(stateFile), vbCrLf)
        cells = Split(line, vbTab)
        If IsCheckpointJobRow(cells) Then
            handled = (cells(CHECKPOINT_STATUS_FIELD) = STATE_FAILED)
            If cells(CHECKPOINT_STATUS_FIELD) = STATE_DONE Then
                handled = CheckpointArtifactsValid(path, cells)
            ElseIf cells(CHECKPOINT_STATUS_FIELD) = STATE_RUNNING And CheckpointJobAttempts(cells) > 1 Then
                handled = True
            End If
            If handled Then
                CheckpointHandled = CheckpointHandled + 1
            End If
        End If
    Next
End Function
