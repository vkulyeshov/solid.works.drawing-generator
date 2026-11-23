Attribute VB_Name = "execution_budget"
Option Explicit

Private jobTimeoutSeconds As Long, layoutTimeoutSeconds As Long
Private executionJobActive As Boolean, executionLayoutActive As Boolean
Private executionJobStarted As Date, executionLayoutStarted As Date
Private executionJobName As String, executionLayoutName As String
Private executionTimeoutMessage As String, executionPollCount As Long
Private executionStrictActive As Boolean, executionStrictExpired As Boolean, executionStrictStarted As Date

Private Function ExecutionTimeoutSetting(ByVal key As String, ByVal fallback As Long) As Long
    Dim value As String, number As Double
    value = Setting(key, CStr(fallback))
    If Not IsNumeric(value) Then
        Err.Raise vbObjectError + ERR_DRAWING_SETTINGS, , key & " must be an integer from 0 to 86400."
    End If
    number = CDbl(value)
    If number < 0 Or number > MAX_EXECUTION_TIMEOUT_SECONDS Or number <> Fix(number) Then
        Err.Raise vbObjectError + ERR_DRAWING_SETTINGS, , key & " must be an integer from 0 to 86400."
    End If
    ExecutionTimeoutSetting = CLng(number)
End Function

Private Sub ConfigureExecutionBudgets()
    jobTimeoutSeconds = ExecutionTimeoutSetting("JobTimeoutSeconds", DEFAULT_JOB_TIMEOUT_SECONDS)
    layoutTimeoutSeconds = ExecutionTimeoutSetting("LayoutTimeoutSeconds", DEFAULT_LAYOUT_TIMEOUT_SECONDS)
    EndJobBudget
End Sub

Private Sub BeginJobBudget(ByVal name As String)
    executionJobName = name
    executionJobStarted = Now
    executionJobActive = True
    executionLayoutActive = False
    executionStrictActive = False
    executionStrictExpired = False
    executionTimeoutMessage = ""
    executionPollCount = 0
End Sub

Private Sub EndJobBudget()
    executionJobActive = False
    executionLayoutActive = False
    executionTimeoutMessage = ""
    executionStrictActive = False
    executionStrictExpired = False
End Sub

Private Sub BeginLayoutBudget(ByVal name As String)
    executionLayoutName = name
    executionLayoutStarted = Now
    executionLayoutActive = True
    CheckExecutionBudget
End Sub

Private Sub EndLayoutBudget()
    executionLayoutActive = False
    CheckExecutionBudget
End Sub

Private Sub CheckExecutionBudget()
    If executionTimeoutMessage = "" Then
        If executionJobActive And jobTimeoutSeconds > 0 Then
            If DateDiff("s", executionJobStarted, Now) >= jobTimeoutSeconds Then
                executionTimeoutMessage = EXECUTION_TIMEOUT_PREFIX & "job exceeded " & jobTimeoutSeconds & " seconds: " & executionJobName
            End If
        End If
        If executionLayoutActive And layoutTimeoutSeconds > 0 Then
            If DateDiff("s", executionLayoutStarted, Now) >= layoutTimeoutSeconds Then
                executionTimeoutMessage = EXECUTION_TIMEOUT_PREFIX & "sheet layout exceeded " & layoutTimeoutSeconds & " seconds: " & executionLayoutName
            End If
        End If
    End If
    If executionTimeoutMessage <> "" Then
        Err.Raise vbObjectError + ERR_EXECUTION_TIMEOUT, , executionTimeoutMessage
    End If
End Sub

' Check inside candidate filtering as well as recursive search nodes.
Private Sub PollExecutionBudget()
    executionPollCount = executionPollCount + 1
    If executionPollCount >= EXECUTION_POLL_INTERVAL Then
        executionPollCount = 0
        CheckExecutionBudget
        If executionStrictActive Then
            executionStrictExpired = (DateDiff("s", executionStrictStarted, Now) >= MEMORY_STRICT_SEARCH_SECONDS)
        End If
    End If
End Sub

Private Sub BeginStrictSearch()
    executionStrictStarted = Now
    executionStrictActive = True
    executionStrictExpired = False
End Sub

Private Sub EndStrictSearch()
    executionStrictActive = False
    executionStrictExpired = False
End Sub

Private Function MemoryStrictSearchExpired() As Boolean
    MemoryStrictSearchExpired = executionStrictActive And executionStrictExpired
End Function

Private Function JobTimedOut(ByVal job As Object) As Boolean
    If Setting("RetryTimedOutJobs", "0") = "1" Then
        Exit Function
    End If
    JobTimedOut = (InStr(1, CStr(job("stage")), EXECUTION_TIMEOUT_PREFIX, vbBinaryCompare) > 0)
End Function

Private Sub ApplyJobTimeoutMarkers()
    Dim path As String, stream As Object, fields As Variant, key As Variant, job As Object, markers As Object
    path = fs.BuildPath(runDir, TIMEOUT_JOBS_FILENAME)
    If Not fs.FileExists(path) Then
        Exit Sub
    End If
    Set markers = CreateObject("Scripting.Dictionary")
    Set stream = fs.OpenTextFile(path, FILE_FOR_READING, False, FILE_UNICODE)
    Do While Not stream.AtEndOfStream
        fields = Split(stream.ReadLine, vbTab)
        If UBound(fields) = 1 Then
            If Left$(fields(1), Len(EXECUTION_TIMEOUT_PREFIX)) = EXECUTION_TIMEOUT_PREFIX Then
                markers(fields(0)) = fields(1)
            End If
        End If
    Loop
    stream.Close
    For Each key In jobs.Keys
        Set job = jobs(key)
        If job("status") <> STATE_DONE And markers.Exists(CStr(job("out"))) Then
            job("status") = STATE_FAILED
            job("stage") = markers(CStr(job("out")))
            If interruptedJobNumber = CLng(job("number")) Then
                interruptedJobNumber = 0
            End If
        End If
    Next key
End Sub
