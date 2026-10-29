Attribute VB_Name = "background"
Option Explicit
Private Sub PrepareDocumentWindow(ByVal document As SldWorks.ModelDoc2)
    Dim frame As SldWorks.Frame
    ' Drawing export needs a document window and graphics context, even with a hidden frame.
    document.Visible = True
    If backgroundOwned Then
        sw.UserControlBackground = True
        Set frame = sw.Frame: frame.KeepInvisible = True
        HideBackgroundFrame
    End If
End Sub

Private Sub HideBackgroundFrame()
    Dim frame As SldWorks.Frame
    If backgroundOwned = False Then
        Exit Sub
    End If
    On Error Resume Next
    Set frame = sw.Frame
    ' This SolidWorks build needs a visible graphics window for PDF export.
    ' Place only our dedicated frame beyond the entire virtual desktop, without activation.
    If FrameIsOffscreen() And IsWindowVisible(frame.GetHWndx64) <> 0 Then
        Exit Sub
    End If
    sw.FrameState = SW_WINDOW_NORMAL
    ' NOZORDER | NOACTIVATE | SHOWWINDOW | NOSENDCHANGING. SolidWorks FrameLeft/Top clamp to the desktop.
    SetWindowPos frame.GetHWndx64, WIN_INSERT_AFTER_DEFAULT, GetSystemMetrics(WIN_METRIC_VIRTUAL_LEFT) + GetSystemMetrics(WIN_METRIC_VIRTUAL_WIDTH) + BACKGROUND_SCREEN_OFFSET_PX, GetSystemMetrics(WIN_METRIC_VIRTUAL_TOP) + GetSystemMetrics(WIN_METRIC_VIRTUAL_HEIGHT) + BACKGROUND_SCREEN_OFFSET_PX, BACKGROUND_WINDOW_WIDTH_PX, BACKGROUND_WINDOW_HEIGHT_PX, BACKGROUND_WINDOW_POSITION_FLAGS
End Sub

Private Function FrameIsOffscreen() As Boolean
    Dim bounds As WindowBounds, frame As SldWorks.Frame
    Set frame = sw.Frame
    If GetWindowRect(frame.GetHWndx64, bounds) = 0 Then
        Exit Function
    End If
    FrameIsOffscreen = bounds.Left >= GetSystemMetrics(WIN_METRIC_VIRTUAL_LEFT) + GetSystemMetrics(WIN_METRIC_VIRTUAL_WIDTH) And bounds.Top >= GetSystemMetrics(WIN_METRIC_VIRTUAL_TOP) + GetSystemMetrics(WIN_METRIC_VIRTUAL_HEIGHT)
End Function

Private Sub FinishBackground()
    Dim state As Object, frame As SldWorks.Frame, bounds As WindowBounds
    If backgroundOwned = False Then
        Exit Sub
    End If
    On Error Resume Next
    Set frame = sw.Frame
    HideBackgroundFrame
    Set state = fs.CreateTextFile(sw.GetCurrentMacroPathFolder & BACKGROUND_STATE_RELATIVE_PATH, True, True)
    state.WriteLine "PID=" & sw.GetProcessID
    state.WriteLine "Hidden=" & CStr(IsWindowVisible(frame.GetHWndx64) = 0)
    state.WriteLine "Offscreen=" & CStr(FrameIsOffscreen())
    GetWindowRect frame.GetHWndx64, bounds
    state.WriteLine "FrameLeft=" & bounds.Left
    state.WriteLine "FrameTop=" & bounds.Top
    state.WriteLine "ApplicationVisibleFlag=" & CStr(sw.Visible)
    state.Close
    sw.CloseAllDocuments True
    ShowWindow frame.GetHWndx64, WIN_SHOW_HIDE
    ' Avoid persisting an inaccessible position for the next interactive launch.
    If backgroundLeft < -BACKGROUND_RESET_DISTANCE_PX Or backgroundTop < -BACKGROUND_RESET_DISTANCE_PX Then
        backgroundLeft = BACKGROUND_DEFAULT_POSITION_PX
        backgroundTop = BACKGROUND_DEFAULT_POSITION_PX
    End If
    If backgroundWidth < BACKGROUND_MIN_WINDOW_SIZE_PX Then
        backgroundWidth = BACKGROUND_WINDOW_WIDTH_PX
    End If
    If backgroundHeight < BACKGROUND_MIN_WINDOW_SIZE_PX Then
        backgroundHeight = BACKGROUND_WINDOW_HEIGHT_PX
    End If
    sw.FrameLeft = backgroundLeft: sw.FrameTop = backgroundTop
    sw.FrameWidth = backgroundWidth: sw.FrameHeight = backgroundHeight
    Set currentDrawing = Nothing: Set rootModel = Nothing: Set original = Nothing
    Set jobs = Nothing: Set quantities = Nothing: Set flatBody = Nothing
    ' Document writes and window restoration are complete before this marker.
    Set state = fs.OpenTextFile(sw.GetCurrentMacroPathFolder & BACKGROUND_STATE_RELATIVE_PATH, FILE_FOR_APPENDING, False, FILE_UNICODE)
    state.WriteLine "ReadyToExit=1"
    state.Close
    sw.ExitApp
    Set sw = Nothing
End Sub

