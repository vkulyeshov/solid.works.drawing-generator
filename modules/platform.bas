Attribute VB_Name = "platform"
Option Explicit
Private Type WindowBounds
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

#If VBA7 Then
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal windowHandle As LongPtr, ByVal command As Long) As Long
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal windowHandle As LongPtr) As Long
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal metric As Long) As Long
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal windowHandle As LongPtr, ByVal insertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal windowHandle As LongPtr, ByRef bounds As WindowBounds) As Long
#Else
Private Declare Function ShowWindow Lib "user32" (ByVal windowHandle As Long, ByVal command As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal windowHandle As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal metric As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal windowHandle As Long, ByVal insertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal flags As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal windowHandle As Long, ByRef bounds As WindowBounds) As Long
#End If

' SolidWorks 2025 VBA. All geometry remains native and associative.
' Entry point: main. Source documents are not passed to Save3.
' Pack and Go may reserialize an assembly; exact source backups are retained first.
' A Pack and Go snapshot is made first; derived flat configurations are saved there.
#If VBA7 Then
Private Declare PtrSafe Function MoveFileExW Lib "kernel32" (ByVal existingName As LongPtr, ByVal newName As LongPtr, ByVal flags As Long) As Long
#Else
Private Declare Function MoveFileExW Lib "kernel32" (ByVal existingName As Long, ByVal newName As Long, ByVal flags As Long) As Long
#End If
