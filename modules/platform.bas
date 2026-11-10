Attribute VB_Name = "platform"
Option Explicit
Private Type WindowBounds
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type DrawingTextExtent
    Width As Long
    Height As Long
End Type

Private Type DrawingFontMetrics
    Height As Long
    Ascent As Long
    Descent As Long
    InternalLeading As Long
    ExternalLeading As Long
    AverageWidth As Long
    MaximumWidth As Long
    Weight As Long
    Overhang As Long
    DigitizedAspectX As Long
    DigitizedAspectY As Long
    FirstCharacter As Integer
    LastCharacter As Integer
    DefaultCharacter As Integer
    BreakCharacter As Integer
    Italic As Byte
    Underlined As Byte
    StruckOut As Byte
    PitchAndFamily As Byte
    CharacterSet As Byte
End Type

#If VBA7 Then
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal windowHandle As LongPtr, ByVal command As Long) As Long
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal windowHandle As LongPtr) As Long
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal metric As Long) As Long
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal windowHandle As LongPtr, ByVal insertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal windowHandle As LongPtr, ByRef bounds As WindowBounds) As Long
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal device As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateFontW Lib "gdi32" (ByVal height As Long, ByVal width As Long, ByVal escapement As Long, ByVal orientation As Long, ByVal weight As Long, ByVal italic As Long, ByVal underline As Long, ByVal strikeout As Long, ByVal characterSet As Long, ByVal outputPrecision As Long, ByVal clipPrecision As Long, ByVal quality As Long, ByVal pitch As Long, ByVal face As LongPtr) As LongPtr
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal device As LongPtr, ByVal item As LongPtr) As LongPtr
Private Declare PtrSafe Function GetTextExtentPoint32W Lib "gdi32" (ByVal device As LongPtr, ByVal text As LongPtr, ByVal count As Long, ByRef extent As DrawingTextExtent) As Long
Private Declare PtrSafe Function GetTextMetricsW Lib "gdi32" (ByVal device As LongPtr, ByRef metrics As DrawingFontMetrics) As Long
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal item As LongPtr) As Long
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal device As LongPtr) As Long
#Else
Private Declare Function ShowWindow Lib "user32" (ByVal windowHandle As Long, ByVal command As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal windowHandle As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal metric As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal windowHandle As Long, ByVal insertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long, ByVal flags As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal windowHandle As Long, ByRef bounds As WindowBounds) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal device As Long) As Long
Private Declare Function CreateFontW Lib "gdi32" (ByVal height As Long, ByVal width As Long, ByVal escapement As Long, ByVal orientation As Long, ByVal weight As Long, ByVal italic As Long, ByVal underline As Long, ByVal strikeout As Long, ByVal characterSet As Long, ByVal outputPrecision As Long, ByVal clipPrecision As Long, ByVal quality As Long, ByVal pitch As Long, ByVal face As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal device As Long, ByVal item As Long) As Long
Private Declare Function GetTextExtentPoint32W Lib "gdi32" (ByVal device As Long, ByVal text As Long, ByVal count As Long, ByRef extent As DrawingTextExtent) As Long
Private Declare Function GetTextMetricsW Lib "gdi32" (ByVal device As Long, ByRef metrics As DrawingFontMetrics) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal item As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal device As Long) As Long
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
