Attribute VB_Name = "orientation"
Option Explicit
Private Sub UprightView(ByVal v As SldWorks.View)
    Dim axis(2) As Double, vector As SldWorks.MathVector, util As SldWorks.MathUtility
    Dim p As Variant, up As String, angle As Double, direction As Double
    up = UCase$(Setting("UpAxis", DEFAULT_UP_AXIS)): direction = 1
    If Left$(up, 1) = "-" Then
        direction = -1
    End If
    Select Case Right$(up, 1)
        Case "X": axis(0) = direction
        Case "Y": axis(1) = direction
        Case "Z": axis(2) = direction
        Case Else: Err.Raise vbObjectError + ERR_UP_AXIS, , "Invalid UpAxis."
    End Select
    Set util = sw.GetMathUtility: Set vector = util.CreateVector(axis)
    Set vector = vector.MultiplyTransform(v.ModelToViewTransform): p = vector.ArrayData
    If Sqr(p(0) ^ 2 + p(1) ^ 2) < MIN_PROJECTED_UP_VECTOR_LENGTH Then ' looking along up axis
        Exit Sub
    End If
    If Abs(p(1)) < UP_VECTOR_TOLERANCE Then
        angle = Sgn(p(0)) * PI / 2
    Else
        angle = Atn(p(0) / p(1))
        If p(1) < 0 Then
            angle = angle + PI
        End If
    End If
    v.Angle = v.Angle + angle
End Sub

Private Sub ImportModelItems(ByVal d As SldWorks.ModelDoc2)
    Dim annotations As Variant, types As Long
    On Error GoTo Failed
    If Setting("ImportModelDimensions", "0") = "0" Then
        Exit Sub
    End If
    types = SW_MODEL_ITEMS_MARKED_DIMENSIONS + SW_MODEL_ITEMS_HOLE_PROFILES + SW_MODEL_ITEMS_HOLE_LOCATIONS ' marked dimensions and Hole Wizard dimensions
    If LCase$(Setting("ImportModelDimensions", "0")) = "all" Then
        types = SW_MODEL_ITEMS_ALL_DIMENSIONS
    End If
    AsDrawing(d).ActivateView "": d.ClearSelection2 True
    annotations = AsDrawing(d).InsertModelAnnotations3(0, types, True, True, False, False)
    Exit Sub
Failed:
    Warn "Import model dimensions: " & Err.Description
End Sub

