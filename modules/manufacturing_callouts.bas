Attribute VB_Name = "manufacturing_callouts"
Option Explicit

Private Function CollectChamferSpecifications(ByVal model As SldWorks.ModelDoc2) As Object
    Dim result As Object, feature As SldWorks.Feature, display As SldWorks.DisplayDimension
    Dim distance As Double, angle As Double, distances As Long, angles As Long
    Set result = CreateObject("Scripting.Dictionary")
    Set feature = model.FirstFeature
    Do While Not feature Is Nothing
        If feature.GetTypeName2 = FEATURE_TYPE_CHAMFER Then
            distance = 0
            angle = 0
            distances = 0
            angles = 0
            Set display = feature.GetFirstDisplayDimension
            Do While Not display Is Nothing
                If display.Type2 = SW_DIMENSION_LINEAR Then
                    distance = display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue
                    distances = distances + 1
                ElseIf display.Type2 = SW_DIMENSION_ANGULAR Then
                    angle = display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue
                    angles = angles + 1
                End If
                Set display = feature.GetNextDisplayDimension(display)
            Loop
            If distances = 1 And angles = 1 And distance > NATIVE_GEOMETRY_KEY_TOLERANCE_M And angle > DIRECTION_ZERO_TOLERANCE Then
                result.Add feature.Name, Array(distance, angle)
                Log "FEATURE CHAMFER SPEC " & feature.Name & "; distance=" & distance & "; angle=" & angle
            End If
        End If
        Set feature = feature.GetNextFeature
    Loop
    Set CollectChamferSpecifications = result
End Function

Private Function ChamferDistanceCovered(ByVal display As SldWorks.DisplayDimension, ByVal source As String, ByVal specifications As Object) As Boolean
    Dim specification As Variant
    If Not specifications.Exists(source) Or display.Type2 <> SW_DIMENSION_LINEAR Then
        Exit Function
    End If
    specification = specifications(source)
    ChamferDistanceCovered = Abs(display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue - specification(CHAMFER_DISTANCE_INDEX)) < NATIVE_GEOMETRY_KEY_TOLERANCE_M
End Function

Private Sub FormatChamferSpecification(ByVal display As SldWorks.DisplayDimension, ByVal source As String, ByVal specifications As Object)
    Dim specification As Variant
    If Not specifications.Exists(source) Or display.Type2 <> SW_DIMENSION_ANGULAR Then
        Exit Sub
    End If
    specification = specifications(source)
    If Abs(display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue - specification(CHAMFER_ANGLE_INDEX)) > DIRECTION_ZERO_TOLERANCE Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Chamfer angle differs from the source specification."
    End If
    display.SetText SW_DIMENSION_TEXT_PREFIX, "Фаска " & CStr(Round(specification(CHAMFER_DISTANCE_INDEX) * MILLIMETRES_PER_METRE, LINEAR_DECIMAL_PLACES)) & "x"
End Sub

Private Function AxialProfileSpan(ByVal display As SldWorks.DisplayDimension, ByVal view As SldWorks.View, ByVal owners As Object) As Double
    Dim fields As Variant, source As String, references As Variant, first As SldWorks.MathPoint, last As SldWorks.MathPoint
    Dim a As Variant, b As Variant
    If display.Type2 <> SW_DIMENSION_LINEAR Then
        Exit Function
    End If
    fields = Split(display.GetDimension2(SW_DIMENSION_INDEX_FIRST).FullName, "@")
    source = CStr(fields(SOURCE_DIMENSION_FEATURE_INDEX))
    If Not owners.Exists(source) Then
        Exit Function
    End If
    If owners(source) <> FEATURE_TYPE_REVOLUTION And owners(source) <> FEATURE_TYPE_INSTANT_EXTRUDE Then
        Exit Function
    End If
    references = display.GetDimension2(SW_DIMENSION_INDEX_FIRST).ReferencePoints
    If IsEmpty(references) Then
        Exit Function
    End If
    Set first = references(DIMENSION_REFERENCE_FIRST - 1)
    Set last = references(DIMENSION_REFERENCE_SECOND - 1)
    a = PatternSheetPoint(view, first.ArrayData)
    b = PatternSheetPoint(view, last.ArrayData)
    If Abs(a(SCREEN_Y_AXIS_INDEX) - b(SCREEN_Y_AXIS_INDEX)) < NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
        If Abs(Abs(a(SCREEN_X_AXIS_INDEX) - b(SCREEN_X_AXIS_INDEX)) / view.ScaleDecimal - display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue) < NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
            AxialProfileSpan = Abs(a(SCREEN_X_AXIS_INDEX) - b(SCREEN_X_AXIS_INDEX)) / view.ScaleDecimal
        End If
    ElseIf Abs(a(SCREEN_X_AXIS_INDEX) - b(SCREEN_X_AXIS_INDEX)) < NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
        If Abs(Abs(a(SCREEN_Y_AXIS_INDEX) - b(SCREEN_Y_AXIS_INDEX)) / view.ScaleDecimal - display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue) < NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
            AxialProfileSpan = Abs(a(SCREEN_Y_AXIS_INDEX) - b(SCREEN_Y_AXIS_INDEX)) / view.ScaleDecimal
        End If
    End If
End Function

' Keep the overall axial span and its steps on opposite sides of their view.
Private Function AxialReferenceDirection(ByVal display As SldWorks.DisplayDimension) As String
    Dim references As Variant, first As SldWorks.MathPoint, last As SldWorks.MathPoint
    Dim a As Variant, b As Variant, direction As Variant, length As Double, axis As Long
    references = display.GetDimension2(SW_DIMENSION_INDEX_FIRST).ReferencePoints
    Set first = references(DIMENSION_REFERENCE_FIRST - 1)
    Set last = references(DIMENSION_REFERENCE_SECOND - 1)
    a = first.ArrayData
    b = last.ArrayData
    direction = Array(b(0) - a(0), b(1) - a(1), b(2) - a(2))
    For axis = SCREEN_X_AXIS_INDEX To SPACE_LAST_AXIS_INDEX
        length = length + direction(axis) ^ 2
    Next axis
    length = Sqr(length)
    For axis = SCREEN_X_AXIS_INDEX To SPACE_LAST_AXIS_INDEX
        direction(axis) = direction(axis) / length
    Next axis
    AxialReferenceDirection = NativeDirectionKey(direction)
End Function

Private Sub PreferPatternProfileDimensions(ByVal d As SldWorks.ModelDoc2, ByVal patternView As SldWorks.View)
    Dim owners As Object, sheetNames As Variant, sheetName As Variant, view As SldWorks.View
    Dim display As SldWorks.DisplayDimension, span As Double, maximum As Double
    Dim references As Variant, first As SldWorks.MathPoint, last As SldWorks.MathPoint, a As Variant, b As Variant
    Dim maxima As Object, directionKey As String
    Set owners = FeatureOwners(patternView.ReferencedDocument)
    Set maxima = CreateObject("Scripting.Dictionary")
    sheetNames = AsDrawing(d).GetSheetNames
    For Each sheetName In sheetNames
        AsDrawing(d).ActivateSheet CStr(sheetName)
        Set view = AsDrawing(d).GetFirstView.GetNextView
        Do While Not view Is Nothing
            Set display = view.GetFirstDisplayDimension5
            Do While Not display Is Nothing
                span = AxialProfileSpan(display, view, owners)
                If span > NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
                    directionKey = AxialReferenceDirection(display)
                    If Not maxima.Exists(directionKey) Then
                        maxima.Add directionKey, span
                    Else
                        maxima(directionKey) = -MinD(-CDbl(maxima(directionKey)), -span)
                    End If
                End If
                Set display = display.GetNext5
            Loop
            Set view = view.GetNextView
        Loop
    Next sheetName
    For Each sheetName In sheetNames
        AsDrawing(d).ActivateSheet CStr(sheetName)
        Set view = AsDrawing(d).GetFirstView.GetNextView
        Do While Not view Is Nothing
            Set display = view.GetFirstDisplayDimension5
            Do While Not display Is Nothing
                span = AxialProfileSpan(display, view, owners)
                If span > NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
                    maximum = CDbl(maxima(AxialReferenceDirection(display)))
                    If CStr(sheetName) = CStr(sheetNames(LBound(sheetNames))) Then
                        featureDimensionViewPriorities(display.GetNameForSelection) = FEATURE_PROFILE_FRONT_PRIORITY
                        references = display.GetDimension2(SW_DIMENSION_INDEX_FIRST).ReferencePoints
                        Set first = references(DIMENSION_REFERENCE_FIRST - 1)
                        Set last = references(DIMENSION_REFERENCE_SECOND - 1)
                        a = PatternSheetPoint(view, first.ArrayData)
                        b = PatternSheetPoint(view, last.ArrayData)
                        If Abs(a(SCREEN_Y_AXIS_INDEX) - b(SCREEN_Y_AXIS_INDEX)) < NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
                            If Abs(span - maximum) < NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
                                featureDimensionPlacement(display.GetNameForSelection) = NATIVE_LABEL_SIDE_BELOW
                            Else
                                featureDimensionPlacement(display.GetNameForSelection) = NATIVE_LABEL_SIDE_ABOVE
                            End If
                        Else
                            If Abs(span - maximum) < NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
                                featureDimensionPlacement(display.GetNameForSelection) = NATIVE_LABEL_SIDE_LEFT
                            Else
                                featureDimensionPlacement(display.GetNameForSelection) = NATIVE_LABEL_SIDE_RIGHT
                            End If
                        End If
                    End If
                End If
                Set display = display.GetNext5
            Loop
            Set view = view.GetNextView
        Loop
    Next sheetName
End Sub
