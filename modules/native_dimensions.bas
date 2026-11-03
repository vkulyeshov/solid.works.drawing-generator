Attribute VB_Name = "native_dimensions"
Option Explicit
Private Function NativeDimensionSheets(ByVal d As SldWorks.ModelDoc2) As Long
    Dim sheetName As Variant, view As SldWorks.View, transform As Variant, axis As Long, orthographic As Boolean
    Set featureDimensionOwners = Nothing
    Set featureDimensionRepeatedKeys = Nothing
    Set featureDimensionViewPriorities = Nothing
    Set featureDimensionOccurrenceCounts = Nothing
    Set featureCircleClassifications = Nothing
    Set featureDimensionPlacement = Nothing
    Set featurePatternGraphics = Nothing
    If Setting("AutoDimensions", "1") <> "1" Then
        Exit Function
    End If
    If SixViewMode() Then
        NormalizeSixViewScales d
    End If
    If FeatureDimensionMode() Then
        NativeDimensionSheets = FeatureDimensionSheets(d)
        If NativeDimensionSheets > 0 Then
            Exit Function
        End If
        If LCase$(Setting("OrthographicViews", "auto")) = "auto" And Not SixViewMode() Then
            Err.Raise vbObjectError + ERR_FEATURE_VIEW_REQUIRED, , "An unparameterized solid requires opposite views for geometry coverage."
        End If
    End If
    For Each sheetName In AsDrawing(d).GetSheetNames
        AsDrawing(d).ActivateSheet CStr(sheetName)
        DoEvents
        Set view = AsDrawing(d).GetFirstView.GetNextView
        Do While Not view Is Nothing
            transform = view.ModelToViewTransform.ArrayData
            orthographic = False
            For axis = MATRIX_XZ_INDEX To MATRIX_ZZ_INDEX Step 3
                If Abs(transform(axis)) >= 1 - DIRECTION_ZERO_TOLERANCE Then
                    orthographic = True
                End If
            Next axis
            If orthographic Then
                NativeDimensionSheets = NativeDimensionSheets + NativeDimensionView(d, view)
            End If
            Set view = view.GetNextView
        Loop
    Next sheetName
End Function

Private Function TryNativeDimensionSheets(ByVal d As SldWorks.ModelDoc2, ByRef needsViews As Boolean) As Long
    Dim errorNumber As Long, errorText As String
    On Error GoTo Failed
    TryNativeDimensionSheets = NativeDimensionSheets(d)
    Exit Function
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    If errorNumber = vbObjectError + ERR_FEATURE_VIEW_REQUIRED And LCase$(Setting("OrthographicViews", "auto")) = "auto" And Not SixViewMode() Then
        needsViews = True
        Log "FEATURE VIEWS adding opposite projections: " & errorText
        Err.Clear
        Exit Function
    End If
    Err.Raise errorNumber, , errorText
End Function

Private Sub ClearNativeDrawingDimensions(ByVal d As SldWorks.ModelDoc2)
    Dim sheetName As Variant, view As SldWorks.View, display As SldWorks.DisplayDimension, following As SldWorks.DisplayDimension
    For Each sheetName In AsDrawing(d).GetSheetNames
        AsDrawing(d).ActivateSheet CStr(sheetName)
        Set view = AsDrawing(d).GetFirstView.GetNextView
        Do While Not view Is Nothing
            Set display = view.GetFirstDisplayDimension5
            Do While Not display Is Nothing
                Set following = display.GetNext5
                RemoveFeatureDimension d, display, "replanning dimensions after adding required views"
                Set display = following
            Loop
            Set view = view.GetNextView
        Loop
    Next sheetName
End Sub

Private Function NativeDimensionView(ByVal d As SldWorks.ModelDoc2, ByVal v As SldWorks.View) As Long
    Dim status As Long, display As SldWorks.DisplayDimension, active As SldWorks.ModelDoc2
    Dim error As Long, position As Variant, selected As Boolean, feature As SldWorks.Feature
    Dim components As Variant, component As Variant, edges As Variant, item As Variant, edge As SldWorks.Edge
    Dim extension As Object, objects(0) As Object, selectionData As SldWorks.SelectData
    If Setting("AutoDimensions", "1") <> "1" Then
        Exit Function
    End If
    Set active = sw.ActivateDoc3(d.GetPathName, False, SW_REBUILD_USER_DECISION, error)
    If active Is Nothing Then
        Err.Raise vbObjectError + ERR_VIEW_ACTIVATION, , "Cannot activate the drawing for native dimensioning."
    End If
    DoEvents
    d.ClearSelection2 True
    AsDrawing(d).EditSheet
    AsDrawing(d).ActivateView ""
    d.GraphicsRedraw2
    DoEvents
    position = v.Position
    selected = d.Extension.SelectByID2(v.Name, "DRAWINGVIEW", position(0), position(1), 0, False, 0, Nothing, 0)
    If Not selected Then
        Set extension = d.Extension
        selected = extension.SelectByID2(CStr(v.Name), "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
    End If
    If Not selected Then
        Set objects(0) = v
        selected = d.Extension.MultiSelect2(objects, False, Nothing) > 0
    End If
    If Not selected Then
        Set feature = d.FeatureByName(v.Name)
        If Not feature Is Nothing Then
            selected = feature.Select2(False, 0)
        End If
    End If
    If Not selected Then
        AsDrawing(d).ActivateView v.Name
        components = v.GetVisibleComponents
        If Not IsEmpty(components) Then
            For Each component In components
                edges = v.GetVisibleEntities2(component, SW_VISIBLE_ENTITY_EDGE)
                If Not IsEmpty(edges) Then
                    For Each item In edges
                        Set edge = item
                        selected = v.SelectEntity(edge, False)
                        If Not selected Then
                            Set selectionData = d.SelectionManager.CreateSelectData
                            Set selectionData.View = v
                            selected = edge.Select4(False, selectionData)
                        End If
                        If selected Then
                            Exit For
                        End If
                    Next item
                End If
                If selected Then
                    Exit For
                End If
            Next component
        End If
    End If
    If Not selected Then
        Log "NATIVE SELECTION PID=" & sw.GetProcessID & "; document=" & d.GetTitle & "; type=" & d.GetType & "; active=" & active.GetTitle & "; view=" & v.GetName2 & "; feature=" & CStr(Not feature Is Nothing) & "; components=" & CStr(v.GetVisibleComponentCount)
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot select a view for native dimensioning: " & v.Name
    End If
    status = AsDrawing(d).AutoDimension(SW_AUTODIM_ALL_ENTITIES, SW_AUTODIM_CHAIN, SW_AUTODIM_ABOVE, SW_AUTODIM_CHAIN, SW_AUTODIM_LEFT)
    d.ClearSelection2 True
    Set display = v.GetFirstDisplayDimension5
    Do While Not display Is Nothing
        NativeDimensionView = NativeDimensionView + 1
        Set display = display.GetNext5
    Loop
    Log "NATIVE DIMENSIONS " & v.Name & "; status=" & status & "; dimensions=" & NativeDimensionView
    If NativeDimensionView = 0 Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Native dimensioning produced no dimensions: " & v.Name
    End If
End Function

