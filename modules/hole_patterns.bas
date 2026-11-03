Attribute VB_Name = "hole_patterns"
Option Explicit
Private Sub HoleMemberNames(ByVal feature As SldWorks.Feature, ByVal members As Object)
    Dim child As SldWorks.Feature
    members(feature.Name) = True
    Set child = feature.GetFirstSubFeature
    Do While Not child Is Nothing
        HoleMemberNames child, members
        Set child = child.GetNextSubFeature
    Loop
End Sub

Private Function CollectHolePatterns(ByVal model As SldWorks.ModelDoc2) As Object
    Dim result As Object, patterns As New Collection, members As Object, feature As SldWorks.Feature
    Dim data As SldWorks.WizardHoleFeatureData2, faces As Variant, item As Variant, face As SldWorks.Face2
    Dim surface As SldWorks.Surface, parameters As Variant, pattern As Object, centers As Object, origin As Variant
    Dim station As Double, key As String, normal As Variant, profile As String
    Set result = CreateObject("Scripting.Dictionary")
    Set members = CreateObject("Scripting.Dictionary")
    Set feature = model.FirstFeature
    Do While Not feature Is Nothing
        If feature.GetTypeName2 = FEATURE_TYPE_HOLE_WIZARD Then
            HoleMemberNames feature, members
            Set data = feature.GetDefinition
            Set pattern = CreateObject("Scripting.Dictionary")
            Set centers = CreateObject("Scripting.Dictionary")
            pattern.Add "feature", feature.Name
            pattern.Add "centers", centers
            pattern.Add "size", data.FastenerSize
            pattern.Add "class", data.ThreadClass
            pattern.Add "end", data.EndCondition
            profile = CStr(data.FastenerType2) & "|" & data.FastenerSize & "|" & data.ThreadClass & "|" & CStr(data.EndCondition)
            If data.EndCondition <> SW_END_THROUGH_ALL Then
                profile = profile & "|" & CStr(Round(data.Depth / NATIVE_GEOMETRY_KEY_TOLERANCE_M, 0))
            End If
            pattern.Add "profile", profile
            faces = feature.GetFaces
            If Not IsEmpty(faces) Then
                For Each item In faces
                    Set face = item
                    Set surface = face.GetSurface
                    If surface.IsCylinder Then
                        parameters = surface.CylinderParams
                        normal = Array(parameters(CIRCLE_AXIS_X_INDEX), parameters(CIRCLE_AXIS_Y_INDEX), parameters(CIRCLE_AXIS_Z_INDEX))
                        station = parameters(0) * normal(0) + parameters(1) * normal(1) + parameters(2) * normal(2)
                        origin = Array(parameters(0) - station * normal(0), parameters(1) - station * normal(1), parameters(2) - station * normal(2))
                        key = NativePointKey(origin)
                        If Not centers.Exists(key) Then
                            centers.Add key, parameters
                        End If
                        pattern("radius") = parameters(CIRCLE_RADIUS_INDEX)
                        pattern("normal") = normal
                    End If
                Next item
            End If
            If centers.Count > 0 Then
                patterns.Add pattern
                Log "HOLE PATTERN " & feature.Name & "; count=" & centers.Count & "; profile=" & profile
            End If
        End If
        Set feature = feature.GetNextFeature
    Loop
    result.Add "patterns", patterns
    result.Add "members", members
    Set CollectHolePatterns = result
End Function

Private Function SelectFeatureEdge(ByVal d As SldWorks.ModelDoc2, ByVal view As SldWorks.View, ByVal edge As SldWorks.Edge, Optional ByVal append As Boolean = False) As Boolean
    SelectFeatureEdge = SelectFeatureReference(d, view, edge, append)
End Function

Private Function SelectFeatureReference(ByVal d As SldWorks.ModelDoc2, ByVal view As SldWorks.View, ByVal reference As Object, Optional ByVal append As Boolean = False) As Boolean
    Dim data As SldWorks.SelectData, context As Object, entity As SldWorks.Entity, manager As SldWorks.SelectionMgr
    If Not append Then
        d.ClearSelection2 True
    End If
    Set context = view
    Set entity = reference
    SelectFeatureReference = context.SelectEntity(entity, append)
    If Not SelectFeatureReference Then
        AsDrawing(d).ActivateView CStr(view.Name)
        Set manager = d.SelectionManager
        Set data = manager.CreateSelectData
        If Not data Is Nothing Then
            Set data.View = view
        End If
        SelectFeatureReference = entity.Select4(append, data)
    End If
End Function

Private Function VisibleHoleEdge(ByVal view As SldWorks.View, ByVal pattern As Object) As SldWorks.Edge
    Dim components As Variant, component As Variant, edges As Variant, item As Variant
    Dim edge As SldWorks.Edge, origin As Variant, station As Double, key As String
    Dim curve As SldWorks.Curve, parameters As Variant, transform As Variant, normal As Variant, aligned As Double
    Dim stage As String, errorNumber As Long, errorText As String
    On Error GoTo Failed
    stage = "view transform"
    transform = view.ModelToViewTransform.ArrayData
    normal = pattern("normal")
    aligned = Abs(normal(0) * transform(MATRIX_XZ_INDEX) + normal(1) * transform(MATRIX_YZ_INDEX) + normal(2) * transform(MATRIX_ZZ_INDEX))
    If aligned < FACE_NORMAL_ALIGNMENT_COSINE Then
        Exit Function
    End If
    components = view.GetVisibleComponents
    If IsEmpty(components) Then
        Exit Function
    End If
    For Each component In components
        stage = "visible edges"
        edges = view.GetVisibleEntities2(component, SW_VISIBLE_ENTITY_EDGE)
        If Not IsEmpty(edges) Then
            For Each item In edges
                Set edge = item
                stage = "edge curve"
                Set curve = edge.GetCurve
                If curve Is Nothing Then
                    GoTo NextHoleEdge
                End If
                If curve.IsCircle Then
                    parameters = curve.CircleParams
                    If Abs(parameters(CIRCLE_RADIUS_INDEX) - CDbl(pattern("radius"))) < NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
                        station = parameters(0) * normal(0) + parameters(1) * normal(1) + parameters(2) * normal(2)
                        origin = Array(parameters(0) - station * normal(0), parameters(1) - station * normal(1), parameters(2) - station * normal(2))
                        key = NativePointKey(origin)
                        If pattern("centers").Exists(key) Then
                            Set VisibleHoleEdge = edge
                            Exit Function
                        End If
                    End If
                End If
NextHoleEdge:
            Next item
        End If
    Next component
    Exit Function
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    Err.Raise errorNumber, , "Visible hole edge, " & stage & ": " & errorText
End Function

Private Sub AddSharedHoleSizes(ByVal d As SldWorks.ModelDoc2, ByVal evidence As Object)
    Dim groups As Object, item As Variant, pattern As Object, group As Object, key As Variant
    Dim sheetName As Variant, view As SldWorks.View, edge As SldWorks.Edge, display As SldWorks.DisplayDimension
    Dim outline As Variant, finished As Boolean, definition As String
    Set groups = CreateObject("Scripting.Dictionary")
    For Each item In evidence("patterns")
        Set pattern = item
        key = CStr(pattern("profile")) & "|" & CStr(Round(CDbl(pattern("radius")) / NATIVE_GEOMETRY_KEY_TOLERANCE_M, 0))
        If Not groups.Exists(key) Then
            Set group = CreateObject("Scripting.Dictionary")
            group.Add "pattern", pattern
            group.Add "count", 0
            groups.Add key, group
        End If
        Set group = groups(key)
        group("count") = CLng(group("count")) + pattern("centers").Count
    Next item
    For Each key In groups.Keys
        Set group = groups(key)
        Set pattern = group("pattern")
        Log "HOLE SHARED CREATE count=" & group("count") & "; feature=" & pattern("feature")
        finished = False
        For Each sheetName In AsDrawing(d).GetSheetNames
            AsDrawing(d).ActivateSheet CStr(sheetName)
            Set view = AsDrawing(d).GetFirstView.GetNextView
            Do While Not view Is Nothing
                Log "HOLE SHARED SEARCH " & view.Name
                Set edge = VisibleHoleEdge(view, pattern)
                If Not edge Is Nothing Then
                    If SelectFeatureEdge(d, view, edge) Then
                        Log "HOLE SHARED EDGE selected in " & view.Name
                        outline = view.GetOutline
                        Set display = AsDrawing(d).AddHoleCallout2(outline(2) + DIAMETER_LABEL_OFFSET_M, outline(3), 0)
                        If Not display Is Nothing Then
                            definition = display.GetText(SW_DIMENSION_TEXT_CALLOUT_ABOVE_DEFINITION)
                            definition = Replace(definition, SW_HOLE_INSTANCE_TOKEN, CStr(group("count")))
                            display.SetText SW_DIMENSION_TEXT_CALLOUT_ABOVE, definition
                            featureDimensionOwners.Add display.GetNameForSelection, "hole-profile|" & CStr(key)
                            Log "HOLE SHARED SIZE count=" & group("count") & "; view=" & view.Name
                            finished = True
                            Exit Do
                        End If
                    End If
                End If
                Set view = view.GetNextView
            Loop
            If finished Then
                Exit For
            End If
        Next sheetName
        If Not finished Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot create a shared hole-size annotation."
        End If
    Next key
End Sub
