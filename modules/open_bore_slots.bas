Attribute VB_Name = "open_bore_slots"
Option Explicit

' Replace the rectangle's construction offsets with the exposed keyway geometry.
Private Function OpenBoreSlotPlan(ByVal bounds As Variant, ByVal center As Variant, ByVal radius As Double) As Variant
    Dim axis As Long, across As Long, direction As Long, halfWidth As Double, mouth As Double, inner As Double, floor As Double
    For axis = SCREEN_X_AXIS_INDEX To SCREEN_Y_AXIS_INDEX
        across = SCREEN_Y_AXIS_INDEX - axis
        halfWidth = (bounds(across + PLANAR_AXIS_COUNT) - bounds(across)) / 2
        If halfWidth > NATIVE_GEOMETRY_KEY_TOLERANCE_M And halfWidth < radius Then
            If Abs((bounds(across + PLANAR_AXIS_COUNT) + bounds(across)) / 2 - center(across)) < NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
                mouth = Sqr(radius * radius - halfWidth * halfWidth)
                For direction = -1 To 1 Step 2
                    inner = (bounds(axis) - center(axis)) * direction
                    floor = (bounds(axis + PLANAR_AXIS_COUNT) - center(axis)) * direction
                    If direction < 0 Then
                        inner = (bounds(axis + PLANAR_AXIS_COUNT) - center(axis)) * direction
                        floor = (bounds(axis) - center(axis)) * direction
                    End If
                    If inner >= 0 And inner < mouth And floor > radius Then
                        OpenBoreSlotPlan = Array(axis, center(axis) + direction * mouth, center(axis) + direction * floor, bounds(across), bounds(across + PLANAR_AXIS_COUNT), halfWidth * 2, floor - mouth)
                        Exit Function
                    End If
                Next direction
            End If
        End If
    Next axis
End Function

Private Sub CollectOpenBoreSketch(ByVal feature As SldWorks.Feature, ByVal slots As Object)
    Dim sketch As SldWorks.Sketch, segments As Variant, item As Variant, segment As SldWorks.SketchSegment
    Dim line As SldWorks.SketchLine, first As Variant, last As Variant, bounds As Variant, count As Long, transform As Variant
    Dim key As Variant, record As Object, parameters As Variant, center As Variant, plan As Variant, aligned As Double
    Dim corners As Object, corner As Variant, pointKey As String
    If feature.GetTypeName2 <> FEATURE_TYPE_SKETCH Then
        Exit Sub
    End If
    Set sketch = feature.GetSpecificFeature2
    segments = sketch.GetSketchSegments
    If IsEmpty(segments) Then
        Exit Sub
    End If
    bounds = Array(SEARCH_INITIAL_LIMIT, SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT)
    Set corners = CreateObject("Scripting.Dictionary")
    For Each item In segments
        Set segment = item
        If segment.ConstructionGeometry = False Then
            If segment.GetType <> SW_SKETCH_LINE Then
                Exit Sub
            End If
            Set line = segment
            first = SourceSketchPoint(line.GetStartPoint2)
            last = SourceSketchPoint(line.GetEndPoint2)
            If Abs(first(0) - last(0)) > NATIVE_GEOMETRY_KEY_TOLERANCE_M And Abs(first(1) - last(1)) > NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
                Exit Sub
            End If
            NativeExpandBounds bounds, first(0), first(1)
            NativeExpandBounds bounds, last(0), last(1)
            For Each corner In Array(first, last)
                pointKey = NativePointKey(corner)
                If Not corners.Exists(pointKey) Then
                    corners.Add pointKey, 0
                End If
                corners(pointKey) = CLng(corners(pointKey)) + 1
            Next corner
            count = count + 1
        End If
    Next item
    If count <> RECTANGLE_EDGE_COUNT Then
        Exit Sub
    End If
    If corners.Count <> RECTANGLE_EDGE_COUNT Then
        Exit Sub
    End If
    For Each corner In corners.Keys
        If CLng(corners(corner)) <> RECTANGLE_VERTEX_DEGREE Then
            Exit Sub
        End If
    Next corner
    transform = sketch.ModelToSketchTransform.ArrayData
    For Each key In featureCircleClassifications.Keys
        Set record = featureCircleClassifications(key)
        If CBool(record("bore")) Then
            parameters = record("parameters")
            aligned = Abs(parameters(CIRCLE_AXIS_X_INDEX) * transform(MATRIX_XZ_INDEX) + parameters(CIRCLE_AXIS_Y_INDEX) * transform(MATRIX_YZ_INDEX) + parameters(CIRCLE_AXIS_Z_INDEX) * transform(MATRIX_ZZ_INDEX))
            If aligned > FACE_NORMAL_ALIGNMENT_COSINE Then
                center = SourcePointInSketch(parameters, transform)
                plan = OpenBoreSlotPlan(bounds, center, parameters(CIRCLE_RADIUS_INDEX) * Abs(transform(MATRIX_SCALE_INDEX)))
                If Not IsEmpty(plan) Then
                    Set record = CreateObject("Scripting.Dictionary")
                    record.Add "plan", plan
                    record.Add "transform", transform
                    slots.Add feature.Name, record
                    Log "FEATURE OPEN SLOT " & feature.Name & "; width=" & plan(SLOT_WIDTH_INDEX) & "; exposed depth=" & plan(SLOT_DEPTH_INDEX)
                    Exit Sub
                End If
            End If
        End If
    Next key
End Sub

Private Function NativePlacementSide(ByVal display As SldWorks.DisplayDimension) As Long
    If Not featureDimensionPlacement Is Nothing Then
        If featureDimensionPlacement.Exists(display.GetNameForSelection) Then
            NativePlacementSide = CLng(featureDimensionPlacement(display.GetNameForSelection))
        End If
    End If
End Function

Private Sub DirectOpenSlotDiameters(ByVal d As SldWorks.ModelDoc2)
    Dim sheetName As Variant, view As SldWorks.View, display As SldWorks.DisplayDimension, hasSlot As Boolean, key As String
    For Each sheetName In AsDrawing(d).GetSheetNames
        AsDrawing(d).ActivateSheet CStr(sheetName)
        Set view = AsDrawing(d).GetFirstView.GetNextView
        Do While Not view Is Nothing
            hasSlot = False
            Set display = view.GetFirstDisplayDimension5
            Do While Not display Is Nothing
                If featureDimensionOwners.Exists(display.GetNameForSelection) Then
                    key = CStr(featureDimensionOwners(display.GetNameForSelection))
                    If Left$(key, Len("open-slot-width|")) = "open-slot-width|" Then
                        hasSlot = True
                    End If
                End If
                Set display = display.GetNext5
            Loop
            If hasSlot Then
                Set display = view.GetFirstDisplayDimension5
                Do While Not display Is Nothing
                    If display.Type2 = SW_DIMENSION_DIAMETER Then
                        key = CStr(featureDimensionOwners(display.GetNameForSelection))
                        If Left$(key, Len("pattern-circle|")) = "pattern-circle|" Then
                            featureDimensionPlacement(display.GetNameForSelection) = NATIVE_LABEL_SIDE_RIGHT
                        Else
                            featureDimensionPlacement(display.GetNameForSelection) = NATIVE_LABEL_SIDE_LEFT
                        End If
                    End If
                    Set display = display.GetNext5
                Loop
            End If
            Set view = view.GetNextView
        Loop
    Next sheetName
End Sub

Private Sub PlaceDirectedPitchLabel(ByVal d As SldWorks.ModelDoc2, ByVal display As SldWorks.DisplayDimension, ByVal view As SldWorks.View, ByVal region As Variant)
    Dim outline As Variant, ann As SldWorks.Annotation, bounds As Variant, position As Variant
    Dim references As Variant, textFormat As SldWorks.TextFormat
    Dim targetX As Double, targetY As Double, attempt As Long
    references = display.GetDimension2(SW_DIMENSION_INDEX_FIRST).ReferencePoints
    If IsEmpty(references) Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "A directed pitch-circle dimension has no center reference."
    End If
    outline = view.GetOutline
    Set ann = display.GetAnnotation
    Set textFormat = ann.GetTextFormat(SW_ANNOTATION_TEXT_FORMAT_FIRST)
    textFormat.CharHeight = TEXT_HEIGHT_SMALL_M
    ann.SetTextFormat SW_ANNOTATION_TEXT_FORMAT_FIRST, False, textFormat
    display.SolidLeader = False
    display.OffsetText = False
    display.SetBentLeaderLength False, LEADER_SHOULDER_LENGTH_M
    targetX = outline(OUTLINE_RIGHT_INDEX) + PITCH_LABEL_RIGHT_CLEARANCE_M
    targetY = outline(OUTLINE_TOP_INDEX) + PITCH_LABEL_TOP_CLEARANCE_M
    ' SetPosition2 moves the native annotation origin, not its rendered text box.
    ' Refit the actual text bounds after the diameter leader changes direction.
    For attempt = 1 To LEADER_CLAMP_ATTEMPTS
        d.GraphicsRedraw2
        bounds = NativeTextBounds(display.GetDisplayData)
        targetX = MinD(outline(OUTLINE_RIGHT_INDEX) + PITCH_LABEL_RIGHT_CLEARANCE_M, region(OUTLINE_RIGHT_INDEX) - bounds(OUTLINE_RIGHT_INDEX) + bounds(OUTLINE_LEFT_INDEX))
        targetY = MinD(outline(OUTLINE_TOP_INDEX) + PITCH_LABEL_TOP_CLEARANCE_M, region(OUTLINE_TOP_INDEX) - bounds(OUTLINE_TOP_INDEX) + bounds(OUTLINE_BOTTOM_INDEX))
        position = ann.GetPosition
        If Abs(bounds(OUTLINE_LEFT_INDEX) - targetX) < DIMENSION_GEOMETRY_GAP_M And Abs(bounds(OUTLINE_BOTTOM_INDEX) - targetY) < DIMENSION_GEOMETRY_GAP_M Then
            Exit For
        End If
        ann.SetPosition2 position(0) + targetX - bounds(OUTLINE_LEFT_INDEX), position(1) + targetY - bounds(OUTLINE_BOTTOM_INDEX), position(2)
    Next attempt
End Sub

Private Function CollectOpenBoreSlots(ByVal model As SldWorks.ModelDoc2) As Object
    Dim result As Object, feature As SldWorks.Feature, child As SldWorks.Feature
    Set result = CreateObject("Scripting.Dictionary")
    Set feature = model.FirstFeature
    Do While Not feature Is Nothing
        If feature.GetTypeName = FEATURE_LEGACY_CUT Then
            Set child = feature.GetFirstSubFeature
            Do While Not child Is Nothing
                CollectOpenBoreSketch child, result
                Set child = child.GetNextSubFeature
            Loop
        End If
        Set feature = feature.GetNextFeature
    Loop
    Set CollectOpenBoreSlots = result
End Function

Private Function OpenSlotVertex(ByVal view As SldWorks.View, ByVal transform As Variant, ByVal target As Variant) As SldWorks.Vertex
    Dim components As Variant, component As Variant, edges As Variant, item As Variant, edge As SldWorks.Edge
    Dim vertex As SldWorks.Vertex, point As Variant, endIndex As Long
    components = view.GetVisibleComponents
    If IsEmpty(components) Then
        Exit Function
    End If
    For Each component In components
        edges = view.GetVisibleEntities2(component, SW_VISIBLE_ENTITY_EDGE)
        If Not IsEmpty(edges) Then
            For Each item In edges
                Set edge = item
                For endIndex = EDGE_START_INDEX To EDGE_END_INDEX
                    Set vertex = edge.GetStartVertex
                    If endIndex = EDGE_END_INDEX Then
                        Set vertex = edge.GetEndVertex
                    End If
                    If Not vertex Is Nothing Then
                        point = SourcePointInSketch(vertex.GetPoint, transform)
                        If Abs(point(0) - target(0)) < NATIVE_GEOMETRY_KEY_TOLERANCE_M And Abs(point(1) - target(1)) < NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
                            Set OpenSlotVertex = vertex
                            Exit Function
                        End If
                    End If
                Next endIndex
            Next item
        End If
    Next component
End Function

Private Function AddOpenSlotSpan(ByVal d As SldWorks.ModelDoc2, ByVal view As SldWorks.View, ByVal first As SldWorks.Vertex, ByVal last As SldWorks.Vertex, ByVal value As Double, ByVal key As String, ByVal above As Boolean) As SldWorks.DisplayDimension
    Dim a As Variant, b As Variant, outline As Variant, display As SldWorks.DisplayDimension
    a = PatternSheetPoint(view, first.GetPoint)
    b = PatternSheetPoint(view, last.GetPoint)
    outline = view.GetOutline
    d.ClearSelection2 True
    If Not SelectFeatureReference(d, view, first) Or Not SelectFeatureReference(d, view, last, True) Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot select open slot vertices."
    End If
    If Abs(a(0) - b(0)) > Abs(a(1) - b(1)) Then
        Set display = d.AddHorizontalDimension2((a(0) + b(0)) / 2, outline(3) + GOST_OBJECT_OFFSET_M, 0)
    Else
        Set display = d.AddVerticalDimension2(outline(0) - GOST_OBJECT_OFFSET_M, (a(1) + b(1)) / 2, 0)
    End If
    If display Is Nothing Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot create an open slot dimension."
    End If
    If Abs(display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue - value) > NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Open slot dimension does not match its model references."
    End If
    featureDimensionOwners(display.GetNameForSelection) = key
    If above Then
        featureDimensionPlacement(display.GetNameForSelection) = NATIVE_LABEL_SIDE_ABOVE
    End If
    Set AddOpenSlotSpan = display
End Function

Private Sub AddOpenBoreSlots(ByVal d As SldWorks.ModelDoc2, ByVal slots As Object)
    Dim name As Variant, record As Object, plan As Variant, transform As Variant, matrix As Variant, sheetName As Variant
    Dim view As SldWorks.View, target As Variant, first As SldWorks.Vertex, last As SldWorks.Vertex, floorVertex As SldWorks.Vertex
    Dim axis As Long, across As Long, aligned As Double, display As SldWorks.DisplayDimension, added As Boolean
    For Each name In slots.Keys
        Set record = slots(name)
        plan = record("plan")
        transform = record("transform")
        axis = CLng(plan(SLOT_RADIAL_AXIS_INDEX))
        across = SCREEN_Y_AXIS_INDEX - axis
        added = False
        For Each sheetName In AsDrawing(d).GetSheetNames
            AsDrawing(d).ActivateSheet CStr(sheetName)
            Set view = AsDrawing(d).GetFirstView.GetNextView
            Do While Not view Is Nothing
                matrix = view.ModelToViewTransform.ArrayData
                aligned = transform(MATRIX_XZ_INDEX) * matrix(MATRIX_XZ_INDEX) + transform(MATRIX_YZ_INDEX) * matrix(MATRIX_YZ_INDEX) + transform(MATRIX_ZZ_INDEX) * matrix(MATRIX_ZZ_INDEX)
                If aligned > FACE_NORMAL_ALIGNMENT_COSINE Then
                    target = Array(0#, 0#, 0#)
                    target(axis) = plan(SLOT_MOUTH_INDEX)
                    target(across) = plan(SLOT_TRANSVERSE_LOW_INDEX)
                    Set first = OpenSlotVertex(view, transform, target)
                    target(across) = plan(SLOT_TRANSVERSE_HIGH_INDEX)
                    Set last = OpenSlotVertex(view, transform, target)
                    target(axis) = plan(SLOT_FLOOR_INDEX)
                    target(across) = plan(SLOT_TRANSVERSE_LOW_INDEX)
                    Set floorVertex = OpenSlotVertex(view, transform, target)
                    If Not first Is Nothing And Not last Is Nothing And Not floorVertex Is Nothing Then
                        Set display = AddOpenSlotSpan(d, view, first, last, CDbl(plan(SLOT_WIDTH_INDEX)), "open-slot-width|" & name, True)
                        Set display = AddOpenSlotSpan(d, view, first, floorVertex, CDbl(plan(SLOT_DEPTH_INDEX)), "open-slot-depth|" & name, False)
                        added = True
                        Exit Do
                    End If
                End If
                Set view = view.GetNextView
            Loop
            If added Then
                Exit For
            End If
        Next sheetName
        If Not added Then
            Err.Raise vbObjectError + ERR_FEATURE_VIEW_REQUIRED, , "No visible circle/slot intersection for " & name
        End If
    Next name
End Sub
