Attribute VB_Name = "dimensions"
Option Explicit
Private Function DiameterDimensions(ByVal d As SldWorks.ModelDoc2, ByVal v As SldWorks.View) As Long
    Dim comps As Variant, edges As Variant, ci As Variant, ei As Variant, edge As SldWorks.Edge
    Dim curve As SldWorks.Curve, cp As Variant, start As SldWorks.Vertex, finish As SldWorks.Vertex
    Dim p As Variant, q As Variant, axis(2) As Double, vector As SldWorks.MathVector, util As SldWorks.MathUtility
    Dim seen As Object, key As String, radius As Double, outline As Variant, dd As SldWorks.DisplayDimension
    Dim group As Object, centers As Object, groupName As Variant, center(2) As Double, point As SldWorks.MathPoint, centerKey As String
    On Error GoTo Failed
    Set seen = CreateObject("Scripting.Dictionary")
    Set util = sw.GetMathUtility
    comps = v.GetVisibleComponents: outline = v.GetOutline
    If IsEmpty(comps) Then
        Exit Function
    End If
    For Each ci In comps
        edges = v.GetVisibleEntities2(ci, SW_VISIBLE_ENTITY_EDGE)
        If Not IsEmpty(edges) Then
            For Each ei In edges
                Set edge = ei: Set curve = edge.GetCurve
                If curve.IsCircle Then
                    cp = curve.CircleParams: radius = cp(CIRCLE_RADIUS_INDEX)
                    If radius < MIN_CIRCLE_RADIUS_M Then
                        GoTo NextCircle
                    End If
                    Set start = edge.GetStartVertex: Set finish = edge.GetEndVertex
                    If Not start Is Nothing And Not finish Is Nothing Then
                        p = start.GetPoint: q = finish.GetPoint
                        If Abs(p(0) - q(0)) + Abs(p(1) - q(1)) + Abs(p(2) - q(2)) > CIRCLE_CLOSURE_TOLERANCE_M Then
                            GoTo NextCircle
                        End If
                    End If
                    axis(0) = cp(CIRCLE_AXIS_X_INDEX): axis(1) = cp(CIRCLE_AXIS_Y_INDEX): axis(2) = cp(CIRCLE_AXIS_Z_INDEX)
                    Set vector = util.CreateVector(axis)
                    Set vector = vector.MultiplyTransform(v.ModelToViewTransform)
                    p = vector.ArrayData
                    If Abs(p(2)) < FACE_NORMAL_ALIGNMENT_COSINE * Sqr(p(0) ^ 2 + p(1) ^ 2 + p(2) ^ 2) Then
                        GoTo NextCircle
                    End If
                    key = "diameter|" & PhysicalAxisKey(v, 2) & "|" & Format$(radius, GEOMETRY_KEY_FORMAT)
                    If dimensionKeys.Exists(key) Then
                        GoTo NextCircle
                    End If
                    If Not seen.Exists(key) Then
                        Set group = CreateObject("Scripting.Dictionary")
                        Set centers = CreateObject("Scripting.Dictionary")
                        group.Add "edge", edge
                        group.Add "radius", radius
                        group.Add "centers", centers
                        seen.Add key, group
                    End If
                    Set group = seen(key)
                    Set centers = group("centers")
                    center(0) = cp(CIRCLE_CENTER_X_INDEX): center(1) = cp(CIRCLE_CENTER_Y_INDEX): center(2) = cp(CIRCLE_CENTER_Z_INDEX)
                    Set point = util.CreatePoint(center)
                    Set point = point.MultiplyTransform(v.ModelToViewTransform)
                    p = point.ArrayData
                    centerKey = CStr(Round(p(0) / v.ScaleDecimal, COORDINATE_KEY_DECIMALS)) & "|" & CStr(Round(p(1) / v.ScaleDecimal, COORDINATE_KEY_DECIMALS))
                    If Not centers.Exists(centerKey) Then
                        centers.Add centerKey, True
                    End If
                End If
NextCircle:
            Next ei
        End If
    Next ci
    For Each groupName In seen.Keys
        Set group = seen(groupName)
        Set centers = group("centers")
        Set edge = group("edge")
        radius = CDbl(group("radius"))
        d.ClearSelection2 True: AsDrawing(d).ActivateView ""
        If v.SelectEntity(edge, False) Then
            Set dd = d.AddDiameterDimension2(outline(OUTLINE_RIGHT_INDEX) + DIAMETER_LABEL_OFFSET_M, outline(OUTLINE_TOP_INDEX) - DIAMETER_LABEL_PITCH_M * DiameterDimensions, 0)
            If Not dd Is Nothing Then
                If ValidateDimension(d, v, dd, 2 * radius * v.ScaleDecimal) Then
                    If centers.Count > 1 Then
                        dd.SetText SW_DIMENSION_TEXT_PREFIX, CStr(centers.Count) & "x " & dd.GetText(SW_DIMENSION_TEXT_PREFIX)
                    End If
                    dimensionKeys.Add CStr(groupName), True
                    DiameterDimensions = DiameterDimensions + 1
                End If
            End If
        End If
    Next groupName
    d.ClearSelection2 True
    Exit Function
Failed:
    Warn "Diameter dimensions " & v.Name & ": " & Err.Description
End Function

Private Function OverallDimensions(ByVal d As SldWorks.ModelDoc2, ByVal v As SldWorks.View, Optional ByVal lowerRight As Boolean = False) As Long
    Dim comps As Variant, verts As Variant, i As Long, k As Long, p As Variant, mp As SldWorks.MathPoint
    Dim edges As Variant, edgeItem As Variant, edge As SldWorks.Edge, vertexList As Collection, endpoint As SldWorks.Vertex
    Dim leftV As Object, rightV As Object, topV As Object, bottomV As Object, vertex As SldWorks.Vertex
    Dim comp As SldWorks.Component2, math As SldWorks.MathUtility, stage As String
    Dim refDoc As SldWorks.ModelDoc2, realComp As SldWorks.Component2, modelComps As Object
    Dim ca As Variant, ci As Variant, compName As String, partDoc As SldWorks.ModelDoc2, sourceEntity As Object
    Dim xmin As Double, xmax As Double, ymin As Double, ymax As Double, outline As Variant, dimObj As Object
    Dim horizontalY As Double, verticalX As Double, horizontalKey As String, verticalKey As String
    On Error GoTo Failed
    stage = "components"
    Set math = sw.GetMathUtility
    Set refDoc = v.ReferencedDocument
    Set modelComps = CreateObject("Scripting.Dictionary"): modelComps.CompareMode = 1
    If refDoc.GetType = SW_DOC_ASSEMBLY Then
        ca = AsAssembly(refDoc).GetComponents(False)
        For Each ci In ca
            Set realComp = ci
            modelComps.Add realComp.Name2, realComp
        Next ci
    End If
    xmin = SEARCH_INITIAL_LIMIT: ymin = SEARCH_INITIAL_LIMIT: xmax = -SEARCH_INITIAL_LIMIT: ymax = -SEARCH_INITIAL_LIMIT
    comps = v.GetVisibleComponents
    If IsEmpty(comps) Then
        Exit Function
    End If
    For i = LBound(comps) To UBound(comps)
        Set comp = comps(i)
        Set realComp = Nothing
        If refDoc.GetType = SW_DOC_ASSEMBLY Then
            compName = Mid$(comp.Name2, InStr(comp.Name2, "/") + 1)
            If modelComps.Exists(compName) Then
                Set realComp = modelComps(compName)
            End If
        End If
        Set vertexList = New Collection
            edges = v.GetVisibleEntities2(comp, SW_VISIBLE_ENTITY_EDGE)
            If Not IsEmpty(edges) Then
                For Each edgeItem In edges
                    Set edge = edgeItem
                    Set endpoint = edge.GetStartVertex
                    If Not endpoint Is Nothing Then
                        vertexList.Add endpoint
                    End If
                    Set endpoint = edge.GetEndVertex
                    If Not endpoint Is Nothing Then
                        vertexList.Add endpoint
                    End If
                Next edgeItem
            End If
        If vertexList.Count > 0 Then
            For k = 1 To vertexList.Count
                Set vertex = vertexList(k)
                stage = "point"
                Set mp = math.CreatePoint(vertex.GetPoint)
                Dim corresponding As Object
                stage = "transform"
                Set corresponding = vertex
                If refDoc.GetType = SW_DOC_ASSEMBLY Then
                    If realComp Is Nothing Then
                        GoTo NextVertex
                    End If
                    Set mp = mp.MultiplyTransform(realComp.Transform2)
                End If
                Set mp = mp.MultiplyTransform(v.ModelToViewTransform)
                p = mp.ArrayData
                stage = "corresponding"
                If refDoc.GetType = SW_DOC_ASSEMBLY Then
                    If p(0) >= xmin And p(0) <= xmax And p(1) >= ymin And p(1) <= ymax Then
                        GoTo NextVertex
                    End If
                    Set partDoc = realComp.GetModelDoc2
                    Set sourceEntity = partDoc.Extension.GetCorresponding2(vertex)
                    If sourceEntity Is Nothing Then
                        Set sourceEntity = vertex
                    End If
                    Set corresponding = realComp.GetCorrespondingEntity(sourceEntity)
                    If corresponding Is Nothing Then
                        GoTo NextVertex
                    End If
                End If
                If p(0) < xmin Then
                    xmin = p(0)
                    Set leftV = corresponding
                End If
                If p(0) > xmax Then
                    xmax = p(0)
                    Set rightV = corresponding
                End If
                If p(1) < ymin Then
                    ymin = p(1)
                    Set bottomV = corresponding
                End If
                If p(1) > ymax Then
                    ymax = p(1)
                    Set topV = corresponding
                End If
NextVertex:
            Next k
        End If
        IncludeCircleExtremes v, edges, realComp, refDoc, xmin, xmax, ymin, ymax, leftV, rightV, bottomV, topV
    Next i
    outline = v.GetOutline
    horizontalY = outline(OUTLINE_TOP_INDEX) + OVERALL_LABEL_OFFSET_M: verticalX = outline(OUTLINE_LEFT_INDEX) - OVERALL_LABEL_OFFSET_M
    If lowerRight Then
        outline = ProjectedFlatBounds(v)
        horizontalY = COORDINATE_BOTTOM_LIMIT_M - OVERALL_FOOTER_CLEARANCE_M
        verticalX = sheetWidth - OVERALL_RIGHT_MARGIN_M
    End If
    stage = "select horizontal"
    horizontalKey = "overall|" & CStr(lowerRight) & "|" & PhysicalAxisKey(v, 0) & "|" & Format$((xmax - xmin) / v.ScaleDecimal, GEOMETRY_KEY_FORMAT)
    verticalKey = "overall|" & CStr(lowerRight) & "|" & PhysicalAxisKey(v, 1) & "|" & Format$((ymax - ymin) / v.ScaleDecimal, GEOMETRY_KEY_FORMAT)
    AsDrawing(d).ActivateView "": d.ClearSelection2 True
    If Not leftV Is Nothing And xmax - xmin > MIN_DIMENSION_SPAN_M And Not dimensionKeys.Exists(horizontalKey) Then
        If SelectInView(d, v, leftV, False) And SelectInView(d, v, rightV, True) Then
            stage = "add horizontal"
            Set dimObj = d.AddHorizontalDimension2((outline(OUTLINE_LEFT_INDEX) + outline(OUTLINE_RIGHT_INDEX)) / 2, horizontalY, 0)
            If Not dimObj Is Nothing Then
                SetMaximumArcEnds dimObj
                stage = "precision horizontal"
                dimObj.SetPrecision3 LINEAR_DECIMAL_PLACES, LINEAR_DECIMAL_PLACES, LINEAR_DECIMAL_PLACES, LINEAR_DECIMAL_PLACES
                stage = "validate horizontal"
                If ValidateDimension(d, v, dimObj, xmax - xmin) Then
                    dimensionKeys.Add horizontalKey, dimObj
                    OverallDimensions = OverallDimensions + 1
                End If
            End If
        End If
    End If
    d.ClearSelection2 True
    stage = "select vertical"
    If Not bottomV Is Nothing And ymax - ymin > MIN_DIMENSION_SPAN_M And Not dimensionKeys.Exists(verticalKey) Then
        If SelectInView(d, v, bottomV, False) And SelectInView(d, v, topV, True) Then
            stage = "add vertical"
            Set dimObj = d.AddVerticalDimension2(verticalX, (outline(OUTLINE_BOTTOM_INDEX) + outline(OUTLINE_TOP_INDEX)) / 2, 0)
            If Not dimObj Is Nothing Then
                SetMaximumArcEnds dimObj
                stage = "precision vertical"
                dimObj.SetPrecision3 LINEAR_DECIMAL_PLACES, LINEAR_DECIMAL_PLACES, LINEAR_DECIMAL_PLACES, LINEAR_DECIMAL_PLACES
                stage = "validate vertical"
                If ValidateDimension(d, v, dimObj, ymax - ymin) Then
                    dimensionKeys.Add verticalKey, dimObj
                    OverallDimensions = OverallDimensions + 1
                End If
            End If
        End If
    End If
    d.ClearSelection2 True
    Exit Function
Failed:
    Warn "Overall dimensions " & v.Name & " [" & stage & "]: " & Err.Description
End Function

Private Sub SetMaximumArcEnds(ByVal display As SldWorks.DisplayDimension)
    Dim dimension As SldWorks.Dimension
    Set dimension = display.GetDimension2(SW_DIMENSION_INDEX_FIRST)
    If dimension Is Nothing Then
        Exit Sub
    End If
    dimension.SetArcEndCondition 1, SW_ARC_MAXIMUM
    dimension.SetArcEndCondition 2, SW_ARC_MAXIMUM
End Sub

Private Sub IncludeCircleExtremes(ByVal v As SldWorks.View, ByVal edges As Variant, ByVal comp As SldWorks.Component2, ByVal doc As SldWorks.ModelDoc2, ByRef xmin As Double, ByRef xmax As Double, ByRef ymin As Double, ByRef ymax As Double, ByRef leftV As Object, ByRef rightV As Object, ByRef bottomV As Object, ByRef topV As Object)
    Dim item As Variant, edge As SldWorks.Edge, curve As SldWorks.Curve, cp As Variant, p As Variant, q As Variant
    Dim first As SldWorks.Vertex, last As SldWorks.Vertex, util As SldWorks.MathUtility, point As SldWorks.MathPoint
    Dim axis As SldWorks.MathVector, entity As Object, sourceEntity As Object, part As SldWorks.ModelDoc2, radius As Double
    Dim center(2) As Double, direction(2) As Double
    If IsEmpty(edges) Then
        Exit Sub
    End If
    Set util = sw.GetMathUtility
    For Each item In edges
        Set edge = item: Set curve = edge.GetCurve
        If curve.IsCircle Then
            Set first = edge.GetStartVertex: Set last = edge.GetEndVertex
            If Not first Is Nothing And Not last Is Nothing Then
                p = first.GetPoint: q = last.GetPoint
                If Abs(p(0) - q(0)) + Abs(p(1) - q(1)) + Abs(p(2) - q(2)) > CIRCLE_CLOSURE_TOLERANCE_M Then
                    GoTo NextCircleBound
                End If
            End If
            cp = curve.CircleParams
            center(0) = cp(CIRCLE_CENTER_X_INDEX): center(1) = cp(CIRCLE_CENTER_Y_INDEX): center(2) = cp(CIRCLE_CENTER_Z_INDEX)
            direction(0) = cp(CIRCLE_AXIS_X_INDEX): direction(1) = cp(CIRCLE_AXIS_Y_INDEX): direction(2) = cp(CIRCLE_AXIS_Z_INDEX)
            Set point = util.CreatePoint(center)
            Set axis = util.CreateVector(direction)
            Set entity = edge
            If doc.GetType = SW_DOC_ASSEMBLY Then
                If comp Is Nothing Then
                    GoTo NextCircleBound
                End If
                Set point = point.MultiplyTransform(comp.Transform2)
                Set axis = axis.MultiplyTransform(comp.Transform2)
                Set part = comp.GetModelDoc2
                Set sourceEntity = part.Extension.GetCorresponding2(edge)
                If sourceEntity Is Nothing Then
                    Set sourceEntity = edge
                End If
                Set entity = comp.GetCorrespondingEntity(sourceEntity)
                If entity Is Nothing Then
                    GoTo NextCircleBound
                End If
            End If
            Set point = point.MultiplyTransform(v.ModelToViewTransform): p = point.ArrayData
            Set axis = axis.MultiplyTransform(v.ModelToViewTransform): q = axis.ArrayData
            If Abs(q(2)) < FACE_NORMAL_ALIGNMENT_COSINE * Sqr(q(0) ^ 2 + q(1) ^ 2 + q(2) ^ 2) Then
                GoTo NextCircleBound
            End If
            radius = cp(CIRCLE_RADIUS_INDEX) * v.ScaleDecimal
            If p(0) - radius < xmin Then
                xmin = p(0) - radius
                Set leftV = entity
            End If
            If p(0) + radius > xmax Then
                xmax = p(0) + radius
                Set rightV = entity
            End If
            If p(1) - radius < ymin Then
                ymin = p(1) - radius
                Set bottomV = entity
            End If
            If p(1) + radius > ymax Then
                ymax = p(1) + radius
                Set topV = entity
            End If
        End If
NextCircleBound:
    Next item
End Sub

Private Function SelectInView(ByVal d As SldWorks.ModelDoc2, ByVal v As Object, ByVal entity As Object, ByVal append As Boolean) As Boolean
    Dim e As SldWorks.Entity, typedView As SldWorks.View
    Set typedView = v
    SelectInView = typedView.SelectEntity(entity, append)
End Function

Private Function ValidateDimension(ByVal d As SldWorks.ModelDoc2, ByVal v As Object, ByVal display As Object, ByVal paperSpan As Double) As Boolean
    Dim dimension As SldWorks.Dimension, actual As Double, expected As Double
    Dim typedDisplay As SldWorks.DisplayDimension
    Set typedDisplay = display
    Set dimension = typedDisplay.GetDimension2(SW_DIMENSION_INDEX_FIRST)
    If dimension Is Nothing Then
        d.ForceRebuild3 False
        Set dimension = typedDisplay.GetDimension2(SW_DIMENSION_INDEX_FIRST)
    End If
    If dimension Is Nothing Then
        Err.Raise vbObjectError + ERR_DIMENSION_PARAMETER, , "New dimension has no parameter."
    End If
    actual = dimension.SystemValue
    expected = paperSpan / v.ScaleDecimal
    If Abs(actual - expected) <= DIMENSION_MATCH_TOLERANCE_M Then
        ValidateDimension = True
    Else
        Warn "Dimension rejected: " & v.Name & " measured=" & actual & " expected=" & expected
        d.ClearSelection2 True
        Dim ann As SldWorks.Annotation
        Set ann = typedDisplay.GetAnnotation
        If ann.Select3(False, Nothing) = False Then
            Err.Raise vbObjectError + ERR_REMOVE_DIMENSION, , "Cannot select an invalid dimension for removal."
        End If
        d.EditDelete
    End If
End Function

Private Sub NormalizeDimensions(ByVal d As SldWorks.ModelDoc2)
    Dim sheets As Variant, sheetName As Variant, v As SldWorks.View
    Dim display As SldWorks.DisplayDimension, ann As SldWorks.Annotation, tf As SldWorks.TextFormat
    sheets = AsDrawing(d).GetSheetNames
    Set tf = d.GetUserPreferenceTextFormat(SW_TEXT_FORMAT_DIMENSION)
    tf.TypeFaceName = Setting("Font", DEFAULT_FONT)
    tf.CharHeight = TEXT_HEIGHT_STANDARD_M
    For Each sheetName In sheets
        AsDrawing(d).ActivateSheet CStr(sheetName)
        Set v = AsDrawing(d).GetFirstView
        Set v = v.GetNextView
        Do While Not v Is Nothing
            Set display = v.GetFirstDisplayDimension5
            Do While Not display Is Nothing
                display.SetPrecision3 LINEAR_DECIMAL_PLACES, LINEAR_DECIMAL_PLACES, LINEAR_DECIMAL_PLACES, LINEAR_DECIMAL_PLACES
                ApplyGostDimensionStyle display
                Set ann = display.GetAnnotation
                ann.SetTextFormat SW_ANNOTATION_TEXT_FORMAT_FIRST, False, tf
                Set display = display.GetNext5
            Loop
            Set v = v.GetNextView
        Loop
    Next sheetName
End Sub

Private Sub ApplyGostDimensionStyle(ByVal display As SldWorks.DisplayDimension)
    Dim firstStyle As Long, secondStyle As Long, annotation As SldWorks.Annotation, i As Long
    ' GOST 2.307: retain complete extension lines and place arrows by actual fit.
    display.SmartWitness = False
    display.ExtensionLineUseDocumentDisplay = True
    display.GetArrowHeadStyle2 firstStyle, secondStyle
    If firstStyle = SW_ARROW_DOT And secondStyle = SW_ARROW_DOT Then
        display.ArrowSide = SW_DIMENSION_ARROWS_INSIDE
        Set annotation = display.GetAnnotation
        For i = ARROW_HEAD_INDEX_FIRST To annotation.GetArrowHeadCount - 1
            annotation.SetArrowHeadSizeAtIndex i, False, GOST_CHAIN_DOT_DIAMETER_M, GOST_CHAIN_DOT_DIAMETER_M, GOST_CHAIN_DOT_DIAMETER_M
        Next i
    Else
        display.SetArrowHeadStyle2 False, SW_ARROW_CLOSED, SW_ARROW_CLOSED
    End If
    display.SetBrokenLeader2 False, SW_SOLID_LEADER_ALIGNED_TEXT
End Sub

