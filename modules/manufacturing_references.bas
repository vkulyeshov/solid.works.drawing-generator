Attribute VB_Name = "manufacturing_references"
Option Explicit
Private Function ExactCircleKey(ByVal edge As SldWorks.Edge) As String
    Dim curve As SldWorks.Curve, parameters As Variant
    Set curve = edge.GetCurve
    If curve.IsCircle Then
        parameters = curve.CircleParams
        ExactCircleKey = NativePointKey(Array(parameters(0), parameters(1), parameters(2))) & NativeDirectionKey(Array(parameters(3), parameters(4), parameters(5))) & "|" & CStr(Round(parameters(CIRCLE_RADIUS_INDEX) / NATIVE_GEOMETRY_KEY_TOLERANCE_M, 0))
    End If
End Function

Private Function QueueModelDiameter(ByVal display As SldWorks.DisplayDimension, ByVal view As SldWorks.View, ByVal source As String, ByVal kind As String, ByVal symmetry As Object, ByVal requests As Object) As Boolean
    Dim key As String, request As Object, dimension As SldWorks.Dimension, radius As Double, parameters As Variant
    If display.Type2 <> SW_DIMENSION_DIAMETER And Not (display.Type2 = SW_DIMENSION_RADIAL And kind = FEATURE_TYPE_REVOLUTION) Then
        Exit Function
    End If
    key = NativeCircularOwnershipKey(display, view, symmetry)
    If Left$(key, Len(CIRCULAR_SIZE_KEY_PREFIX)) <> CIRCULAR_SIZE_KEY_PREFIX Then
        Exit Function
    End If
    If Not requests.Exists(key) Then
        Set dimension = display.GetDimension2(SW_DIMENSION_INDEX_FIRST)
        radius = dimension.SystemValue / CIRCULAR_DIAMETER_RADIUS_RATIO
        If display.Type2 = SW_DIMENSION_RADIAL Then
            radius = dimension.SystemValue
        End If
        Set request = CreateObject("Scripting.Dictionary")
        request.Add "radius", radius
        request.Add "source", source
        request.Add "kind", kind
        For Each parameters In symmetry("cylinders")
            If Abs(parameters(CIRCLE_RADIUS_INDEX) - radius) < NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
                If CircularAxisKey(parameters, radius, symmetry) = key Then
                    request.Add "axis", parameters
                    Exit For
                End If
            End If
        Next parameters
        requests.Add key, request
    End If
    QueueModelDiameter = True
End Function

Private Sub RemoveCoveredRadialWalls(ByVal d As SldWorks.ModelDoc2, ByVal requests As Object)
    Dim sheetName As Variant, view As SldWorks.View, display As SldWorks.DisplayDimension, following As SldWorks.DisplayDimension
    Dim dimension As SldWorks.Dimension, refs As Variant, first As SldWorks.MathPoint, last As SldWorks.MathPoint
    Dim a As Variant, b As Variant, parameters As Variant, radialA As Variant, radialB As Variant, normal As Variant
    Dim key As Variant, other As Variant, request As Object, stationA As Double, stationB As Double, radiusA As Double, radiusB As Double
    Dim dot As Double, axis As Long, covered As Boolean, prefix As String
    For Each sheetName In AsDrawing(d).GetSheetNames
        AsDrawing(d).ActivateSheet CStr(sheetName)
        Set view = AsDrawing(d).GetFirstView.GetNextView
        Do While Not view Is Nothing
            Set display = view.GetFirstDisplayDimension5
            Do While Not display Is Nothing
                Set following = display.GetNext5
                covered = False
                If display.Type2 = SW_DIMENSION_LINEAR Then
                    Set dimension = display.GetDimension2(SW_DIMENSION_INDEX_FIRST)
                    refs = dimension.ReferencePoints
                    If Not IsEmpty(refs) Then
                        If UBound(refs) >= DIMENSION_REFERENCE_SECOND - 1 Then
                            Set first = refs(DIMENSION_REFERENCE_FIRST - 1)
                            Set last = refs(DIMENSION_REFERENCE_SECOND - 1)
                            a = first.ArrayData
                            b = last.ArrayData
                            For Each key In requests.Keys
                                Set request = requests(key)
                                If request.Exists("axis") Then
                                    parameters = request("axis")
                                    normal = Array(parameters(CIRCLE_AXIS_X_INDEX), parameters(CIRCLE_AXIS_Y_INDEX), parameters(CIRCLE_AXIS_Z_INDEX))
                                    stationA = 0
                                    stationB = 0
                                    For axis = 0 To SPACE_LAST_AXIS_INDEX
                                        stationA = stationA + (a(axis) - parameters(axis)) * normal(axis)
                                        stationB = stationB + (b(axis) - parameters(axis)) * normal(axis)
                                    Next axis
                                    If Abs(stationA - stationB) < NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
                                        radialA = Array(0, 0, 0)
                                        radialB = Array(0, 0, 0)
                                        radiusA = 0
                                        radiusB = 0
                                        dot = 0
                                        For axis = 0 To SPACE_LAST_AXIS_INDEX
                                            radialA(axis) = a(axis) - parameters(axis) - stationA * normal(axis)
                                            radialB(axis) = b(axis) - parameters(axis) - stationB * normal(axis)
                                            radiusA = radiusA + radialA(axis) ^ 2
                                            radiusB = radiusB + radialB(axis) ^ 2
                                            dot = dot + radialA(axis) * radialB(axis)
                                        Next axis
                                        radiusA = Sqr(radiusA)
                                        radiusB = Sqr(radiusB)
                                        If radiusA > NATIVE_GEOMETRY_KEY_TOLERANCE_M And radiusB > NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
                                            If Abs(dot / (radiusA * radiusB) - 1) < DIRECTION_ZERO_TOLERANCE And Abs(dimension.SystemValue - Abs(radiusA - radiusB)) < NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
                                                If Abs(radiusA - CDbl(request("radius"))) < NATIVE_GEOMETRY_KEY_TOLERANCE_M Or Abs(radiusB - CDbl(request("radius"))) < NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
                                                    prefix = Left$(CStr(key), InStrRev(CStr(key), "|"))
                                                    For Each other In requests.Keys
                                                        If CStr(other) <> CStr(key) And Left$(CStr(other), Len(prefix)) = prefix Then
                                                            If Abs(radiusA - CDbl(requests(other)("radius"))) < NATIVE_GEOMETRY_KEY_TOLERANCE_M Or Abs(radiusB - CDbl(requests(other)("radius"))) < NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
                                                                covered = True
                                                                Exit For
                                                            End If
                                                        End If
                                                    Next other
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                                If covered Then
                                    Exit For
                                End If
                            Next key
                        End If
                    End If
                    If covered Then
                        RemoveFeatureDimension d, display, "radial wall is defined by the two coaxial diameters"
                    End If
                End If
                Set display = following
            Loop
            Set view = view.GetNextView
        Loop
    Next sheetName
End Sub

Private Sub AddRequestedModelDiameters(ByVal d As SldWorks.ModelDoc2, ByVal model As SldWorks.ModelDoc2, ByVal requests As Object, ByVal covered As Object, ByVal axes As Object, ByVal symmetry As Object)
    Dim key As Variant, request As Object, sheets As Variant, sheetName As Variant, view As SldWorks.View
    Dim components As Variant, component As Variant, edges As Variant, item As Variant, edge As SldWorks.Edge, curve As SldWorks.Curve
    Dim parameters As Variant, transform As Variant, aligned As Double, quality As Long, bestQuality As Long, depth As Double, bestDepth As Double
    Dim bestView As SldWorks.View, bestEdge As SldWorks.Edge, bestSheet As String, display As SldWorks.DisplayDimension, outline As Variant
    sheets = AsDrawing(d).GetSheetNames
    For Each key In requests.Keys
        Set request = requests(key)
        Set bestEdge = Nothing
        bestQuality = -1
        bestDepth = -SEARCH_INITIAL_LIMIT
        For Each sheetName In sheets
            AsDrawing(d).ActivateSheet CStr(sheetName)
            Set view = AsDrawing(d).GetFirstView.GetNextView
            Do While Not view Is Nothing
                transform = view.ModelToViewTransform.ArrayData
                quality = FeatureCircularQuality(FEATURE_GEOMETRY_PREFIX & CStr(key), view, symmetry, CStr(sheetName) <> CStr(sheets(LBound(sheets))))
                If request("kind") = FEATURE_TYPE_INSTANT_EXTRUDE Then
                    quality = quality + ProfileViewPriority(model, CStr(request("source")), view)
                End If
                components = view.GetVisibleComponents
                If Not IsEmpty(components) Then
                    For Each component In components
                        edges = view.GetVisibleEntities2(component, SW_VISIBLE_ENTITY_EDGE)
                        If Not IsEmpty(edges) Then
                            For Each item In edges
                                Set edge = item
                                Set curve = edge.GetCurve
                                If Not curve Is Nothing Then
                                    If curve.IsCircle Then
                                        parameters = curve.CircleParams
                                        aligned = Abs(parameters(CIRCLE_AXIS_X_INDEX) * transform(MATRIX_XZ_INDEX) + parameters(CIRCLE_AXIS_Y_INDEX) * transform(MATRIX_YZ_INDEX) + parameters(CIRCLE_AXIS_Z_INDEX) * transform(MATRIX_ZZ_INDEX))
                                        If aligned >= FACE_NORMAL_ALIGNMENT_COSINE Then
                                            If CircularAxisKey(parameters, parameters(CIRCLE_RADIUS_INDEX), symmetry) = CStr(key) Then
                                                depth = parameters(0) * transform(MATRIX_XZ_INDEX) + parameters(1) * transform(MATRIX_YZ_INDEX) + parameters(2) * transform(MATRIX_ZZ_INDEX)
                                                If quality > bestQuality Or (quality = bestQuality And depth > bestDepth) Then
                                                    bestQuality = quality
                                                    bestDepth = depth
                                                    Set bestEdge = edge
                                                    Set bestView = view
                                                    bestSheet = CStr(sheetName)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next item
                        End If
                    Next component
                End If
                Set view = view.GetNextView
            Loop
        Next sheetName
        If bestEdge Is Nothing Then
            Err.Raise vbObjectError + ERR_FEATURE_VIEW_REQUIRED, , "No visible model edge supports a requested diameter."
        End If
        AsDrawing(d).ActivateSheet bestSheet
        If Not SelectFeatureEdge(d, bestView, bestEdge) Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot select the requested diameter edge."
        End If
        outline = bestView.GetOutline
        Set display = d.AddDiameterDimension2(outline(2) + DIAMETER_LABEL_OFFSET_M, outline(3), 0)
        If display Is Nothing Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot create the requested model diameter."
        End If
        If Not ValidateDimension(d, bestView, display, CIRCULAR_DIAMETER_RADIUS_RATIO * CDbl(request("radius")) * bestView.ScaleDecimal) Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "A model diameter does not match its supporting edge."
        End If
        featureDimensionOwners.Add display.GetNameForSelection, FEATURE_GEOMETRY_PREFIX & CStr(key)
        covered(CStr(key)) = True
        If request("kind") = FEATURE_TYPE_REVOLUTION Then
            axes(Left$(CStr(key), InStrRev(CStr(key), "|"))) = True
        End If
        Log "FEATURE MODEL DIAMETER value=" & display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue & "; view=" & bestView.Name
    Next key
End Sub

Private Function ModelCircleClassifications(ByVal view As SldWorks.View) As Object
    Dim result As Object, bodies As Variant, body As Variant, model As SldWorks.ModelDoc2, edges As Variant, item As Variant
    Dim edge As SldWorks.Edge, curve As SldWorks.Curve, record As Object, key As String
    Set result = CreateObject("Scripting.Dictionary")
    Set model = view.ReferencedDocument
    bodies = AsPart(model).GetBodies2(SW_SOLID_BODY, True)
    For Each body In bodies
        edges = body.GetEdges
        For Each item In edges
            Set edge = item
            Set curve = edge.GetCurve
            If curve.IsCircle Then
                key = ExactCircleKey(edge)
                If Not result.Exists(key) Then
                    Set record = CreateObject("Scripting.Dictionary")
                    record.Add "bore", CircleIsBore(edge)
                    record.Add "rim", DerivedChamferRim(edge)
                    record.Add "parameters", curve.CircleParams
                    result.Add key, record
                    If CBool(record("rim")) Then
                        Log "FEATURE CHAMFER RIM " & key
                    End If
                End If
            End If
        Next item
    Next body
    Set ModelCircleClassifications = result
End Function

Private Function CircleIsBore(ByVal edge As SldWorks.Edge) As Boolean
    Dim faces As Variant, item As Variant, face As SldWorks.Face2, surface As SldWorks.Surface
    Dim key As String
    If Not featureCircleClassifications Is Nothing Then
        key = ExactCircleKey(edge)
        If featureCircleClassifications.Exists(key) Then
            CircleIsBore = CBool(featureCircleClassifications(key)("bore"))
            Exit Function
        End If
    End If
    faces = edge.GetTwoAdjacentFaces2
    If IsEmpty(faces) Then
        Exit Function
    End If
    For Each item In faces
        If Not item Is Nothing Then
            Set face = item
            Set surface = face.GetSurface
            If surface.IsCylinder And face.FaceInSurfaceSense Then
                CircleIsBore = True
                Exit Function
            End If
        End If
    Next item
End Function

Private Function DerivedChamferRim(ByVal edge As SldWorks.Edge) As Boolean
    Dim faces As Variant, item As Variant, face As SldWorks.Face2, surface As SldWorks.Surface, cone As Boolean, cylinder As Boolean
    Dim key As String
    If Not featureCircleClassifications Is Nothing Then
        key = ExactCircleKey(edge)
        If featureCircleClassifications.Exists(key) Then
            DerivedChamferRim = CBool(featureCircleClassifications(key)("rim"))
            Exit Function
        End If
    End If
    faces = edge.GetTwoAdjacentFaces2
    If IsEmpty(faces) Then
        Exit Function
    End If
    For Each item In faces
        If Not item Is Nothing Then
            Set face = item
            Set surface = face.GetSurface
            cone = cone Or surface.IsCone
            cylinder = cylinder Or surface.IsCylinder
        End If
    Next item
    DerivedChamferRim = cone And Not cylinder
End Function

Private Function ProfileViewPriority(ByVal model As SldWorks.ModelDoc2, ByVal source As String, ByVal view As SldWorks.View) As Long
    Dim feature As SldWorks.Feature, child As SldWorks.Feature, sketch As SldWorks.Sketch, transform As Variant, matrix As Variant
    Dim normal As Variant, aligned As Double, horizontal As Double, vertical As Double, profile As Boolean
    Set feature = model.FeatureByName(source)
    If feature Is Nothing Then
        Exit Function
    End If
    profile = feature.GetTypeName2 = FEATURE_TYPE_SKETCH
    If profile Then
        Set sketch = feature.GetSpecificFeature2
    ElseIf feature.GetTypeName2 = FEATURE_TYPE_INSTANT_EXTRUDE Then
        Set child = feature.GetFirstSubFeature
        Do While Not child Is Nothing
            If child.GetTypeName2 = FEATURE_TYPE_SKETCH Then
                Set sketch = child.GetSpecificFeature2
                Exit Do
            End If
            Set child = child.GetNextSubFeature
        Loop
    End If
    If sketch Is Nothing Then
        Exit Function
    End If
    transform = sketch.ModelToSketchTransform.ArrayData
    normal = Array(transform(MATRIX_XZ_INDEX), transform(MATRIX_YZ_INDEX), transform(MATRIX_ZZ_INDEX))
    matrix = view.ModelToViewTransform.ArrayData
    aligned = normal(0) * matrix(MATRIX_XZ_INDEX) + normal(1) * matrix(MATRIX_YZ_INDEX) + normal(2) * matrix(MATRIX_ZZ_INDEX)
    If profile Then
        If aligned > FACE_NORMAL_ALIGNMENT_COSINE Then
            ProfileViewPriority = FEATURE_PROFILE_FRONT_PRIORITY
        End If
    Else
        horizontal = normal(0) * matrix(MATRIX_XX_INDEX) + normal(1) * matrix(MATRIX_YX_INDEX) + normal(2) * matrix(MATRIX_ZX_INDEX)
        vertical = normal(0) * matrix(MATRIX_XY_INDEX) + normal(1) * matrix(MATRIX_YY_INDEX) + normal(2) * matrix(MATRIX_ZY_INDEX)
        If horizontal > FACE_NORMAL_ALIGNMENT_COSINE Or vertical > FACE_NORMAL_ALIGNMENT_COSINE Then
            ProfileViewPriority = FEATURE_PROFILE_FRONT_PRIORITY
        End If
    End If
End Function

Private Function ConstraintBoundaryKey(ByVal segment As SldWorks.SketchSegment, ByVal center As Variant) As String
    Dim line As SldWorks.SketchLine, a As Variant, b As Variant, candidate As String, turn As Long
    If segment.GetType <> SW_SKETCH_LINE Or segment.ConstructionGeometry Then
        Exit Function
    End If
    Set line = segment
    a = SourceSketchPoint(line.GetStartPoint2)
    b = SourceSketchPoint(line.GetEndPoint2)
    ConstraintBoundaryKey = NativePairKey(a, b)
    For turn = 1 To QUARTER_TURN_EQUIVALENT_COUNT - 1
        a = QuarterTurnPoint(a, center)
        b = QuarterTurnPoint(b, center)
        candidate = NativePairKey(a, b)
        If candidate < ConstraintBoundaryKey Then
            ConstraintBoundaryKey = candidate
        End If
    Next turn
End Function

Private Function ConstructionHalfCoverage(ByVal model As SldWorks.ModelDoc2, ByVal owners As Object) As Object
    Dim result As Object, source As Variant, feature As SldWorks.Feature, evidence As Object, records As Collection
    Dim display As SldWorks.DisplayDimension, dimension As SldWorks.Dimension, attached As Variant, item As Variant
    Dim segment As SldWorks.SketchSegment, record As Object, half As Variant, full As Variant, boundary As String
    Dim centers As Variant, hasDatum As Boolean, hasBoundary As Boolean, boundaries As Object, sourceBoundary As Variant
    Dim annotation As SldWorks.Annotation
    Set result = CreateObject("Scripting.Dictionary")
    Set featureDimensionOccurrenceCounts = CreateObject("Scripting.Dictionary")
    For Each source In owners.Keys
        Set feature = model.FeatureByName(CStr(source))
        If Not feature Is Nothing Then
            If feature.GetTypeName2 = FEATURE_TYPE_SKETCH Then
                Set evidence = SketchQuarterEvidence(feature)
                If CBool(evidence("valid")) Then
                    Set records = New Collection
                    Set display = feature.GetFirstDisplayDimension
                    Do While Not display Is Nothing
                        If display.Type2 = SW_DIMENSION_LINEAR Then
                            Set dimension = display.GetDimension2(SW_DIMENSION_INDEX_FIRST)
                            Set annotation = display.GetAnnotation
                            attached = annotation.GetAttachedEntities3
                            hasDatum = False
                            Set boundaries = CreateObject("Scripting.Dictionary")
                            If Not IsEmpty(attached) Then
                                For Each item In attached
                                    If TypeOf item Is SldWorks.SketchSegment Then
                                        Set segment = item
                                        hasDatum = hasDatum Or segment.ConstructionGeometry
                                        boundary = ConstraintBoundaryKey(segment, evidence("center"))
                                        If Len(boundary) > 0 Then
                                            boundaries(boundary) = True
                                        End If
                                    End If
                                Next item
                            End If
                            Set record = CreateObject("Scripting.Dictionary")
                            record.Add "name", dimension.FullName
                            record.Add "value", dimension.SystemValue
                            record.Add "datum", hasDatum
                            record.Add "centered", CenteredConstructionHalf(attached, evidence("center"))
                            record.Add "boundaries", boundaries
                            records.Add record
                            If Not hasDatum And boundaries.Count > 0 Then
                                featureDimensionOccurrenceCounts(dimension.FullName) = QUARTER_TURN_EQUIVALENT_COUNT
                            End If
                        End If
                        Set display = feature.GetNextDisplayDimension(display)
                    Loop
                    For Each half In records
                        If CBool(half("datum")) And CBool(half("centered")) And CDbl(half("value")) > NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
                            For Each full In records
                                If Not CBool(full("datum")) Then
                                    If Abs(CDbl(full("value")) - CIRCULAR_DIAMETER_RADIUS_RATIO * CDbl(half("value"))) < NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
                                        For Each sourceBoundary In half("boundaries").Keys
                                            If full("boundaries").Exists(CStr(sourceBoundary)) Then
                                                result(half("name")) = CStr(full("name"))
                                                Exit For
                                            End If
                                        Next sourceBoundary
                                    End If
                                End If
                            Next full
                        End If
                    Next half
                End If
            End If
        End If
    Next source
    Set ConstructionHalfCoverage = result
End Function

Private Function CenteredConstructionHalf(ByVal entities As Variant, ByVal center As Variant) As Boolean
    Dim item As Variant, other As Variant, datumSegment As SldWorks.SketchSegment, boundarySegment As SldWorks.SketchSegment
    Dim datumLine As SldWorks.SketchLine, boundaryLine As SldWorks.SketchLine
    Dim a As Variant, b As Variant, c As Variant, e As Variant, lengthA As Double, lengthB As Double, dot As Double
    If IsEmpty(entities) Then
        Exit Function
    End If
    For Each item In entities
        If TypeOf item Is SldWorks.SketchSegment Then
            Set datumSegment = item
            If datumSegment.GetType = SW_SKETCH_LINE And datumSegment.ConstructionGeometry Then
                Set datumLine = datumSegment
                a = SourceSketchPoint(datumLine.GetStartPoint2)
                b = SourceSketchPoint(datumLine.GetEndPoint2)
                lengthA = Sqr((b(0) - a(0)) ^ 2 + (b(1) - a(1)) ^ 2)
                If lengthA > NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
                    If Abs((b(0) - a(0)) * (center(1) - a(1)) - (b(1) - a(1)) * (center(0) - a(0))) / lengthA <= NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
                        For Each other In entities
                            If TypeOf other Is SldWorks.SketchSegment Then
                                Set boundarySegment = other
                                If boundarySegment.GetType = SW_SKETCH_LINE And Not boundarySegment.ConstructionGeometry Then
                                    Set boundaryLine = boundarySegment
                                    c = SourceSketchPoint(boundaryLine.GetStartPoint2)
                                    e = SourceSketchPoint(boundaryLine.GetEndPoint2)
                                    lengthB = Sqr((e(0) - c(0)) ^ 2 + (e(1) - c(1)) ^ 2)
                                    If lengthB > NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
                                        dot = ((b(0) - a(0)) * (e(0) - c(0)) + (b(1) - a(1)) * (e(1) - c(1))) / (lengthA * lengthB)
                                        If Abs(Abs(dot) - 1) < DIRECTION_ZERO_TOLERANCE Then
                                            CenteredConstructionHalf = True
                                            Exit Function
                                        End If
                                    End If
                                End If
                            End If
                        Next other
                    End If
                End If
            End If
        End If
    Next item
End Function
