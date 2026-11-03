Attribute VB_Name = "semantic_dimensions"
Option Explicit
Private Function FeatureDimensionMode() As Boolean
    Dim strategy As String
    strategy = LCase$(Setting("DimensionStrategy", DIMENSION_STRATEGY_FEATURES))
    If strategy <> DIMENSION_STRATEGY_FEATURES And strategy <> DIMENSION_STRATEGY_CHAIN Then
        Err.Raise vbObjectError + ERR_DRAWING_SETTINGS, , "DimensionStrategy must be features or chain."
    End If
    FeatureDimensionMode = strategy = DIMENSION_STRATEGY_FEATURES
End Function

Private Sub RemoveFeatureDimension(ByVal d As SldWorks.ModelDoc2, ByVal display As SldWorks.DisplayDimension, ByVal reason As String)
    Dim ann As SldWorks.Annotation
    Log "FEATURE OMIT " & display.GetNameForSelection & "; " & reason
    d.ClearSelection2 True
    Set ann = display.GetAnnotation
    If Not ann.Select3(False, Nothing) Then
        Err.Raise vbObjectError + ERR_REMOVE_DIMENSION, , "Cannot select a redundant feature annotation."
    End If
    d.EditDelete
End Sub

Private Sub FeatureOwnerNames(ByVal feature As SldWorks.Feature, ByVal kind As String, ByVal owners As Object)
    Dim child As SldWorks.Feature
    owners(feature.Name) = kind
    Set child = feature.GetFirstSubFeature
    Do While Not child Is Nothing
        FeatureOwnerNames child, kind, owners
        Set child = child.GetNextSubFeature
    Loop
End Sub

Private Function FeatureOwners(ByVal model As SldWorks.ModelDoc2) As Object
    Dim owners As Object, feature As SldWorks.Feature
    Set owners = CreateObject("Scripting.Dictionary")
    Set feature = model.FirstFeature
    Do While Not feature Is Nothing
        FeatureOwnerNames feature, feature.GetTypeName2, owners
        Set feature = feature.GetNextFeature
    Loop
    Set FeatureOwners = owners
End Function

Private Function FeatureDimensionSheets(ByVal d As SldWorks.ModelDoc2) As Long
    Dim sheetName As Variant, imported As Variant, drawing As Object, view As SldWorks.View
    Dim display As SldWorks.DisplayDimension, following As SldWorks.DisplayDimension, dimension As SldWorks.Dimension
    Dim key As String, value As Double, seen As Object, audit As Object, typeName As String
    Dim active As SldWorks.ModelDoc2, error As Long, extension As Object, selected As Boolean, matrix As Variant
    Dim evidence As Object, owners As Object, symmetry As Object, model As SldWorks.ModelDoc2, parts As Variant, sourceFeature As String, geometric As String
    Dim sourceKeys As Object, halfConstraints As Object, circleSizes As Object, turningAxes As Object
    Dim sketchEvidence As Object, rotatedKeys As Object, rotationKey As String
    Dim constructionCoverage As Object, constructionName As Variant
    Dim diameterRequests As Object, openSlots As Object, chamfers As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Set sourceKeys = CreateObject("Scripting.Dictionary")
    Set circleSizes = CreateObject("Scripting.Dictionary")
    Set turningAxes = CreateObject("Scripting.Dictionary")
    Set sketchEvidence = CreateObject("Scripting.Dictionary")
    Set rotatedKeys = CreateObject("Scripting.Dictionary")
    Set diameterRequests = CreateObject("Scripting.Dictionary")
    Set featureDimensionOwners = CreateObject("Scripting.Dictionary")
    Set featureDimensionRepeatedKeys = CreateObject("Scripting.Dictionary")
    Set featureDimensionViewPriorities = CreateObject("Scripting.Dictionary")
    Set featureDimensionPlacement = CreateObject("Scripting.Dictionary")
    Set drawing = AsDrawing(d)
    Set view = drawing.GetFirstView.GetNextView
    Set model = view.ReferencedDocument
    Set featureCircleClassifications = ModelCircleClassifications(view)
    Set evidence = CollectHolePatterns(model)
    Set owners = FeatureOwners(model)
    Set openSlots = CollectOpenBoreSlots(model)
    Set chamfers = CollectChamferSpecifications(model)
    Set constructionCoverage = ConstructionHalfCoverage(model, owners)
    Set symmetry = NativeSymmetryEvidence(view)
    Set active = sw.ActivateDoc3(d.GetPathName, False, SW_REBUILD_USER_DECISION, error)
    DoEvents
    Set extension = d.Extension
    For Each sheetName In drawing.GetSheetNames
        drawing.ActivateSheet CStr(sheetName)
        Set view = drawing.GetFirstView.GetNextView
        Do While Not view Is Nothing
            matrix = view.ModelToViewTransform.ArrayData
            If Abs(matrix(MATRIX_XZ_INDEX)) > FACE_NORMAL_ALIGNMENT_COSINE Or Abs(matrix(MATRIX_YZ_INDEX)) > FACE_NORMAL_ALIGNMENT_COSINE Or Abs(matrix(MATRIX_ZZ_INDEX)) > FACE_NORMAL_ALIGNMENT_COSINE Then
                drawing.ActivateView ""
                d.ClearSelection2 True
                selected = extension.SelectByID2(CStr(view.Name), "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
                imported = AsDrawing(d).InsertModelAnnotations4(SW_IMPORT_ENTIRE_MODEL, SW_MODEL_ITEMS_MARKED_DIMENSIONS Or SW_MODEL_ITEMS_UNMARKED_DIMENSIONS Or SW_MODEL_ITEMS_HOLE_PROFILES Or SW_MODEL_ITEMS_HOLE_LOCATIONS, False, True, True, False, False, False)
                Log "FEATURE IMPORT " & view.Name & "; selected=" & CStr(selected) & "; annotations=" & CStr(Not IsEmpty(imported))
                d.ForceRebuild3 False
            End If
            Set view = view.GetNextView
        Loop
    Next sheetName
    Set halfConstraints = FeatureHalfConstraints(d, symmetry)
    For Each constructionName In constructionCoverage.Keys
        halfConstraints(constructionName) = CStr(constructionCoverage(constructionName))
    Next constructionName
    Set audit = fs.OpenTextFile(fs.BuildPath(runDir, FEATURE_DIMENSIONS_FILENAME), FILE_FOR_APPENDING, True, FILE_UNICODE)
    For Each sheetName In drawing.GetSheetNames
        drawing.ActivateSheet CStr(sheetName)
        Set view = drawing.GetFirstView.GetNextView
        Do While Not view Is Nothing
            Set display = view.GetFirstDisplayDimension5
            Do While Not display Is Nothing
                Set following = display.GetNext5
                Set dimension = display.GetDimension2(SW_DIMENSION_INDEX_FIRST)
                value = dimension.SystemValue
                key = "model|" & dimension.FullName
                parts = Split(dimension.FullName, "@")
                sourceFeature = CStr(parts(SOURCE_DIMENSION_FEATURE_INDEX))
                audit.WriteLine CStr(sheetName) & vbTab & view.Name & vbTab & display.GetNameForSelection & vbTab & dimension.FullName & vbTab & CStr(display.Type2) & vbTab & CStr(value)
                If openSlots.Exists(sourceFeature) Then
                    RemoveFeatureDimension d, display, "open slot uses exposed circle intersections instead of rectangle construction offsets"
                ElseIf ChamferDistanceCovered(display, sourceFeature, chamfers) Then
                    RemoveFeatureDimension d, display, "distance is included in the chamfer specification"
                ElseIf evidence("members").Exists(sourceFeature) Then
                    RemoveFeatureDimension d, display, "hole size is covered by the shared profile callout"
                ElseIf halfConstraints.Exists(dimension.FullName) Then
                    RemoveFeatureDimension d, display, "verified symmetry and full span " & CStr(halfConstraints(dimension.FullName))
                ElseIf Abs(value) < NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
                    RemoveFeatureDimension d, display, "zero construction constraint"
                ElseIf display.Type2 = SW_DIMENSION_ANGULAR And Abs(value - FULL_REVOLUTION_RAD) < DIRECTION_ZERO_TOLERANCE Then
                    RemoveFeatureDimension d, display, "complete revolution"
                ElseIf QueueModelDiameter(display, view, sourceFeature, CStr(owners(sourceFeature)), symmetry, diameterRequests) Then
                    RemoveFeatureDimension d, display, "diameter is recreated from a visible model edge"
                    FeatureDimensionSheets = FeatureDimensionSheets + 1
                Else
                    FormatChamferSpecification display, sourceFeature, chamfers
                    If sourceKeys.Exists(dimension.FullName) Then
                        key = CStr(sourceKeys(dimension.FullName))
                    Else
                        geometric = NativeCircularOwnershipKey(display, view, symmetry)
                        If Len(geometric) = 0 Then
                            geometric = NativeMeasurementKey(display, view, symmetry)
                        End If
                        If Len(geometric) > 0 Then
                            key = "feature|" & geometric
                        End If
                        If Left$(geometric, Len(CIRCULAR_SIZE_KEY_PREFIX)) = CIRCULAR_SIZE_KEY_PREFIX Then
                            circleSizes(geometric) = True
                            If owners.Exists(sourceFeature) Then
                                If owners(sourceFeature) = FEATURE_TYPE_REVOLUTION Then
                                    turningAxes(Left$(geometric, InStrRev(geometric, "|"))) = True
                                End If
                            End If
                        End If
                        rotationKey = SketchRotationKey(display, model, sourceFeature, sketchEvidence)
                        If Len(rotationKey) > 0 Then
                            If rotatedKeys.Exists(rotationKey) Then
                                key = CStr(rotatedKeys(rotationKey))
                                featureDimensionRepeatedKeys(key) = True
                                Log "FEATURE LOCAL REPEAT " & dimension.FullName & "; matching rotated references in the same sketch"
                            Else
                                rotatedKeys.Add rotationKey, key
                            End If
                        End If
                        sourceKeys.Add dimension.FullName, key
                    End If
                    featureDimensionOwners.Add display.GetNameForSelection, key
                    If owners.Exists(sourceFeature) Then
                        If owners(sourceFeature) = FEATURE_TYPE_INSTANT_EXTRUDE Then
                            featureDimensionViewPriorities(display.GetNameForSelection) = ProfileViewPriority(model, sourceFeature, view)
                        End If
                    End If
                    FeatureDimensionSheets = FeatureDimensionSheets + 1
                End If
                Set display = following
            Loop
            Set view = view.GetNextView
        Loop
    Next sheetName
    audit.Close
    AddSharedHoleSizes d, evidence
    AddRequestedModelDiameters d, model, diameterRequests, circleSizes, turningAxes, symmetry
    AddCoaxialCircleSizes d, circleSizes, turningAxes, symmetry
    AddHolePatternLocations d, evidence
    RemoveCoveredRadialWalls d, diameterRequests
    AddOpenBoreSlots d, openSlots
    DirectOpenSlotDiameters d
    Log "FEATURE DIMENSIONS imported=" & FeatureDimensionSheets
    If FeatureDimensionSheets = 0 Then
        Log "FEATURE FALLBACK source dimensions unavailable; using bounded native geometry dimensions"
    End If
End Function
