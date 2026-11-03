Attribute VB_Name = "six_views"
Option Explicit
Private Function SixViewMode() As Boolean
    Dim value As String
    value = LCase$(Setting("OrthographicViews", "auto"))
    If value = "auto" Then
        SixViewMode = autoExpandedViews
        Exit Function
    End If
    If value <> CStr(ORTHOGRAPHIC_VIEW_COUNT_STANDARD) And value <> CStr(ORTHOGRAPHIC_VIEW_COUNT_SIX) Then
        Err.Raise vbObjectError + ERR_DRAWING_SETTINGS, , "OrthographicViews must be auto, 3 or 6."
    End If
    SixViewMode = value = CStr(ORTHOGRAPHIC_VIEW_COUNT_SIX)
End Function

Private Function SixPrimaryOrientation(ByVal principal As String, ByVal side As Boolean) As String
    If side Then
        If principal = "*Right" Then
            SixPrimaryOrientation = "*Front"
        Else
            SixPrimaryOrientation = "*Right"
        End If
    Else
        If principal = "*Top" Then
            SixPrimaryOrientation = "*Front"
        Else
            SixPrimaryOrientation = "*Top"
        End If
    End If
End Function

Private Function OppositeViewTransform(ByVal source As Variant) As Variant
    Dim values As Variant, index As Long
    values = source
    ' A half-turn about screen Y preserves handedness and reverses the viewing side.
    For index = 0 To MATRIX_ZZ_INDEX Step MATRIX_ROW_STRIDE
        values(index) = -values(index)
        values(index + MATRIX_XZ_INDEX) = -values(index + MATRIX_XZ_INDEX)
    Next index
    For index = SIX_VIEW_FIRST_TRANSLATION_INDEX To SIX_VIEW_DEPTH_TRANSLATION_INDEX
        values(index) = 0
    Next index
    values(MATRIX_SCALE_INDEX) = SIX_VIEW_MODEL_VIEW_SCALE
    OppositeViewTransform = values
End Function

Private Function CircularQuantity(ByVal signature As String) As Long
    Dim index As Long, character As String
    CircularQuantity = 1
    For index = 1 To Len(signature)
        character = Mid$(signature, index, 1)
        If InStr("0123456789", character) = 0 Then
            If index > 1 And (character = "X" Or character = ChrW(CIRCULAR_MULTIPLICATION_CHARACTER)) Then
                CircularQuantity = CLng(Left$(signature, index - 1))
            End If
            Exit Function
        End If
    Next index
End Function

Private Function CircularSpecification(ByVal signature As String) As String
    Dim index As Long
    CircularSpecification = signature
    If CircularQuantity(signature) > 1 Then
        index = InStr(signature, "X")
        If index = 0 Then
            index = InStr(signature, ChrW(CIRCULAR_MULTIPLICATION_CHARACTER))
        End If
        CircularSpecification = Mid$(signature, index + 1)
    End If
End Function

Private Function CircularAxisKey(ByVal parameters As Variant, ByVal radius As Double, ByVal symmetry As Object) As String
    Dim origin As Variant, normal As Variant, station As Double, index As Long, flags As Variant, reflected As Variant, key As String, candidate As String
    normal = Array(parameters(3), parameters(4), parameters(5))
    station = parameters(0) * normal(0) + parameters(1) * normal(1) + parameters(2) * normal(2)
    origin = Array(parameters(0) - station * normal(0), parameters(1) - station * normal(1), parameters(2) - station * normal(2))
    key = NativePointKey(origin)
    flags = symmetry("axes")
    For index = 0 To 2
        If flags(index) Then
            reflected = NativeReflectedPoint(origin, index, symmetry("center"))
            candidate = NativePointKey(reflected)
            If candidate < key Then
                key = candidate
            End If
        End If
    Next index
    CircularAxisKey = "circular-size|" & NativeDirectionKey(normal) & "|" & key & "|" & CStr(Round(radius / NATIVE_GEOMETRY_KEY_TOLERANCE_M, 0))
End Function

Private Function MakeOppositeView(ByVal d As SldWorks.ModelDoc2, ByVal doc As SldWorks.ModelDoc2, ByVal cfg As String, ByVal body As Object, ByVal source As SldWorks.View, ByVal x As Double, ByVal y As Double) As SldWorks.View
    Dim active As SldWorks.ModelDoc2, modelView As SldWorks.ModelView, util As SldWorks.MathUtility
    Dim values As Variant, name As String, error As Long, view As SldWorks.View
    Set active = sw.ActivateDoc3(fs.GetFileName(doc.GetPathName), False, SW_REBUILD_USER_DECISION, error)
    If active Is Nothing Then
        Err.Raise vbObjectError + ERR_VIEW_ACTIVATION, , "Cannot activate the snapshot for an opposite view."
    End If
    Set modelView = doc.ActiveView
    Set util = sw.GetMathUtility
    values = OppositeViewTransform(source.ModelToViewTransform.ArrayData)
    modelView.Orientation3 = util.CreateTransform(values)
    viewCounter = viewCounter + 1
    name = SIX_VIEW_ORIENTATION_NAME & CStr(viewCounter)
    doc.NameView name
    Set active = sw.ActivateDoc3(d.GetTitle, False, SW_REBUILD_USER_DECISION, error)
    Set view = MakeView(d, doc, cfg, body, name, x, y)
    If view Is Nothing Then
        Err.Raise vbObjectError + ERR_MAIN_VIEW, , "Cannot create an opposite drawing view."
    End If
    Set MakeOppositeView = view
End Function

Private Sub AuditSixView(ByVal sheetName As String, ByVal view As SldWorks.View, ByVal role As String)
    Dim audit As Object, values As Variant, text As String, index As Long
    values = view.ModelToViewTransform.ArrayData
    For index = 0 To UBound(values)
        text = text & vbTab & CStr(values(index))
    Next index
    Set audit = fs.OpenTextFile(fs.BuildPath(runDir, SIX_VIEW_TRANSFORMS_FILENAME), FILE_FOR_APPENDING, True, FILE_UNICODE)
    audit.WriteLine CStr(jobNumber) & vbTab & sheetName & vbTab & view.Name & vbTab & role & text
    audit.Close
End Sub

Private Function NativeCalloutSignature(ByVal display As SldWorks.DisplayDimension) As String
    Dim data As SldWorks.DisplayData, index As Long, text As String
    Set data = display.GetDisplayData
    For index = 0 To data.GetTextCount - 1
        text = text & data.GetTextAtIndex(index)
    Next index
    NativeCalloutSignature = UCase$(Replace(Replace(Replace(Replace(text, " ", ""), vbTab, ""), vbCr, ""), vbLf, ""))
End Function

Private Function NativeCircularOwnershipKey(ByVal display As SldWorks.DisplayDimension, ByVal view As SldWorks.View, ByVal symmetry As Object) As String
    Dim ann As SldWorks.Annotation, entities As Variant, entity As Variant, edge As SldWorks.Edge
    Dim faces As Variant, item As Variant, face As SldWorks.Face2, surface As SldWorks.Surface, feature As SldWorks.Feature
    Dim curve As SldWorks.Curve, parameters As Variant, normal As Variant, featureName As String, key As String
    Dim dimension As SldWorks.Dimension, refs As Variant, point As SldWorks.MathPoint, position As Variant, radius As Double
    Dim signature As String, specification As String, ordinary As Boolean, collections As Variant, collection As Variant, geometry As Variant, radial As Double, station As Double
    signature = NativeCalloutSignature(display)
    specification = CircularSpecification(signature)
    ordinary = IsNumeric(Replace(specification, SW_DIMENSION_TOKEN_DIAMETER, "")) And InStr(specification, SW_DIMENSION_TOKEN_DIAMETER) > 0
    If display.Type2 <> SW_DIMENSION_RADIAL And display.Type2 <> SW_DIMENSION_DIAMETER And Not ordinary Then
        Exit Function
    End If
    ordinary = ordinary Or display.Type2 = SW_DIMENSION_RADIAL
    Set ann = display.GetAnnotation
    Set dimension = display.GetDimension2(SW_DIMENSION_INDEX_FIRST)
    radius = dimension.SystemValue / CIRCULAR_DIAMETER_RADIUS_RATIO
    If display.Type2 = SW_DIMENSION_RADIAL Then
        radius = dimension.SystemValue
    End If
    entities = ann.GetAttachedEntities3
    If Not IsEmpty(entities) Then
        For Each entity In entities
            If IsObject(entity) Then
                If TypeOf entity Is SldWorks.Edge Then
                    Set edge = entity
                    Set curve = edge.GetCurve
                    If curve.IsCircle Then
                        parameters = curve.CircleParams
                        normal = Array(parameters(CIRCLE_AXIS_X_INDEX), parameters(CIRCLE_AXIS_Y_INDEX), parameters(CIRCLE_AXIS_Z_INDEX))
                        If ordinary And Abs(parameters(CIRCLE_RADIUS_INDEX) - radius) <= NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
                            NativeCircularOwnershipKey = CircularAxisKey(parameters, radius, symmetry)
                            Exit Function
                        End If
                        faces = edge.GetTwoAdjacentFaces2
                        If Not IsEmpty(faces) Then
                            For Each item In faces
                                If Not item Is Nothing Then
                                    Set face = item
                                    Set surface = face.GetSurface
                                    If surface.IsCylinder Then
                                        Set feature = face.GetFeature
                                        If Not feature Is Nothing Then
                                            featureName = feature.Name
                                        End If
                                        Exit For
                                    End If
                                End If
                            Next item
                        End If
                        If Len(featureName) > 0 Then
                            key = "circular-feature|" & featureName & "|" & NativeDirectionKey(normal)
                        Else
                            key = "circular-axis|" & NativeDirectionKey(normal)
                            position = Array(parameters(CIRCLE_CENTER_X_INDEX), parameters(CIRCLE_CENTER_Y_INDEX), parameters(CIRCLE_CENTER_Z_INDEX))
                            key = key & "|" & NativePointKey(position)
                        End If
                        NativeCircularOwnershipKey = key & "|" & CStr(display.Type2) & "|" & specification
                        Exit Function
                    End If
                End If
            End If
        Next entity
    End If
    refs = dimension.ReferencePoints
    If Not IsEmpty(refs) Then
        For Each entity In refs
            If TypeOf entity Is SldWorks.MathPoint Then
                Set point = entity
                position = point.ArrayData
                If ordinary Then
                    collections = Array("circles", "cylinders")
                    For Each collection In collections
                        For Each geometry In symmetry(CStr(collection))
                            If Abs(geometry(CIRCLE_RADIUS_INDEX) - radius) <= NATIVE_GEOMETRY_KEY_TOLERANCE_M Then
                                radial = (position(0) - geometry(0)) ^ 2 + (position(1) - geometry(1)) ^ 2 + (position(2) - geometry(2)) ^ 2
                                station = (position(0) - geometry(0)) * geometry(3) + (position(1) - geometry(1)) * geometry(4) + (position(2) - geometry(2)) * geometry(5)
                                If radial - station * station <= NATIVE_GEOMETRY_KEY_TOLERANCE_M ^ 2 Then
                                    NativeCircularOwnershipKey = CircularAxisKey(geometry, radius, symmetry)
                                    Exit Function
                                End If
                            End If
                        Next geometry
                    Next collection
                End If
            End If
        Next entity
    End If
End Function

Private Sub AssignSixViewDimensions(ByVal d As SldWorks.ModelDoc2)
    Dim groups As Object, loads As Object, candidates As Collection, record As Object, winner As Object
    Dim sheetName As Variant, view As SldWorks.View, display As SldWorks.DisplayDimension, symmetry As Object
    Dim key As String, group As Variant, candidate As Variant, minimum As Long, value As Double, ann As SldWorks.Annotation
    Dim audit As Object, removed As Long, retained As Long, measure As Double, quantity As Long, maximumQuantity As Long
    Dim places As Object
    Dim priority As Long, maximumPriority As Long, sheetNames As Variant, placeCount As Long, sourceName As Variant
    Set groups = CreateObject("Scripting.Dictionary")
    Set loads = CreateObject("Scripting.Dictionary")
    sheetNames = AsDrawing(d).GetSheetNames
    For Each sheetName In sheetNames
        AsDrawing(d).ActivateSheet CStr(sheetName)
        Set view = AsDrawing(d).GetFirstView.GetNextView
        If symmetry Is Nothing Then
            Set symmetry = NativeSymmetryEvidence(view)
        End If
        Do While Not view Is Nothing
            loads.Add CStr(sheetName) & "|" & view.Name, 0
            Set display = view.GetFirstDisplayDimension5
            Do While Not display Is Nothing
                key = NativeCircularOwnershipKey(display, view, symmetry)
                If Not featureDimensionOwners Is Nothing Then
                    If featureDimensionOwners.Exists(display.GetNameForSelection) Then
                        key = CStr(featureDimensionOwners(display.GetNameForSelection))
                    End If
                End If
                If Len(key) = 0 Then
                    key = NativeMeasurementKey(display, view, symmetry)
                End If
                If Len(key) = 0 Then
                    key = "unmatched|" & CStr(sheetName) & "|" & display.GetNameForSelection
                End If
                Set record = CreateObject("Scripting.Dictionary")
                record.Add "sheet", CStr(sheetName)
                record.Add "view", view.Name
                record.Add "owner", CStr(sheetName) & "|" & view.Name
                record.Add "display", display
                record.Add "name", display.GetNameForSelection
                record.Add "value", display.GetDimension2(SW_DIMENSION_INDEX_FIRST).SystemValue
                measure = CDbl(record("value"))
                If Left$(key, Len("circular-size|")) = "circular-size|" And display.Type2 <> SW_DIMENSION_RADIAL Then
                    measure = measure / CIRCULAR_DIAMETER_RADIUS_RATIO
                End If
                record.Add "measure", measure
                record.Add "quantity", CircularQuantity(NativeCalloutSignature(display))
                priority = FeatureCircularQuality(key, view, symmetry, CStr(sheetName) <> CStr(sheetNames(LBound(sheetNames))))
                If Not featureDimensionViewPriorities Is Nothing Then
                    If featureDimensionViewPriorities.Exists(display.GetNameForSelection) Then
                        If CLng(featureDimensionViewPriorities(display.GetNameForSelection)) > priority Then
                            priority = CLng(featureDimensionViewPriorities(display.GetNameForSelection))
                        End If
                    End If
                End If
                record.Add "priority", priority
                If Not groups.Exists(key) Then
                    Set candidates = New Collection
                    groups.Add key, candidates
                End If
                Set candidates = groups(key)
                candidates.Add record
                Set display = display.GetNext5
            Loop
            Set view = view.GetNextView
        Loop
    Next sheetName
    Set audit = fs.OpenTextFile(fs.BuildPath(runDir, SIX_VIEW_AUDIT_FILENAME), FILE_FOR_APPENDING, True, FILE_UNICODE)
    For Each group In groups.Keys
        Set candidates = groups(group)
        Set winner = Nothing
        minimum = groups.Count + 1
        maximumQuantity = 0
        maximumPriority = -1
        For Each candidate In candidates
            Set record = candidate
            quantity = CLng(record("quantity"))
            priority = CLng(record("priority"))
            If priority > maximumPriority Or (priority = maximumPriority And quantity > maximumQuantity) Or (priority = maximumPriority And quantity = maximumQuantity And CLng(loads(record("owner"))) < minimum) Then
                minimum = CLng(loads(record("owner")))
                maximumQuantity = quantity
                maximumPriority = priority
                Set winner = record
            End If
        Next candidate
        loads(winner("owner")) = CLng(loads(winner("owner"))) + 1
        If Not featureDimensionRepeatedKeys Is Nothing Then
            If featureDimensionRepeatedKeys.Exists(CStr(group)) Then
                Set places = CreateObject("Scripting.Dictionary")
                For Each candidate In candidates
                    Set record = candidate
                    Set display = record("display")
                    places(display.GetDimension2(SW_DIMENSION_INDEX_FIRST).FullName) = True
                Next candidate
                If places.Count > 1 Then
                    placeCount = places.Count
                    If Not featureDimensionOccurrenceCounts Is Nothing Then
                        For Each sourceName In places.Keys
                            If featureDimensionOccurrenceCounts.Exists(sourceName) Then
                                If CLng(featureDimensionOccurrenceCounts(sourceName)) > placeCount Then
                                    placeCount = CLng(featureDimensionOccurrenceCounts(sourceName))
                                End If
                            End If
                        Next sourceName
                    End If
                    Set display = winner("display")
                    display.SetText SW_DIMENSION_TEXT_SUFFIX, display.GetText(SW_DIMENSION_TEXT_SUFFIX) & " (" & placeCount & " місця)"
                End If
            End If
        End If
        retained = retained + 1
        For Each candidate In candidates
            Set record = candidate
            Set display = record("display")
            value = CDbl(record("value"))
            If Abs(CDbl(record("measure")) - CDbl(winner("measure"))) > LEADER_VALUE_TOLERANCE_M Then
                Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "A six-view dimension group contains different measurements."
            End If
            If record Is winner Then
                audit.WriteLine "KEEP" & vbTab & record("sheet") & vbTab & record("view") & vbTab & record("name") & vbTab & CStr(group) & vbTab & CStr(value) & vbTab & CStr(record("measure"))
            Else
                AsDrawing(d).ActivateSheet CStr(record("sheet"))
                d.ClearSelection2 True
                Set ann = display.GetAnnotation
                If Not ann.Select3(False, Nothing) Then
                    Err.Raise vbObjectError + ERR_REMOVE_DIMENSION, , "Cannot select a repeated six-view dimension."
                End If
                d.EditDelete
                removed = removed + 1
                audit.WriteLine "REMOVE" & vbTab & record("sheet") & vbTab & record("view") & vbTab & record("name") & vbTab & CStr(group) & vbTab & CStr(value) & vbTab & winner("name") & vbTab & winner("sheet") & vbTab & CStr(record("measure"))
            End If
        Next candidate
    Next group
    audit.Close
    Log "SIX VIEW DIMENSIONS retained=" & retained & "; repeated=" & removed & "; each group has one owner across both sheets"
End Sub

Private Sub AddOppositeViewSheet(ByVal d As SldWorks.ModelDoc2, ByVal doc As SldWorks.ModelDoc2, ByVal cfg As String, ByVal body As Object, ByVal first As SldWorks.View, ByVal lower As SldWorks.View, ByVal side As SldWorks.View, ByVal iso As SldWorks.View, ByVal viewScaleValue As Double, ByVal qty As Long)
    Dim sheet As SldWorks.Sheet, name As String, sources As Collection, item As Variant
    Dim view As SldWorks.View, opposite As SldWorks.View, position As Variant, oppositeIso As SldWorks.View, count As Long
    Set sheet = AsDrawing(d).GetCurrentSheet
    name = sheet.GetName
    Set sources = New Collection
    sources.Add first
    sources.Add lower
    sources.Add side
    For Each item In sources
        Set view = item
        AuditSixView name, view, "orthographic"
    Next item
    AuditSixView name, iso, "isometric"
    NewFormattedSheet d, SIX_VIEW_SHEET_NAME
    Set sheet = AsDrawing(d).GetCurrentSheet
    SetPaperScale sheet, viewScaleValue
    For Each item In sources
        Set view = item
        position = view.Position
        Set opposite = MakeOppositeView(d, doc, cfg, body, view, position(0), position(1))
        opposite.UseSheetScale = SW_USE_SHEET_SCALE
        d.ForceRebuild3 False
        CenterView opposite, position(0), position(1)
        AuditSixView sheet.GetName, opposite, "orthographic"
    Next item
    position = iso.Position
    Set oppositeIso = MakeOppositeView(d, doc, cfg, body, iso, position(0), position(1))
    oppositeIso.UseSheetScale = SW_USE_INDEPENDENT_SCALE
    oppositeIso.ScaleDecimal = iso.ScaleDecimal
    d.ForceRebuild3 False
    CenterView oppositeIso, position(0), position(1)
    AuditSixView sheet.GetName, oppositeIso, "isometric"
    AddText d, "Ізометрія (" & ScaleText(oppositeIso.ScaleDecimal) & ")", LX(AUXILIARY_CAPTION_X_M), AUXILIARY_CAPTION_Y_M, TEXT_HEIGHT_SMALL_M, 0
    AddTechnicalNotes d, doc, cfg, body, qty, False
    SetPaperScale sheet, viewScaleValue
    RefreshTemplateFields d
    AsDrawing(d).ActivateSheet name
End Sub

Private Sub NormalizeSixViewScales(ByVal d As SldWorks.ModelDoc2)
    Dim sheetName As Variant, sheet As SldWorks.Sheet, properties As Variant, viewScaleValue As Double
    Dim view As SldWorks.View, index As Long, x As Double, y As Double
    For Each sheetName In AsDrawing(d).GetSheetNames
        AsDrawing(d).ActivateSheet CStr(sheetName)
        Set sheet = AsDrawing(d).GetCurrentSheet
        sheet.SheetFormatVisible = True
        If viewScaleValue = 0 Then
            properties = sheet.GetProperties
            viewScaleValue = properties(SHEET_SCALE_NUMERATOR_PROPERTY) / properties(SHEET_SCALE_DENOMINATOR_PROPERTY)
        End If
        properties = sheet.GetProperties
        sheetWidth = CDbl(properties(SHEET_WIDTH_PROPERTY))
        sheetHeight = CDbl(properties(SHEET_HEIGHT_PROPERTY))
        layoutSx = (sheetWidth - LAYOUT_HORIZONTAL_RESERVED_M) / REFERENCE_USABLE_WIDTH_M
        layoutSy = (sheetHeight - LAYOUT_VERTICAL_RESERVED_M) / REFERENCE_USABLE_HEIGHT_M
        SetPaperScale sheet, viewScaleValue
        index = 0
        Set view = AsDrawing(d).GetFirstView.GetNextView
        Do While Not view Is Nothing
            index = index + 1
            x = LX(ORTHOGRAPHIC_COLUMN_X_M)
            y = LY(ORTHOGRAPHIC_ROW_Y_M)
            If index Mod 2 = 0 Then
                y = LY(AUXILIARY_ROW_Y_M)
            End If
            If index > 2 Then
                x = LX(AUXILIARY_COLUMN_X_M)
            End If
            If index < SIX_VIEW_ISOMETRIC_INDEX Then
                view.UseSheetScale = SW_USE_SHEET_SCALE
            End If
            d.ForceRebuild3 False
            CenterView view, x, y
            Set view = view.GetNextView
        Loop
    Next sheetName
End Sub

Private Sub AuditFinalSixViews(ByVal d As SldWorks.ModelDoc2)
    Dim sheetName As Variant, view As SldWorks.View, index As Long, role As String
    For Each sheetName In AsDrawing(d).GetSheetNames
        AsDrawing(d).ActivateSheet CStr(sheetName)
        index = 0
        Set view = AsDrawing(d).GetFirstView.GetNextView
        Do While Not view Is Nothing
            index = index + 1
            role = "orthographic"
            If index = SIX_VIEW_ISOMETRIC_INDEX Then
                role = "isometric"
            End If
            AuditSixView CStr(sheetName), view, role
            Set view = view.GetNextView
        Loop
    Next sheetName
End Sub
