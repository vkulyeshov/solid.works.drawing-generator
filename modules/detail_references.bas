Attribute VB_Name = "detail_references"
Option Explicit

Private Function DetailReferenceLabel(ByVal number As Long) As String
    Dim letters As Variant, count As Long
    letters = Split(PLANAR_DETAIL_LABELS, "|")
    count = UBound(letters) + 1
    Do While number > 0
        number = number - 1
        DetailReferenceLabel = letters(number Mod count) & DetailReferenceLabel
        number = number \ count
    Loop
End Function

' Cropped profile views are not native DetailView objects, so explicitly mark
' their source areas and use the same Ukrainian letter in their captions.
Private Sub AddProfileDetailReferences(ByVal d As SldWorks.ModelDoc2)
    Dim group As Variant, request As Variant, first As Variant, last As Variant, bounds As Variant
    Dim overview As SldWorks.View, profile As SldWorks.SketchSegment, index As Long, label As String
    Dim originalSheet As String, cx As Double, cy As Double, radius As Double, cell As Variant, rx As Double, ry As Double
    Dim marker As Object, captions As Collection, workspace As Object, frame As Variant
    Dim position As Variant, outline As Variant, candidates As Collection, item As Variant, free As Boolean
    Dim direction As Long, angle As Double, x As Double, y As Double, box As Variant
    Dim sourceViews As New Collection, sourceView As SldWorks.View, material As Variant, distance As Double
    If profilePageGroups Is Nothing Then
        Exit Sub
    End If
    originalSheet = AsDrawing(d).GetCurrentSheet.GetName
    AsDrawing(d).ActivateSheet profilePageFirstSheet
    Set overview = AsDrawing(d).GetFirstView.GetNextView
    Set sourceView = overview
    Do While Not sourceView Is Nothing
        sourceViews.Add sourceView
        Set sourceView = sourceView.GetNextView
    Loop
    CacheProfileContours sourceViews
    If profileSheetContours.Exists(overview.Name) Then
        material = profileSheetContours(overview.Name)
    End If
    Set workspace = SheetWorkspace(d)
    frame = workspace("frame")
    Set captions = workspace("obstacles")
    For Each group In profilePageGroups
        If group.Exists("view") And group.Exists("sheet") Then
            index = index + 1
            label = DetailReferenceLabel(index)
            If Not group.Exists("referenceLabel") Then
                group.Add "referenceLabel", label
                bounds = Array(SEARCH_INITIAL_LIMIT, SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT)
                AsDrawing(d).ActivateSheet profilePageFirstSheet
                For Each request In group("requests")
                    first = PatternSheetPoint(overview, request("first"))
                    last = PatternSheetPoint(overview, request("last"))
                    NativeExpandBounds bounds, first(0), first(1)
                    NativeExpandBounds bounds, last(0), last(1)
                Next request
                cx = (bounds(0) + bounds(2)) / 2
                cy = (bounds(1) + bounds(3)) / 2
                bounds = Array(MaxD(frame(0), bounds(0) - MEMORY_REFERENCE_PADDING_M), MaxD(frame(1), bounds(1) - MEMORY_REFERENCE_PADDING_M), MinD(frame(2), bounds(2) + MEMORY_REFERENCE_PADDING_M), MinD(frame(3), bounds(3) + MEMORY_REFERENCE_PADDING_M))
                bounds = ReferenceOvalBounds(bounds, frame)
                cx = (bounds(0) + bounds(2)) / 2
                cy = (bounds(1) + bounds(3)) / 2
                rx = (bounds(2) - bounds(0)) / 2
                ry = (bounds(3) - bounds(1)) / 2
                CreateProfileReferenceBoundary d, overview, bounds
                AsDrawing(d).ActivateView ""
                free = False
                For distance = MEMORY_REFERENCE_CAPTION_GAP_M To PLANAR_DETAIL_LABEL_REACH_M Step TEXT_HEIGHT_SMALL_M
                  For direction = 0 To MEMORY_LAYOUT_DIRECTIONS - 1
                    angle = direction * FULL_REVOLUTION_RAD / MEMORY_LAYOUT_DIRECTIONS
                    x = cx + (rx + distance) * Cos(angle)
                    y = cy + (ry + distance) * Sin(angle)
                    box = Array(x - TEXT_HEIGHT_STANDARD_M, y - TEXT_HEIGHT_STANDARD_M, x + TEXT_HEIGHT_STANDARD_M, y + TEXT_HEIGHT_STANDARD_M)
                    free = NativeBoundsInside(box, frame)
                    If NativeBoxesOverlap(box, bounds, LEADER_LABEL_CLEARANCE_M) Then
                        free = False
                    End If
                    If Not IsEmpty(material) Then
                        If PanelBoxTouchesMaterial(box, material) Then
                            free = False
                        End If
                    End If
                    For Each item In captions
                        If NativeBoxesOverlap(box, item, LEADER_LABEL_CLEARANCE_M) Then
                            free = False
                        End If
                    Next item
                    If free Then
                        Exit For
                    End If
                  Next direction
                  If free Then
                      Exit For
                  End If
                Next distance
                If free = False Then
                    ' An initial greedy letter placement is not a layout
                    ' decision. The joint solver can move earlier letters too.
                    x = MinD(frame(2) - TEXT_HEIGHT_STANDARD_M, MaxD(frame(0) + TEXT_HEIGHT_STANDARD_M, cx + rx + MEMORY_REFERENCE_CAPTION_GAP_M))
                    y = MinD(frame(3) - TEXT_HEIGHT_STANDARD_M, MaxD(frame(1) + TEXT_HEIGHT_STANDARD_M, cy))
                    box = Array(x - TEXT_HEIGHT_STANDARD_M, y - TEXT_HEIGHT_STANDARD_M, x + TEXT_HEIGHT_STANDARD_M, y + TEXT_HEIGHT_STANDARD_M)
                    Log "DETAIL REFERENCE joint placement required for " & label
                End If
                RegisterMemoryReference d, label, x, y, overview.Name, bounds
                captions.Add box
                AsDrawing(d).ActivateSheet CStr(group("sheet"))
                If group.Exists("cell") Then
                    cell = group("cell")
                Else
                    cell = group("region")
                End If
                AddText d, label & " (" & ScaleText(group("scale")) & ")", cell(0), cell(3) + PROFILE_PAGE_CAPTION_GAP_M, TEXT_HEIGHT_SMALL_M, 0
                Log "DETAIL REFERENCE " & label & "; overview=" & profilePageFirstSheet & "; detail=" & group("sheet") & "/" & group("view")
            End If
        End If
    Next group
    AsDrawing(d).ActivateSheet originalSheet
    d.GraphicsRedraw2
End Sub

' Enclose a feature family with a circle or an oval, using a continuous thin line.
Private Sub CreateProfileReferenceBoundary(ByVal d As SldWorks.ModelDoc2, ByVal view As SldWorks.View, ByVal bounds As Variant)
    Dim sketch As SldWorks.Sketch, util As SldWorks.MathUtility, point As SldWorks.MathPoint
    Dim transform As SldWorks.MathTransform, corners As Variant, converted(2) As Variant, a As Variant, b As Variant, c As Variant
    Dim target(2) As Double, index As Long, oldFocus As Boolean, segment As SldWorks.SketchSegment
    Set sketch = view.GetSketch
    Set transform = sketch.ModelToSketchTransform
    Set util = sw.GetMathUtility
    a = Array((bounds(0) + bounds(2)) / 2, (bounds(1) + bounds(3)) / 2)
    If bounds(2) - bounds(0) >= bounds(3) - bounds(1) Then
        corners = Array(a, Array(bounds(2), a(1)), Array(a(0), bounds(3)))
    Else
        corners = Array(a, Array(a(0), bounds(3)), Array(bounds(2), a(1)))
    End If
    For index = 0 To 2
        a = corners(index)
        target(0) = a(0): target(1) = a(1): target(2) = 0
        Set point = util.CreatePoint(target)
        Set point = point.MultiplyTransform(transform)
        converted(index) = point.ArrayData
    Next index
    AsDrawing(d).ActivateView view.Name
    oldFocus = view.FocusLocked
    view.FocusLocked = True
    a = converted(0): b = converted(1): c = converted(2)
    If Abs((bounds(2) - bounds(0)) - (bounds(3) - bounds(1))) <= PROFILE_MATERIAL_TOLERANCE_M Then
        Set segment = d.SketchManager.CreateCircleByRadius(a(0), a(1), a(2), Sqr((b(0) - a(0)) ^ 2 + (b(1) - a(1)) ^ 2))
    Else
        Set segment = d.SketchManager.CreateEllipse(a(0), a(1), a(2), b(0), b(1), b(2), c(0), c(1), c(2))
    End If
    If segment Is Nothing Then
        view.FocusLocked = oldFocus
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Cannot draw a source-area reference oval."
    End If
    StyleReferenceSegment segment
    view.FocusLocked = oldFocus
End Sub

Private Function ReferenceOvalBounds(ByVal bounds As Variant, ByVal frame As Variant) As Variant
    Dim cx As Double, cy As Double, rx As Double, ry As Double
    cx = (bounds(0) + bounds(2)) / 2: cy = (bounds(1) + bounds(3)) / 2
    rx = MaxD(MEMORY_REFERENCE_MIN_RADIUS_M, (bounds(2) - bounds(0)) * MEMORY_REFERENCE_ENCLOSURE_FACTOR / 2)
    ry = MaxD(MEMORY_REFERENCE_MIN_RADIUS_M, (bounds(3) - bounds(1)) * MEMORY_REFERENCE_ENCLOSURE_FACTOR / 2)
    If MaxD(rx, ry) / MinD(rx, ry) <= MEMORY_REFERENCE_CIRCLE_RATIO Then
        rx = MaxD(rx, ry): ry = rx
    End If
    ReferenceOvalBounds = Array(MaxD(frame(0), cx - rx), MaxD(frame(1), cy - ry), MinD(frame(2), cx + rx), MinD(frame(3), cy + ry))
End Function

Private Sub StyleReferenceSegment(ByVal segment As SldWorks.SketchSegment)
    segment.Color = DIMENSION_COLOR_BLACK
    segment.Style = SW_SKETCH_STYLE_CONTINUOUS
    segment.Width = SW_LINE_WEIGHT_DIMENSION
    segment.LayerOverride = segment.LayerOverride Or SW_LAYER_OVERRIDE_COLOR Or SW_LAYER_OVERRIDE_WIDTH
End Sub
