Attribute VB_Name = "memory_geometry"
Option Explicit

Private Function MemoryBoxPolygon(ByVal box As Variant) As Variant
    MemoryBoxPolygon = Array(Array(box(0), box(1)), Array(box(2), box(1)), Array(box(2), box(3)), Array(box(0), box(3)))
End Function

' Move only free linear text along its measured axis. Witness feet and arrow
' tips stay attached to the same references. This supplies both text sides
' when the native creation order initially put a short size outside the paper.
Private Function MemoryLinearTextCandidate(ByVal seed As Object, ByVal side As Long) As Object
    Dim result As Object, direction As Variant, first As Variant, last As Variant, baseline As Variant
    Dim low As Double, high As Double, textLow As Double, textHigh As Double, projected As Double, shift As Double
    Dim point As Variant, polygon As Variant, moved() As Variant, index As Long, key As Variant, line As Variant
    Dim text As New Collection, lines As New Collection, bases As New Collection, nx As Double, ny As Double, normal As Double
    direction = seed("direction")
    first = seed("first"): last = seed("last"): baseline = seed("baseline")
    low = MinD(first(0) * direction(0) + first(1) * direction(1), last(0) * direction(0) + last(1) * direction(1))
    high = MaxD(first(0) * direction(0) + first(1) * direction(1), last(0) * direction(0) + last(1) * direction(1))
    textLow = SEARCH_INITIAL_LIMIT: textHigh = -SEARCH_INITIAL_LIMIT
    For Each polygon In seed("text")
        For Each point In polygon
            projected = point(0) * direction(0) + point(1) * direction(1)
            textLow = MinD(textLow, projected): textHigh = MaxD(textHigh, projected)
        Next point
    Next polygon
    shift = low - GOST_ARROW_LENGTH_M - (textHigh - textLow) / 2
    If side > 0 Then
        shift = high + GOST_ARROW_LENGTH_M + (textHigh - textLow) / 2
    End If
    shift = shift - (textLow + textHigh) / 2
    Set result = CreateObject("Scripting.Dictionary")
    For Each key In seed.Keys
        result.Add key, seed(key)
    Next key
    For Each polygon In seed("text")
        ReDim moved(UBound(polygon))
        For index = 0 To UBound(polygon)
            moved(index) = PlanarShiftPoint(polygon(index), direction(0) * shift, direction(1) * shift)
        Next index
        text.Add moved
    Next polygon
    nx = -direction(1): ny = direction(0)
    normal = baseline(0) * nx + baseline(1) * ny
    For Each line In seed("lines")
        If Abs(line(0) * nx + line(1) * ny - normal) > PROFILE_MATERIAL_TOLERANCE_M Or Abs(line(2) * nx + line(3) * ny - normal) > PROFILE_MATERIAL_TOLERANCE_M Then
            lines.Add line
        End If
    Next line
    low = MinD(low - GOST_ARROW_LENGTH_M, textLow + shift - MEMORY_FREE_TEXT_TAIL_CLEARANCE_M)
    high = MaxD(high + GOST_ARROW_LENGTH_M, textHigh + shift + MEMORY_FREE_TEXT_TAIL_CLEARANCE_M)
    line = Array(direction(0) * low + nx * normal, direction(1) * low + ny * normal, direction(0) * high + nx * normal, direction(1) * high + ny * normal)
    lines.Add line
    bases.Add line
    Set result("text") = text
    Set result("lines") = lines
    Set result("bases") = bases
    result("position") = PlanarShiftPoint(seed("position"), direction(0) * shift, direction(1) * shift)
    result("controlPoint") = result("position")
    result("centerText") = False
    result("offsetText") = False
    result("cost") = CDbl(seed("cost")) - PlanarGraphicLength(seed("lines")) + PlanarGraphicLength(lines)
    MemoryRebound result
    Set MemoryLinearTextCandidate = result
End Function

' When the baseline crosses its references, the witness overhang must reverse.
' A rigid translation would leave it pointing towards the part and underestimate
' the graphics by twice the native overhang length.
Private Function MemoryTranslateLinearCandidate(ByVal seed As Object, ByVal nx As Double, ByVal ny As Double, ByVal delta As Double) As Object
    Dim result As Object, direction As Variant, baseline As Variant, line As Variant, moved As Variant
    Dim lines As New Collection, first As Variant, last As Variant, foot As Variant, point As Variant
    Dim index As Long, endpoint As Long, along As Double, oldSide As Double, newSide As Double, overhang As Double
    direction = seed("direction"): baseline = seed("baseline")
    first = seed("first"): last = seed("last")
    Set result = PlanarTranslateCandidate(seed, first, last, direction(0), direction(1), nx, ny, delta)
    For index = 1 To seed("lines").Count
        line = seed("lines")(index)
        moved = result("lines")(index)
        along = (line(2) - line(0)) * direction(0) + (line(3) - line(1)) * direction(1)
        If Abs(along) <= PROFILE_MATERIAL_TOLERANCE_M Then
            For Each foot In Array(first, last)
                If Abs((line(0) - foot(0)) * direction(0) + (line(1) - foot(1)) * direction(1)) <= PROFILE_MATERIAL_TOLERANCE_M Then
                    oldSide = (baseline(0) - foot(0)) * nx + (baseline(1) - foot(1)) * ny
                    newSide = oldSide + delta
                    If oldSide * newSide < 0 Then
                        For endpoint = 0 To 2 Step 2
                            If Not PlanarExtensionFoot(line(endpoint), line(endpoint + 1), first, last, direction(0), direction(1), nx, ny) Then
                                overhang = (line(endpoint) - baseline(0)) * nx + (line(endpoint + 1) - baseline(1)) * ny
                                moved(endpoint) = moved(endpoint) - 2 * overhang * nx
                                moved(endpoint + 1) = moved(endpoint + 1) - 2 * overhang * ny
                            End If
                        Next endpoint
                    End If
                    Exit For
                End If
            Next foot
        End If
        lines.Add moved
    Next index
    result("cost") = CDbl(result("cost")) - PlanarGraphicLength(result("lines")) + PlanarGraphicLength(lines)
    Set result("lines") = lines
    result("baseline") = Array(baseline(0) + nx * delta, baseline(1) + ny * delta)
    MemoryRebound result
    Set MemoryTranslateLinearCandidate = result
End Function

Private Sub MemoryRebound(ByVal candidate As Object)
    Dim bounds As Variant, line As Variant, polygon As Variant, point As Variant
    bounds = Array(SEARCH_INITIAL_LIMIT, SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT, -SEARCH_INITIAL_LIMIT)
    For Each line In candidate("lines")
        NativeExpandBounds bounds, line(0), line(1)
        NativeExpandBounds bounds, line(2), line(3)
    Next line
    For Each polygon In candidate("text")
        For Each point In polygon
            NativeExpandBounds bounds, point(0), point(1)
        Next point
    Next polygon
    For Each polygon In candidate("heads")
        For Each point In polygon
            NativeExpandBounds bounds, point(0), point(1)
        Next point
    Next polygon
    candidate("bounds") = bounds
    SpatialClear candidate
End Sub

' Witness feet are captured once. No COM objects are used by candidate generation.
Private Function MemoryWitnessFoot(ByVal arrow As Variant, ByVal lines As Collection, ByVal ux As Double, ByVal uy As Double) As Variant
    Dim line As Variant, point As Variant, best As Double, distance As Double, along As Double
    MemoryWitnessFoot = Array(arrow(0), arrow(1))
    For Each line In lines
        If Abs((line(2) - line(0)) * ux + (line(3) - line(1)) * uy) <= PROFILE_MATERIAL_TOLERANCE_M Then
            For Each point In Array(Array(line(0), line(1)), Array(line(2), line(3)))
                along = Abs((point(0) - arrow(0)) * ux + (point(1) - arrow(1)) * uy)
                distance = Sqr((point(0) - arrow(0)) ^ 2 + (point(1) - arrow(1)) ^ 2)
                If along <= PROFILE_MATERIAL_TOLERANCE_M And distance > best Then
                    best = distance
                    MemoryWitnessFoot = point
                End If
            Next point
        End If
    Next line
End Function

Private Function MemoryCandidateFits(ByVal candidate As Object, ByVal region As Variant, ByVal obstacles As Collection, ByVal material As Variant, ByVal outline As Variant) As Boolean
    Dim polygon As Variant, line As Variant, box As Variant, bounds As Variant
    If Not NativeBoundsInside(candidate("bounds"), region) Then
        RejectPlanarCandidate "workspace bounds"
        Exit Function
    End If
    bounds = LayoutInkBounds(candidate("text"))
    For Each box In obstacles
        If NativeBoxesOverlap(bounds, box, LEADER_LABEL_CLEARANCE_M) Then
            RejectPlanarCandidate "occupied text"
            Exit Function
        End If
        For Each line In candidate("lines")
            If NativeLeaderCrossesBox(line(0), line(1), line(2), line(3), box) Then
                RejectPlanarCandidate "occupied graphics"
                Exit Function
            End If
        Next line
    Next box
    If IsEmpty(material) Then
        If NativeBoxesOverlap(bounds, outline, LEADER_LABEL_CLEARANCE_M) Then
            RejectPlanarCandidate "own outline text"
            Exit Function
        End If
        If candidate("type") = SW_DIMENSION_LINEAR Then
            For Each line In candidate("bases")
                If NativeLeaderCrossesBox(line(0), line(1), line(2), line(3), outline) Then
                    RejectPlanarCandidate "own outline baseline"
                    Exit Function
                End If
            Next line
        End If
    Else
        For Each polygon In candidate("text")
            box = LayoutInkBoundsOfPolygon(polygon)
            If SpatialBoundsOverlap(box, outline) Then
                If PanelBoxTouchesMaterial(box, material) Then
                    RejectPlanarCandidate "own material text"
                    Exit Function
                End If
            End If
        Next polygon
        If candidate("type") = SW_DIMENSION_LINEAR Or candidate("type") = SW_DIMENSION_ANGULAR Then
            For Each line In candidate("bases")
                If SpatialBoundsOverlap(SpatialLineBounds(line), outline) Then
                    If ProfileLineCutsMaterial(Array(line(0), line(1)), Array(line(2), line(3)), material) Then
                        RejectPlanarCandidate "own material baseline"
                        Exit Function
                    End If
                End If
            Next line
        End If
    End If
    Dim ownText As New Collection
    For Each polygon In candidate("text")
        ownText.Add PlanarGrowRectangle(polygon, -LAYOUT_TEXT_INK_MARGIN_M - DIMENSION_GEOMETRY_GAP_M)
    Next polygon
    If PlanarLinesHitText(candidate("lines"), ownText) Then
        RejectPlanarCandidate "own text lines"
        Exit Function
    End If
    If LayoutHeadPolygonsHitText(candidate("heads"), candidate("text")) Then
        RejectPlanarCandidate "own text arrows"
        Exit Function
    End If
    MemoryCandidateFits = True
End Function

Private Function LayoutInkBoundsOfPolygon(ByVal polygon As Variant) As Variant
    Dim polygons As New Collection
    polygons.Add polygon
    LayoutInkBoundsOfPolygon = LayoutInkBounds(polygons)
End Function

' A circular-size/hole leader retains its geometry anchor. All free endpoints
' move with the text; the first leader segment and arrow direction are rebuilt.
Private Function MemoryLeaderCandidate(ByVal seed As Object, ByVal x As Double, ByVal y As Double) As Object
    Dim result As Object, key As Variant, polygon As Variant, moved() As Variant, index As Long
    Dim lines As New Collection, heads As New Collection, text As New Collection, bases As New Collection
    Dim anchor As Variant, point As Variant, line As Variant, dx As Double, dy As Double, tip As Variant
    Dim tx As Double, ty As Double, ux As Double, uy As Double, length As Double, edge As Long
    dx = x - seed("position")(0)
    dy = y - seed("position")(1)
    anchor = seed("anchor")
    Set result = CreateObject("Scripting.Dictionary")
    For Each key In seed.Keys
        result.Add key, seed(key)
    Next key
    For Each polygon In seed("text")
        ReDim moved(UBound(polygon))
        For index = 0 To UBound(polygon)
            moved(index) = PlanarShiftPoint(polygon(index), dx, dy)
        Next index
        text.Add moved
    Next polygon
    For Each line In seed("lines")
        point = Array(line(0) + dx, line(1) + dy, line(2) + dx, line(3) + dy)
        If Sqr((line(0) - anchor(0)) ^ 2 + (line(1) - anchor(1)) ^ 2) <= GOST_ARROW_HEAD_LENGTH_M Then
            point(0) = line(0): point(1) = line(1)
            tx = point(2): ty = point(3)
        End If
        If Sqr((line(2) - anchor(0)) ^ 2 + (line(3) - anchor(1)) ^ 2) <= GOST_ARROW_HEAD_LENGTH_M Then
            point(2) = line(2): point(3) = line(3)
            tx = point(0): ty = point(1)
        End If
        lines.Add point
    Next line
    ux = tx - anchor(0): uy = ty - anchor(1)
    length = Sqr(ux * ux + uy * uy)
    If length > PROFILE_MATERIAL_TOLERANCE_M Then
        ux = ux / length: uy = uy / length
        For Each polygon In seed("heads")
            tip = polygon(0)
            heads.Add Array(tip, Array(tip(0) + ux * GOST_ARROW_HEAD_LENGTH_M - uy * GOST_ARROW_HEAD_WIDTH_M / 2, tip(1) + uy * GOST_ARROW_HEAD_LENGTH_M + ux * GOST_ARROW_HEAD_WIDTH_M / 2), Array(tip(0) + ux * GOST_ARROW_HEAD_LENGTH_M + uy * GOST_ARROW_HEAD_WIDTH_M / 2, tip(1) + uy * GOST_ARROW_HEAD_LENGTH_M - ux * GOST_ARROW_HEAD_WIDTH_M / 2))
        Next polygon
    End If
    Set result("text") = text
    Set result("lines") = lines
    Set result("heads") = heads
    Set result("bases") = bases
    result("position") = Array(x, y, 0)
    result("controlPoint") = result("position")
    result("cost") = PlanarGraphicLength(lines)
    MemoryRebound result
    Set MemoryLeaderCandidate = result
End Function

' Change the radius of an angular annotation, preserving the measured angle,
' witness feet, printed text height and arrowhead size.
Private Function MemoryAngularCandidate(ByVal seed As Object, ByVal radius As Double) As Object
    Dim result As Object, key As Variant, center As Variant, position As Variant, ratio As Double
    Dim lines As New Collection, bases As New Collection, text As New Collection, heads As New Collection
    Dim line As Variant, polygon As Variant, moved() As Variant, index As Long, a As Variant, b As Variant, tip As Variant
    Dim dx As Double, dy As Double
    center = seed("arcCenter")
    ratio = radius / seed("arcRadius")
    position = seed("position")
    dx = (position(0) - center(0)) * (ratio - 1)
    dy = (position(1) - center(1)) * (ratio - 1)
    Set result = CreateObject("Scripting.Dictionary")
    For Each key In seed.Keys
        result.Add key, seed(key)
    Next key
    For Each polygon In seed("text")
        ReDim moved(UBound(polygon))
        For index = 0 To UBound(polygon)
            moved(index) = PlanarShiftPoint(polygon(index), dx, dy)
        Next index
        text.Add moved
    Next polygon
    For Each polygon In seed("heads")
        tip = polygon(0)
        ReDim moved(UBound(polygon))
        For index = 0 To UBound(polygon)
            moved(index) = PlanarShiftPoint(polygon(index), (tip(0) - center(0)) * (ratio - 1), (tip(1) - center(1)) * (ratio - 1))
        Next index
        heads.Add moved
    Next polygon
    For Each line In seed("bases")
        bases.Add Array(center(0) + (line(0) - center(0)) * ratio, center(1) + (line(1) - center(1)) * ratio, center(0) + (line(2) - center(0)) * ratio, center(1) + (line(3) - center(1)) * ratio)
    Next line
    For Each line In seed("lines")
        a = MemoryAngularEndpoint(Array(line(0), line(1)), seed, ratio)
        b = MemoryAngularEndpoint(Array(line(2), line(3)), seed, ratio)
        lines.Add Array(a(0), a(1), b(0), b(1))
    Next line
    Set result("text") = text
    Set result("lines") = lines
    Set result("heads") = heads
    Set result("bases") = bases
    result("position") = PlanarShiftPoint(position, dx, dy)
    result("controlPoint") = result("position")
    result("arcRadius") = radius
    result("cost") = PlanarGraphicLength(lines)
    MemoryRebound result
    Set MemoryAngularCandidate = result
End Function

' Vertically opposite angular sectors have the same measured value. Predict
' both in memory; toggle the native sector only when applying the chosen one.
Private Function MemoryOppositeAngularCandidate(ByVal seed As Object) As Object
    Dim result As Object, key As Variant, polygon As Variant, moved() As Variant, index As Long
    Dim geometry As Collection, point As Variant, center As Variant, segment As Variant, a As Variant, b As Variant
    Dim radius As Double, distance As Double
    Set result = CreateObject("Scripting.Dictionary")
    center = seed("arcCenter")
    radius = seed("arcRadius")
    For Each key In seed.Keys
        Select Case CStr(key)
            Case "text", "heads"
                Set geometry = New Collection
                For Each polygon In seed(key)
                    ReDim moved(UBound(polygon))
                    For index = 0 To UBound(polygon)
                        point = polygon(index)
                        moved(index) = Array(2 * center(0) - point(0), 2 * center(1) - point(1))
                    Next index
                    geometry.Add moved
                Next polygon
                result.Add key, geometry
            Case "bases", "lines"
                Set geometry = New Collection
                For Each segment In seed(key)
                    a = Array(segment(0), segment(1)): b = Array(segment(2), segment(3))
                    For index = 0 To 1
                        point = a
                        If index = 1 Then
                            point = b
                        End If
                        distance = Sqr((point(0) - center(0)) ^ 2 + (point(1) - center(1)) ^ 2)
                        If key = "bases" Or Abs(distance - radius) <= GOST_WITNESS_EXTENSION_M + PROFILE_MATERIAL_TOLERANCE_M Then
                            point = Array(2 * center(0) - point(0), 2 * center(1) - point(1))
                        End If
                        If index = 0 Then
                            a = point
                        Else
                            b = point
                        End If
                    Next index
                    geometry.Add Array(a(0), a(1), b(0), b(1))
                Next segment
                result.Add key, geometry
            Case "position", "controlPoint"
                point = seed(key)
                result.Add key, Array(2 * center(0) - point(0), 2 * center(1) - point(1), point(2))
            Case Else
                result.Add key, seed(key)
        End Select
    Next key
    result("sectorState") = 1 - CLng(seed("sectorState"))
    MemoryRebound result
    Set MemoryOppositeAngularCandidate = result
End Function

Private Function MemoryAngularEndpoint(ByVal point As Variant, ByVal seed As Object, ByVal ratio As Double) As Variant
    Dim center As Variant, radius As Double
    center = seed("arcCenter")
    radius = Sqr((point(0) - center(0)) ^ 2 + (point(1) - center(1)) ^ 2)
    MemoryAngularEndpoint = point
    If Abs(radius - CDbl(seed("arcRadius"))) <= GOST_WITNESS_EXTENSION_M + PROFILE_MATERIAL_TOLERANCE_M Then
        MemoryAngularEndpoint = Array(center(0) + (point(0) - center(0)) * ratio, center(1) + (point(1) - center(1)) * ratio)
    End If
End Function
