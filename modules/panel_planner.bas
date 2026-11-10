Attribute VB_Name = "panel_planner"
Option Explicit

Private Function PanelNode(ByVal points As Collection, ByVal index As Long) As Object
    Set PanelNode = points((index - 1 + points.Count * PANEL_INDEX_WRAP_FACTOR) Mod points.Count + 1)
End Function

Private Function PanelDistance(ByVal a As Object, ByVal b As Object) As Double
    PanelDistance = Sqr((a("x") - b("x")) ^ 2 + (a("y") - b("y")) ^ 2)
End Function

Private Sub PanelCleanFrame(ByVal points As Collection)
    Dim index As Long, a As Object, b As Object, c As Object, length As Double, changed As Boolean
    Do
        changed = False
        For index = points.Count To 1 Step -1
            Set a = PanelNode(points, index - 1)
            Set b = PanelNode(points, index)
            Set c = PanelNode(points, index + 1)
            length = PanelDistance(a, c)
            If length > PROFILE_COORDINATE_TOLERANCE_M Then
                If Abs((b("x") - a("x")) * (c("y") - a("y")) - (b("y") - a("y")) * (c("x") - a("x"))) < PROFILE_COORDINATE_TOLERANCE_M * length Then
                    If (b("x") - a("x")) * (c("x") - b("x")) + (b("y") - a("y")) * (c("y") - b("y")) >= 0 Then
                        points.Remove index
                        changed = True
                        Exit For
                    End If
                End If
            End If
        Next index
    Loop While changed And points.Count > PANEL_MIN_FRAME_VERTICES
End Sub

' A shallow boundary excursion is bracketed by the same supporting line.
' Its original vertices remain available for dimensions and coverage auditing.
Private Function PanelExcursion(ByVal points As Collection, ByVal start As Long, ByVal count As Long, ByVal limit As Double) As Object
    Dim a As Object, b As Object, node As Object, before As Object, after As Object, previous As Object
    Dim dx As Double, dy As Double, length As Double, offset As Double, along As Double, depth As Double, sign As Double, area As Double
    Dim index As Long, result As Object, raw As New Collection, depthA As Object, depthB As Object, edgeLength As Double
    Set a = PanelNode(points, start)
    Set b = PanelNode(points, start + count)
    length = PanelDistance(a, b)
    If length <= PROFILE_COORDINATE_TOLERANCE_M Then
        Exit Function
    End If
    dx = (b("x") - a("x")) / length
    dy = (b("y") - a("y")) / length
    Set before = PanelNode(points, start - 1)
    Set after = PanelNode(points, start + count + 1)
    If Abs((before("x") - a("x")) * dy - (before("y") - a("y")) * dx) > PROFILE_COORDINATE_TOLERANCE_M Then
        Exit Function
    End If
    If Abs((after("x") - b("x")) * dy - (after("y") - b("y")) * dx) > PROFILE_COORDINATE_TOLERANCE_M Then
        Exit Function
    End If
    For index = 0 To count
        Set node = PanelNode(points, start + index)
        raw.Add node
        along = (node("x") - a("x")) * dx + (node("y") - a("y")) * dy
        offset = (node("x") - a("x")) * dy - (node("y") - a("y")) * dx
        If along < -PROFILE_COORDINATE_TOLERANCE_M Or along > length + PROFILE_COORDINATE_TOLERANCE_M Then
            Exit Function
        End If
        If Abs(offset) > PROFILE_COORDINATE_TOLERANCE_M Then
            If sign <> 0 And Sgn(offset) <> sign Then
                Exit Function
            End If
            sign = Sgn(offset)
        End If
        If Abs(offset) > depth Then
            depth = Abs(offset)
        End If
    Next index
    If depth <= PROFILE_COORDINATE_TOLERANCE_M Then
        Exit Function
    End If
    If depth > limit And length > limit * FAMILY_NARROW_SLOT_FACTOR Then
        Exit Function
    End If
    For index = 1 To points.Count
        Set node = PanelNode(points, index)
        Set previous = PanelNode(points, index + 1)
        area = area + node("x") * previous("y") - previous("x") * node("y")
    Next index
    If sign * area >= 0 Then
        Exit Function
    End If
    ' Prefer a complete normal edge: equal depths are not inferred from a
    ' projected bounding box or merely from equally rounded labels.
    For index = 1 To raw.Count - 1
        Set previous = raw(index)
        Set node = raw(index + 1)
        edgeLength = PanelDistance(previous, node)
        If Abs((node("x") - previous("x")) * dx + (node("y") - previous("y")) * dy) <= PROFILE_COORDINATE_TOLERANCE_M Then
            If Abs(edgeLength - depth) <= PROFILE_COORDINATE_TOLERANCE_M Then
                Set depthA = previous
                Set depthB = node
                Exit For
            End If
        End If
    Next index
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "normalDepth", Not depthA Is Nothing
    If depthA Is Nothing Then
        Set depthA = raw(1)
        Set depthB = raw(2)
    End If
    result.Add "first", a
    result.Add "last", b
    result.Add "depthFirst", depthA
    result.Add "depthLast", depthB
    result.Add "raw", raw
    result.Add "width", length
    result.Add "depth", depth
    result.Add "count", count
    Set PanelExcursion = result
End Function

Private Function PanelStripExcursions(ByVal source As Collection, ByVal limit As Double, ByVal features As Collection) As Collection
    Dim points As Collection, nextPoints As Collection, feature As Object, node As Object
    Dim index As Long, count As Long, offset As Long, changed As Boolean, remove As Object
    Set points = source
    Do
        changed = False
        For count = PANEL_MIN_EXCURSION_EDGES To PANEL_MAX_EXCURSION_EDGES
            If count < points.Count - PANEL_MIN_FRAME_VERTICES Then
                For index = 1 To points.Count
                    Set feature = PanelExcursion(points, index, count, limit)
                    If Not feature Is Nothing Then
                        features.Add feature
                        Set remove = CreateObject("Scripting.Dictionary")
                        For offset = 1 To count - 1
                            remove((index + offset - 1) Mod points.Count + 1) = True
                        Next offset
                        Set nextPoints = New Collection
                        For offset = 1 To points.Count
                            If Not remove.Exists(offset) Then
                                nextPoints.Add points(offset)
                            End If
                        Next offset
                        Set points = nextPoints
                        PanelCleanFrame points
                        changed = True
                        Exit For
                    End If
                Next index
            End If
            If changed Then
                Exit For
            End If
        Next count
    Loop While changed
    Set PanelStripExcursions = points
End Function

Private Sub PanelAddRequest(ByVal requests As Collection, ByVal groups As Object, ByVal a As Object, ByVal b As Object, ByVal axis As Long, ByVal role As String, ByVal kind As String, ByVal feature As String, ByVal merge As Boolean, Optional ByVal overall As Boolean = False)
    Dim request As Object, value As Double, key As String, coordinate As String, features As Object
    value = PanelDistance(a, b)
    If axis = SCREEN_X_AXIS_INDEX Then
        value = Abs(a("x") - b("x"))
    ElseIf axis = SCREEN_Y_AXIS_INDEX Then
        value = Abs(a("y") - b("y"))
    End If
    If value <= PROFILE_COORDINATE_TOLERANCE_M Then
        Exit Sub
    End If
    key = kind & "|" & CStr(axis) & "|" & ProfileCoordinateKey(value)
    If kind = "notch-depth" Then
        key = kind & "|" & ProfileCoordinateKey(value)
    End If
    If Not merge Then
        key = key & "|" & feature
    End If
    If groups.Exists(key) Then
        Set request = groups(key)
        request("features")(feature) = True
        Exit Sub
    End If
    Set request = CreateObject("Scripting.Dictionary")
    request.Add "key", "panel|" & key
    request.Add "low", ProfileCoordinateKey(0)
    request.Add "high", ProfileCoordinateKey(value)
    request.Add "axis", axis
    request.Add "role", role
    request.Add "overall", overall
    request.Add "value", value
    request.Add "first", a("point")
    request.Add "last", b("point")
    request.Add "projectedFirst", Array(a("x"), a("y"))
    request.Add "projectedLast", Array(b("x"), b("y"))
    request.Add "panel", True
    request.Add "aligned", axis = PANEL_ALIGNED_AXIS
    request.Add "kind", kind
    Set features = CreateObject("Scripting.Dictionary")
    features(feature) = True
    request.Add "features", features
    groups.Add key, request
    requests.Add request
End Sub

' Recognize local recesses before planning the remaining structural frame.
' Curved boundaries and non-rectangular openings retain their existing planner.
Private Function PlanPanelDimensions(ByVal outer As Collection, ByVal openings As Collection) As Collection
    Dim lowX As Object, highX As Object, lowY As Object, highY As Object, a As Object, b As Object, c As Object, last As Object
    Dim features As New Collection, frame As Collection, request As Object, feature As Object, raw As Collection
    Dim result As New Collection, groups As Object, index As Long, count As Long, axis As Long, depthAxis As Long
    Dim area As Double, span As Double, id As String, key As String, diagonalCount As Long, cutCount As Long
    Dim start As Object, finish As Object, candidate As Object, dx As Double, dy As Double, distance As Double, best As Double, neighbor As Object, datum As Object
    Dim datumDistance As Double, perimeterEdge As Boolean
    For Each raw In openings
        If raw.Count <> RECTANGLE_EDGE_COUNT Then
            Exit Function
        End If
        For index = 1 To raw.Count
            If ProfileSegmentAxis(PanelNode(raw, index), PanelNode(raw, index + 1)) < 0 Then
                Exit Function
            End If
        Next index
    Next raw
    Set lowX = outer(1)
    Set highX = lowX
    Set lowY = lowX
    Set highY = lowX
    For Each a In outer
        If a("x") < lowX("x") Then
            Set lowX = a
        End If
        If a("x") > highX("x") Then
            Set highX = a
        End If
        If a("y") < lowY("y") Then
            Set lowY = a
        End If
        If a("y") > highY("y") Then
            Set highY = a
        End If
    Next a
    span = MinD(highX("x") - lowX("x"), highY("y") - lowY("y"))
    Set frame = PanelStripExcursions(outer, span * PANEL_MINOR_DEPTH_RATIO, features)
    ' A shallow trim that reaches an outside corner has no second mouth on
    ' the supporting line. Preserve it as a feature instead of losing its
    ' length when the closed-notch recognizer reaches the perimeter end.
    For index = 1 To outer.Count
        Set a = PanelNode(outer, index)
        Set b = PanelNode(outer, index + 1)
        Set c = PanelNode(outer, index + 2)
        Set last = PanelNode(outer, index + 3)
        axis = ProfileSegmentAxis(a, b)
        Set start = Nothing
        Set finish = Nothing
        If axis >= 0 Then
            If ProfileSegmentAxis(c, last) = axis And ProfileSegmentAxis(b, c) = 1 - axis Then
                If PanelDistance(b, c) <= span * PANEL_MINOR_DEPTH_RATIO Then
                    Set candidate = lowX
                    Set neighbor = highX
                    key = "x"
                    If axis = SCREEN_X_AXIS_INDEX Then
                        Set candidate = lowY
                        Set neighbor = highY
                        key = "y"
                    End If
                    If Abs(c(key) - candidate(key)) < PROFILE_COORDINATE_TOLERANCE_M Or Abs(c(key) - neighbor(key)) < PROFILE_COORDINATE_TOLERANCE_M Then
                        If ProfileSegmentAxis(PanelNode(outer, index - 1), a) = 1 - axis And PanelDistance(PanelNode(outer, index - 1), a) > span * PANEL_MINOR_DEPTH_RATIO Then
                            Set start = a
                            Set finish = b
                        End If
                    End If
                    If Abs(a(key) - candidate(key)) < PROFILE_COORDINATE_TOLERANCE_M Or Abs(a(key) - neighbor(key)) < PROFILE_COORDINATE_TOLERANCE_M Then
                        If ProfileSegmentAxis(last, PanelNode(outer, index + 4)) = 1 - axis And PanelDistance(last, PanelNode(outer, index + 4)) > span * PANEL_MINOR_DEPTH_RATIO Then
                            Set start = c
                            Set finish = last
                        End If
                    End If
                    If Not start Is Nothing Then
                        If PanelDistance(start, finish) > span * PANEL_MINOR_DEPTH_RATIO Then
                            Set feature = CreateObject("Scripting.Dictionary")
                            feature.Add "first", start
                            feature.Add "last", finish
                            feature.Add "depthFirst", b
                            feature.Add "depthLast", c
                            feature.Add "normalDepth", True
                            feature.Add "width", PanelDistance(start, finish)
                            feature.Add "depth", PanelDistance(b, c)
                            feature.Add "count", PANEL_MIN_EXCURSION_EDGES
                            Set raw = New Collection
                            raw.Add a
                            raw.Add b
                            raw.Add c
                            raw.Add last
                            feature.Add "raw", raw
                            features.Add feature
                        End If
                    End If
                End If
            End If
        End If
    Next index
    For index = 1 To frame.Count
        Set a = PanelNode(frame, index)
        Set b = PanelNode(frame, index + 1)
        area = area + a("x") * b("y") - b("x") * a("y")
    Next index
    Set groups = CreateObject("Scripting.Dictionary")
    PanelAddRequest result, groups, lowX, highX, SCREEN_X_AXIS_INDEX, PROFILE_ROLE_OUTER, "overall-width", "frame", False, True
    PanelAddRequest result, groups, lowY, highY, SCREEN_Y_AXIS_INDEX, PROFILE_ROLE_OUTER, "overall-height", "frame", False, True
    For index = 1 To frame.Count
        Set a = PanelNode(frame, index)
        Set b = PanelNode(frame, index + 1)
        Set c = PanelNode(frame, index + 2)
        Set last = PanelNode(frame, index + 3)
        If ProfileNotch(a, b, c, last, area) And PanelDistance(a, b) > span * PANEL_MINOR_DEPTH_RATIO Then
            cutCount = cutCount + 1
            id = "cut" & CStr(cutCount)
            axis = ProfileSegmentAxis(b, c)
            PanelAddRequest result, groups, b, c, axis, PROFILE_ROLE_OUTER, "major-cut-width", id, True
            PanelAddRequest result, groups, a, b, 1 - axis, PROFILE_ROLE_OUTER, "major-cut-depth", id, True
            If axis = SCREEN_X_AXIS_INDEX Then
                Set candidate = lowX
                Set start = a
                If (a("x") + last("x")) / 2 > (lowX("x") + highX("x")) / 2 Then
                    Set candidate = highX
                    Set start = last
                End If
                PanelAddRequest result, groups, candidate, start, axis, PROFILE_ROLE_OUTER, "major-cut-location", id, True
            End If
        End If
        If ProfileSegmentAxis(a, b) < 0 And PanelDistance(a, b) > span * PANEL_MINOR_DEPTH_RATIO Then
            diagonalCount = diagonalCount + 1
            id = "corner" & CStr(diagonalCount)
            PanelAddRequest result, groups, a, b, SCREEN_X_AXIS_INDEX, PROFILE_ROLE_OUTER, "corner-width", id, True
            PanelAddRequest result, groups, a, b, SCREEN_Y_AXIS_INDEX, PROFILE_ROLE_OUTER, "corner-height", id, True
        End If
    Next index
    If diagonalCount > PANEL_MAX_CLIPPED_CORNERS Then
        Exit Function
    End If
    If cutCount = 0 And diagonalCount > 0 And features.Count = 0 Then
        Exit Function
    End If
    If cutCount = 0 And diagonalCount = 0 Then
        PanelFrameDimensions result, groups, frame
    End If
    index = 0
    For Each feature In features
        index = index + 1
        id = "notch" & CStr(index)
        Set a = feature("first")
        Set b = feature("last")
        axis = ProfileSegmentAxis(a, b)
        If axis < 0 Then
            axis = PANEL_ALIGNED_AXIS
        End If
        PanelAddRequest result, groups, a, b, axis, PROFILE_ROLE_SLOT, "notch-width", id, True
        Set start = feature("depthFirst")
        Set finish = feature("depthLast")
        depthAxis = ProfileSegmentAxis(start, finish)
        If depthAxis < 0 Then
            depthAxis = PANEL_ALIGNED_AXIS
        End If
        key = "notch-depth"
        If CBool(feature("normalDepth")) = False Then
            key = "notch-oblique-edge"
        End If
        PanelAddRequest result, groups, start, finish, depthAxis, PROFILE_ROLE_SLOT, key, id, True
        ' Locate each mouth from an actual vertex on its own supporting line.
        ' Only native references are used; no invented dimension endpoints.
        dx = (b("x") - a("x")) / feature("width")
        dy = (b("y") - a("y")) / feature("width")
        best = SEARCH_INITIAL_LIMIT
        Set candidate = Nothing
        For Each start In frame
            If Abs((start("x") - a("x")) * dy - (start("y") - a("y")) * dx) < PROFILE_COORDINATE_TOLERANCE_M Then
                distance = (a("x") - start("x")) * dx + (a("y") - start("y")) * dy
                If distance > PROFILE_COORDINATE_TOLERANCE_M And distance < best Then
                    best = distance
                    Set candidate = start
                End If
            End If
        Next start
        Set datum = candidate
        datumDistance = best
        Set candidate = Nothing
        best = SEARCH_INITIAL_LIMIT
        For Each neighbor In features
            Set start = neighbor("last")
            If Abs((start("x") - a("x")) * dy - (start("y") - a("y")) * dx) < PROFILE_COORDINATE_TOLERANCE_M Then
                distance = (a("x") - start("x")) * dx + (a("y") - start("y")) * dy
                If distance > PROFILE_COORDINATE_TOLERANCE_M And distance < best Then
                    best = distance
                    Set candidate = start
                End If
            End If
        Next neighbor
        If candidate Is Nothing Then
            Set candidate = datum
        ElseIf Not datum Is Nothing Then
            perimeterEdge = axis = SCREEN_X_AXIS_INDEX And (Abs(a("y") - lowY("y")) < PROFILE_COORDINATE_TOLERANCE_M Or Abs(a("y") - highY("y")) < PROFILE_COORDINATE_TOLERANCE_M)
            If Not perimeterEdge And best > datumDistance Then
                Set candidate = datum
            End If
        End If
        If Not candidate Is Nothing Then
            PanelAddRequest result, groups, candidate, a, axis, PROFILE_ROLE_SLOT, "notch-location", id, False
        End If
    Next feature
    PanelOpeningDimensions result, groups, openings, lowX, lowY
    Log "PANEL PLAN major cuts=" & cutCount & "; clipped corners=" & diagonalCount & "; minor features=" & features.Count & "; dimensions=" & result.Count
    Set profilePanelFrame = frame
    Set profilePanelFeatures = features
    Set profilePanelOpenings = openings
    Set PlanPanelDimensions = result
End Function

Private Sub PanelFrameDimensions(ByVal result As Collection, ByVal groups As Object, ByVal frame As Collection)
    Dim parents As Object, request As Variant, a As Object, b As Object, point As Variant, coordinate As String, other As Variant
    Dim first As String, last As String, index As Long, axis As Long, selected As Long, count As Long
    Dim used As Object, best As Double, length As Double
    Set parents = CreateObject("Scripting.Dictionary")
    For Each request In result
        axis = request("axis")
        point = request("projectedFirst")
        other = request("projectedLast")
        first = ProfileRoot(parents, axis & "|" & ProfileCoordinateKey(point(axis)))
        last = ProfileRoot(parents, axis & "|" & ProfileCoordinateKey(other(axis)))
        parents(first) = last
    Next request
    Set used = CreateObject("Scripting.Dictionary")
    For count = 1 To frame.Count
        best = SEARCH_INITIAL_LIMIT
        selected = 0
        For index = 1 To frame.Count
            If Not used.Exists(index) Then
                Set a = PanelNode(frame, index)
                Set b = PanelNode(frame, index + 1)
                length = PanelDistance(a, b)
                If length < best And ProfileSegmentAxis(a, b) >= 0 Then
                    best = length
                    selected = index
                End If
            End If
        Next index
        If selected = 0 Then
            Exit For
        End If
        used(selected) = True
        Set a = PanelNode(frame, selected)
        Set b = PanelNode(frame, selected + 1)
        axis = ProfileSegmentAxis(a, b)
        If axis >= 0 Then
            coordinate = "x"
            If axis = SCREEN_Y_AXIS_INDEX Then
                coordinate = "y"
            End If
            first = ProfileRoot(parents, axis & "|" & ProfileCoordinateKey(a(coordinate)))
            last = ProfileRoot(parents, axis & "|" & ProfileCoordinateKey(b(coordinate)))
            If first <> last Then
                PanelAddRequest result, groups, a, b, axis, PROFILE_ROLE_OUTER, "frame-edge", "frame" & selected, True
                parents(first) = last
            End If
        End If
    Next count
End Sub

Private Sub PanelOpeningDimensions(ByVal result As Collection, ByVal groups As Object, ByVal openings As Collection, ByVal datumX As Object, ByVal datumY As Object)
    Dim points As Variant, a As Object, b As Object, corner As Object, index As Long, axis As Long, id As String
    For Each points In openings
        index = index + 1
        id = "opening" & index
        Set corner = points(1)
        For Each a In points
            If a("x") + a("y") < corner("x") + corner("y") Then
                Set corner = a
            End If
        Next a
        For axis = SCREEN_X_AXIS_INDEX To SCREEN_Y_AXIS_INDEX
            Set a = PanelNode(points, axis + 1)
            Set b = PanelNode(points, axis + 2)
            If ProfileSegmentAxis(a, b) <> axis Then
                Set a = PanelNode(points, axis + 2)
                Set b = PanelNode(points, axis + 3)
            End If
            PanelAddRequest result, groups, a, b, axis, PROFILE_ROLE_SLOT, "opening-size", id, True
        Next axis
        PanelAddRequest result, groups, datumX, corner, SCREEN_X_AXIS_INDEX, PROFILE_ROLE_OUTER, "opening-location", id, True
        PanelAddRequest result, groups, datumY, corner, SCREEN_Y_AXIS_INDEX, PROFILE_ROLE_OUTER, "opening-location", id, True
    Next points
End Sub
