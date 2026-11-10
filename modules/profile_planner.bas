Attribute VB_Name = "profile_planner"
Option Explicit

Private Function ProfilePlane(ByVal body As SldWorks.Body2) As SldWorks.Face2
    Dim faces As Variant, item As Variant, face As SldWorks.Face2, candidate As SldWorks.Face2
    Dim area As Double, normal As Variant, mainNormal As Variant, box As Variant, axis As Long, depth As Double, span As Double
    faces = body.GetFaces
    For Each item In faces
        Set face = item
        If face.GetSurface.IsPlane Then
            If face.GetArea > area Then
                area = face.GetArea
                Set candidate = face
            End If
        End If
    Next item
    If candidate Is Nothing Then
        Exit Function
    End If
    normal = candidate.Normal
    mainNormal = normal
    box = body.GetBodyBox
    span = SEARCH_INITIAL_LIMIT
    For axis = SCREEN_X_AXIS_INDEX To SPACE_LAST_AXIS_INDEX
        If Abs(normal(axis)) > FACE_NORMAL_ALIGNMENT_COSINE Then
            depth = box(axis + BBOX_MAX_OFFSET) - box(axis)
        Else
            span = MinD(span, box(axis + BBOX_MAX_OFFSET) - box(axis))
        End If
    Next axis
    If depth <= NATIVE_GEOMETRY_KEY_TOLERANCE_M Or depth > span * PROFILE_MAX_THICKNESS_RATIO Then
        Exit Function
    End If
    For Each item In faces
        Set face = item
        If Not face.GetSurface.IsPlane Then
            Exit Function
        End If
        normal = face.Normal
        If Abs(normal(0) * mainNormal(0) + normal(1) * mainNormal(1) + normal(2) * mainNormal(2)) > DIRECTION_ZERO_TOLERANCE Then
            If Abs(face.GetArea - area) > area * PROFILE_PARALLEL_AREA_TOLERANCE Then
                Exit Function
            End If
        End If
    Next item
    Set ProfilePlane = candidate
End Function

Private Function ProfilePointClose(ByVal a As Variant, ByVal b As Variant) As Boolean
    Dim axis As Long
    ProfilePointClose = True
    For axis = SCREEN_X_AXIS_INDEX To SPACE_LAST_AXIS_INDEX
        If Abs(a(axis) - b(axis)) > PROFILE_COORDINATE_TOLERANCE_M Then
            ProfilePointClose = False
        End If
    Next axis
End Function

Private Function ProfileLoopPoints(ByVal boundary As SldWorks.Loop2) As Collection
    Dim edges As Variant, item As Variant, edge As SldWorks.Edge, first As SldWorks.Vertex, last As SldWorks.Vertex
    Dim segments As New Collection, result As New Collection, used As Object, segment As Variant, point As Variant
    Dim index As Long, found As Boolean
    edges = boundary.GetEdges
    For Each item In edges
        Set edge = item
        If Not edge.GetCurve.IsLine Then
            Exit Function
        End If
        Set first = edge.GetStartVertex
        Set last = edge.GetEndVertex
        If first Is Nothing Or last Is Nothing Then
            Exit Function
        End If
        If Not ProfileTopologyPointClose(first.GetPoint, last.GetPoint) Then
            segments.Add Array(first.GetPoint, last.GetPoint)
        End If
    Next item
    If segments.Count < RECTANGLE_EDGE_COUNT Then
        Exit Function
    End If
    Set used = CreateObject("Scripting.Dictionary")
    segment = segments(1)
    result.Add segment(0)
    point = segment(1)
    used(1) = True
    Do While used.Count < segments.Count
        result.Add point
        found = False
        For index = 1 To segments.Count
            If Not used.Exists(index) Then
                segment = segments(index)
                If ProfileTopologyPointClose(point, segment(0)) Then
                    point = segment(1)
                    found = True
                ElseIf ProfileTopologyPointClose(point, segment(1)) Then
                    point = segment(0)
                    found = True
                End If
                If found Then
                    used(index) = True
                    Exit For
                End If
            End If
        Next index
        If Not found Then
            Exit Function
        End If
    Loop
    If Not ProfileTopologyPointClose(point, result(1)) Then
        Exit Function
    End If
    Set ProfileLoopPoints = result
End Function

' A dimensional tolerance must not join different vertices in a narrow slot.
Private Function ProfileTopologyPointClose(ByVal a As Variant, ByVal b As Variant) As Boolean
    Dim axis As Long
    ProfileTopologyPointClose = True
    For axis = SCREEN_X_AXIS_INDEX To SPACE_LAST_AXIS_INDEX
        If Abs(a(axis) - b(axis)) > PROFILE_TOPOLOGY_TOLERANCE_M Then
            ProfileTopologyPointClose = False
        End If
    Next axis
End Function

Private Function ProfileProjection(ByVal points As Collection, ByVal view As SldWorks.View) As Collection
    Dim result As New Collection, item As Variant, matrix As Variant, node As Object
    Dim index As Long, previous As Object, following As Object, changed As Boolean
    matrix = view.ModelToViewTransform.ArrayData
    For Each item In points
        Set node = CreateObject("Scripting.Dictionary")
        node.Add "point", item
        node.Add "x", item(0) * matrix(MATRIX_XX_INDEX) + item(1) * matrix(MATRIX_YX_INDEX) + item(2) * matrix(MATRIX_ZX_INDEX)
        node.Add "y", item(0) * matrix(MATRIX_XY_INDEX) + item(1) * matrix(MATRIX_YY_INDEX) + item(2) * matrix(MATRIX_ZY_INDEX)
        result.Add node
    Next item
    Do
        changed = False
        For index = result.Count To 1 Step -1
            Set previous = result((index + result.Count - 2) Mod result.Count + 1)
            Set node = result(index)
            Set following = result(index Mod result.Count + 1)
            If (Abs(previous("x") - node("x")) < PROFILE_COORDINATE_TOLERANCE_M And Abs(following("x") - node("x")) < PROFILE_COORDINATE_TOLERANCE_M) Or (Abs(previous("y") - node("y")) < PROFILE_COORDINATE_TOLERANCE_M And Abs(following("y") - node("y")) < PROFILE_COORDINATE_TOLERANCE_M) Then
                result.Remove index
                changed = True
                Exit For
            End If
        Next index
    Loop While changed And result.Count > RECTANGLE_EDGE_COUNT
    Set ProfileProjection = result
End Function

Private Function ProfileSegmentAxis(ByVal a As Object, ByVal b As Object) As Long
    ProfileSegmentAxis = -1
    If Abs(a("x") - b("x")) < PROFILE_COORDINATE_TOLERANCE_M And Abs(a("y") - b("y")) > PROFILE_COORDINATE_TOLERANCE_M Then
        ProfileSegmentAxis = SCREEN_Y_AXIS_INDEX
    ElseIf Abs(a("y") - b("y")) < PROFILE_COORDINATE_TOLERANCE_M And Abs(a("x") - b("x")) > PROFILE_COORDINATE_TOLERANCE_M Then
        ProfileSegmentAxis = SCREEN_X_AXIS_INDEX
    End If
End Function

Private Function ProfileNotch(ByVal a As Object, ByVal b As Object, ByVal c As Object, ByVal d As Object, ByVal area As Double) As Boolean
    Dim first As Long, middle As Long, last As Long, turnA As Double, turnB As Double, coordinate As String
    first = ProfileSegmentAxis(a, b)
    middle = ProfileSegmentAxis(b, c)
    last = ProfileSegmentAxis(c, d)
    If first < 0 Or middle < 0 Or first <> last Or first = middle Then
        Exit Function
    End If
    coordinate = "x"
    If first = SCREEN_Y_AXIS_INDEX Then
        coordinate = "y"
    End If
    If Abs(a(coordinate) - d(coordinate)) > PROFILE_COORDINATE_TOLERANCE_M Then
        Exit Function
    End If
    turnA = (b("x") - a("x")) * (c("y") - b("y")) - (b("y") - a("y")) * (c("x") - b("x"))
    turnB = (c("x") - b("x")) * (d("y") - c("y")) - (c("y") - b("y")) * (d("x") - c("x"))
    ProfileNotch = turnA * area < 0 And turnB * area < 0
End Function

Private Function ProfileCoordinateKey(ByVal value As Double) As String
    ProfileCoordinateKey = CStr(Round(value / PROFILE_COORDINATE_TOLERANCE_M, 0))
End Function

Private Sub ProfileCandidate(ByVal candidates As Object, ByVal a As Object, ByVal b As Object, ByVal axis As Long, ByVal role As String, ByVal feature As String, ByVal overall As Boolean)
    Dim coordinate As String, low As String, high As String, key As String, request As Object, features As Object
    coordinate = "x"
    If axis = SCREEN_Y_AXIS_INDEX Then
        coordinate = "y"
    End If
    low = ProfileCoordinateKey(MinD(a(coordinate), b(coordinate)))
    high = ProfileCoordinateKey(-MinD(-a(coordinate), -b(coordinate)))
    If low = high Then
        Exit Sub
    End If
    key = CStr(axis) & "|" & low & "|" & high
    If Not candidates.Exists(key) Then
        Set request = CreateObject("Scripting.Dictionary")
        request.Add "key", key
        request.Add "low", low
        request.Add "high", high
        request.Add "axis", axis
        request.Add "role", role
        request.Add "overall", overall
        request.Add "value", Abs(a(coordinate) - b(coordinate))
        request.Add "first", a("point")
        request.Add "last", b("point")
        Set features = CreateObject("Scripting.Dictionary")
        request.Add "features", features
        candidates.Add key, request
    End If
    Set request = candidates(key)
    If role = PROFILE_ROLE_SLOT Then
        request("role") = role
    End If
    If overall Then
        request("overall") = True
    End If
    If Len(feature) > 0 Then
        request("features")(feature) = True
    End If
End Sub

Private Function ProfileRoot(ByVal parents As Object, ByVal key As String) As String
    If Not parents.Exists(key) Then
        parents.Add key, key
    End If
    Do While CStr(parents(key)) <> key
        key = CStr(parents(key))
    Loop
    ProfileRoot = key
End Function

Private Function PlanProfileDimensions(ByVal outer As Collection, ByVal openings As Collection) As Collection
    Dim candidates As Object, roles As Object, parents As Object, result As New Collection, node As Object, other As Object
    Dim a As Object, b As Object, c As Object, last As Object, points As Collection, opening As Variant
    Dim lowX As Object, highX As Object, lowY As Object, highY As Object
    Dim area As Double, index As Long, offset As Long, axis As Long, role As String, feature As String, key As Variant
    Dim request As Object, winner As Object, priority As Double, best As Double, rootA As String, rootB As String, visited As Object
    Dim panelPlan As Collection
    Set panelPlan = PlanPanelDimensions(outer, openings)
    If Not panelPlan Is Nothing Then
        Set PlanProfileDimensions = panelPlan
        Exit Function
    End If
    Set candidates = CreateObject("Scripting.Dictionary")
    Set roles = CreateObject("Scripting.Dictionary")
    Set parents = CreateObject("Scripting.Dictionary")
    Set visited = CreateObject("Scripting.Dictionary")
    Set lowX = outer(1)
    Set highX = lowX
    Set lowY = lowX
    Set highY = lowX
    For index = 1 To outer.Count
        Set a = outer(index)
        Set b = outer(index Mod outer.Count + 1)
        If ProfileSegmentAxis(a, b) < 0 Then
            Exit Function
        End If
        area = area + a("x") * b("y") - b("x") * a("y")
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
    Next index
    For index = 1 To outer.Count
        Set a = outer(index)
        Set b = outer(index Mod outer.Count + 1)
        Set c = outer((index + 1) Mod outer.Count + 1)
        Set last = outer((index + 2) Mod outer.Count + 1)
        If ProfileNotch(a, b, c, last, area) Then
            For offset = 0 To PROFILE_NOTCH_EDGE_COUNT - 1
                roles((index + offset - 1) Mod outer.Count + 1) = CStr(index)
            Next offset
        End If
    Next index
    ProfileCandidate candidates, lowX, highX, SCREEN_X_AXIS_INDEX, PROFILE_ROLE_OUTER, "", True
    ProfileCandidate candidates, lowY, highY, SCREEN_Y_AXIS_INDEX, PROFILE_ROLE_OUTER, "", True
    For index = 1 To outer.Count
        Set a = outer(index)
        Set b = outer(index Mod outer.Count + 1)
        role = PROFILE_ROLE_OUTER
        feature = ""
        If roles.Exists(index) Then
            role = PROFILE_ROLE_SLOT
            feature = CStr(roles(index))
        End If
        ProfileCandidate candidates, a, b, ProfileSegmentAxis(a, b), role, feature, False
    Next index
    index = 0
    For Each opening In openings
        Set points = opening
        index = index + 1
        For offset = 1 To points.Count
            Set a = points(offset)
            Set b = points(offset Mod points.Count + 1)
            axis = ProfileSegmentAxis(a, b)
            If axis < 0 Then
                Exit Function
            End If
            ProfileCandidate candidates, a, b, axis, PROFILE_ROLE_SLOT, "opening" & index, False
        Next offset
        ProfileCandidate candidates, lowX, points(1), SCREEN_X_AXIS_INDEX, PROFILE_ROLE_SLOT, "", False
        ProfileCandidate candidates, lowY, points(1), SCREEN_Y_AXIS_INDEX, PROFILE_ROLE_SLOT, "", False
    Next opening
    Do While visited.Count < candidates.Count
        best = SEARCH_INITIAL_LIMIT
        Set winner = Nothing
        For Each key In candidates.Keys
            If Not visited.Exists(key) Then
                Set request = candidates(key)
                priority = request("value")
                If request("role") = PROFILE_ROLE_OUTER Then
                    priority = priority + PROFILE_OUTER_PRIORITY_COST
                End If
                If CBool(request("overall")) Then
                    priority = -PROFILE_OUTER_PRIORITY_COST
                End If
                If priority < best Then
                    best = priority
                    Set winner = request
                End If
            End If
        Next key
        visited(winner("key")) = True
        rootA = ProfileRoot(parents, CStr(winner("axis")) & "|" & winner("low"))
        rootB = ProfileRoot(parents, CStr(winner("axis")) & "|" & winner("high"))
        winner.Add "selected", rootA <> rootB
        If rootA <> rootB Then
            parents(rootA) = rootB
            result.Add winner
        End If
        Log "PROFILE PLAN " & winner("key") & "; role=" & winner("role") & "; value=" & winner("value") & "; selected=" & CStr(rootA <> rootB)
    Loop
    Set PlanProfileDimensions = result
End Function
