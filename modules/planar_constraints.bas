Attribute VB_Name = "planar_constraints"
Option Explicit

Private Function PlanarAcuteAngle(ByVal dx As Double, ByVal dy As Double) As Double
    Dim maximum As Double
    maximum = -MinD(-Abs(dx), -Abs(dy))
    If maximum > 0 Then
        PlanarAcuteAngle = Atn(MinD(Abs(dx), Abs(dy)) / maximum)
    End If
End Function

Private Function PlanarDistance(ByVal a As Object, ByVal b As Object) As Double
    PlanarDistance = Sqr((b("x") - a("x")) ^ 2 + (b("y") - a("y")) ^ 2)
End Function

Private Function PlanarMirror(ByVal point As Object, ByVal symmetry As Variant) As Variant
    Dim distance As Double
    distance = (point("x") - symmetry(2)) * symmetry(0) + (point("y") - symmetry(3)) * symmetry(1)
    PlanarMirror = Array(point("x") - 2 * distance * symmetry(0), point("y") - 2 * distance * symmetry(1))
End Function

Private Function PlanarPointIndex(ByVal points As Collection, ByVal point As Variant) As Long
    Dim index As Long, node As Object
    For index = 1 To points.Count
        Set node = points(index)
        If Abs(node("x") - point(0)) <= PROFILE_COORDINATE_TOLERANCE_M And Abs(node("y") - point(1)) <= PROFILE_COORDINATE_TOLERANCE_M Then
            PlanarPointIndex = index
            Exit Function
        End If
    Next index
End Function

' Check the complete boundary and all hole centers, not just a bounding box.
Private Function PlanarSymmetries(ByVal points As Collection, ByVal holes As Collection) As Collection
    Dim result As New Collection, node As Object, following As Object, item As Variant, hole As Object, other As Object
    Dim centerX As Double, centerY As Double, dx As Double, dy As Double, length As Double, symmetry As Variant
    Dim index As Long, perpendicular As Long, a As Long, b As Long, valid As Boolean, found As Boolean, reflected As Variant
    For Each node In points
        centerX = centerX + node("x") / points.Count
        centerY = centerY + node("y") / points.Count
    Next node
    For index = 1 To points.Count
        Set node = points(index)
        Set following = points(index Mod points.Count + 1)
        length = PlanarDistance(node, following)
        dx = (following("x") - node("x")) / length
        dy = (following("y") - node("y")) / length
        For perpendicular = SCREEN_X_AXIS_INDEX To SCREEN_Y_AXIS_INDEX
            symmetry = Array(dx, dy, centerX, centerY)
            If perpendicular = SCREEN_Y_AXIS_INDEX Then
                symmetry = Array(-dy, dx, centerX, centerY)
            End If
            valid = True
            For a = 1 To points.Count
                b = PlanarPointIndex(points, PlanarMirror(points(a), symmetry))
                If b = 0 Then
                    valid = False
                    Exit For
                End If
                reflected = PlanarMirror(points(a Mod points.Count + 1), symmetry)
                If PlanarPointIndex(points, reflected) <> (b + points.Count - 2) Mod points.Count + 1 Then
                    valid = False
                    Exit For
                End If
            Next a
            If valid Then
                For Each item In holes
                    Set hole = item
                    reflected = PlanarMirror(hole, symmetry)
                    found = False
                    For Each other In holes
                        If Abs(other("x") - reflected(0)) <= PROFILE_COORDINATE_TOLERANCE_M And Abs(other("y") - reflected(1)) <= PROFILE_COORDINATE_TOLERANCE_M And Abs(other("radius") - hole("radius")) <= PROFILE_COORDINATE_TOLERANCE_M Then
                            found = True
                            Exit For
                        End If
                    Next other
                    If Not found Then
                        valid = False
                        Exit For
                    End If
                Next item
            End If
            If valid Then
                result.Add symmetry
            End If
        Next perpendicular
    Next index
    Set PlanarSymmetries = result
End Function

Private Sub PlanarJoin(ByVal parents As Object, ByVal first As Long, ByVal last As Long)
    parents(ProfileRoot(parents, CStr(first))) = ProfileRoot(parents, CStr(last))
End Sub

' Length variables are grouped only by a verified whole-profile reflection or a
' locally mirror-symmetric three-edge U feature. Loop closure removes two further
' independent variables; the omitted lengths are reconstructed and checked.
Private Function PlanarEdgePlan(ByVal points As Collection, ByVal holes As Collection) As Collection
    Dim parents As Object, groups As Object, symmetries As Collection, symmetry As Variant, group As Object
    Dim result As New Collection, node As Object, following As Object, a As Object, b As Object, c As Object, last As Object
    Dim index As Long, mapped As Long, other As Long, root As String, length As Double, dot As Double, cross As Double
    Dim firstDrop As Object, lastDrop As Object, candidate As Object, second As Object, determinant As Double, cost As Double, best As Double
    Dim sumX As Double, sumY As Double, inferredA As Double, inferredB As Double
    Set parents = CreateObject("Scripting.Dictionary")
    Set groups = CreateObject("Scripting.Dictionary")
    Set symmetries = PlanarSymmetries(points, holes)
    For index = 1 To points.Count
        root = ProfileRoot(parents, CStr(index))
        For Each symmetry In symmetries
            mapped = PlanarPointIndex(points, PlanarMirror(points(index Mod points.Count + 1), symmetry))
            PlanarJoin parents, index, mapped
        Next symmetry
        Set a = points(index)
        Set b = points(index Mod points.Count + 1)
        Set c = points((index + 1) Mod points.Count + 1)
        Set last = points((index + 2) Mod points.Count + 1)
        length = PlanarDistance(a, b)
        If points.Count > RECTANGLE_EDGE_COUNT And Abs(length - PlanarDistance(c, last)) <= PROFILE_COORDINATE_TOLERANCE_M Then
            dot = (b("x") - a("x")) * (c("x") - b("x")) + (b("y") - a("y")) * (c("y") - b("y"))
            cross = (b("x") - a("x")) * (last("y") - c("y")) - (b("y") - a("y")) * (last("x") - c("x"))
            If Abs(dot) <= PROFILE_COORDINATE_TOLERANCE_M * length And Abs(cross) <= PROFILE_COORDINATE_TOLERANCE_M * length Then
                If (b("x") - a("x")) * (last("x") - c("x")) + (b("y") - a("y")) * (last("y") - c("y")) < 0 Then
                    PlanarJoin parents, index, (index + 1) Mod points.Count + 1
                End If
            End If
        End If
    Next index
    For index = 1 To points.Count
        Set node = points(index)
        Set following = points(index Mod points.Count + 1)
        length = PlanarDistance(node, following)
        root = ProfileRoot(parents, CStr(index))
        If Not groups.Exists(root) Then
            Set group = CreateObject("Scripting.Dictionary")
            group.Add "id", root
            group.Add "index", index
            group.Add "first", node
            group.Add "last", following
            group.Add "value", length
            group.Add "count", 0
            group.Add "dx", 0#
            group.Add "dy", 0#
            group.Add "selected", True
            groups.Add root, group
            result.Add group
        End If
        Set group = groups(root)
        If Abs(group("value") - length) > PROFILE_COORDINATE_TOLERANCE_M Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "A verified planar symmetry has unequal lengths."
        End If
        group("count") = group("count") + 1
        group("dx") = group("dx") + (following("x") - node("x")) / length
        group("dy") = group("dy") + (following("y") - node("y")) / length
    Next index
    best = SEARCH_INITIAL_LIMIT
    For Each candidate In result
        For Each second In result
            determinant = candidate("dx") * second("dy") - candidate("dy") * second("dx")
            cost = candidate("value") + second("value")
            If Abs(determinant) > PLANAR_CLOSURE_DETERMINANT_MIN And cost < best Then
                best = cost
                Set firstDrop = candidate
                Set lastDrop = second
            End If
        Next second
    Next candidate
    If Not firstDrop Is Nothing Then
        firstDrop("selected") = False
        lastDrop("selected") = False
        For Each group In result
            If CBool(group("selected")) Then
                sumX = sumX - group("dx") * group("value")
                sumY = sumY - group("dy") * group("value")
            End If
        Next group
        determinant = firstDrop("dx") * lastDrop("dy") - firstDrop("dy") * lastDrop("dx")
        inferredA = (sumX * lastDrop("dy") - sumY * lastDrop("dx")) / determinant
        inferredB = (firstDrop("dx") * sumY - firstDrop("dy") * sumX) / determinant
        If Abs(inferredA - firstDrop("value")) > PROFILE_COORDINATE_TOLERANCE_M Or Abs(inferredB - lastDrop("value")) > PROFILE_COORDINATE_TOLERANCE_M Then
            Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "Planar contour closure does not reconstruct omitted lengths."
        End If
    End If
    Log "PLANAR CONSTRAINTS edges=" & points.Count & "; independent length groups=" & groups.Count & "; verified reflection candidates=" & symmetries.Count & "; closure variables=" & CStr(Not firstDrop Is Nothing)
    Set PlanarEdgePlan = result
End Function
