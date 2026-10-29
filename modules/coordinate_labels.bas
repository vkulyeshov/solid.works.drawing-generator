Attribute VB_Name = "coordinate_labels"
Option Explicit
Private Function UseCoordinateDots(ByVal span As Double) As Boolean
    UseCoordinateDots = span >= GOST_CHAIN_DOT_DIAMETER_M + DIMENSION_GEOMETRY_GAP_M And span < GOST_ARROW_HEAD_LENGTH_M * 2
End Function

Private Function CoordinateMarkerPadding(ByVal span As Double, ByVal textWidth As Double) As Double
    If UseCoordinateDots(span) Then
        CoordinateMarkerPadding = GOST_CHAIN_DOT_DIAMETER_M / 2 + DIMENSION_GEOMETRY_GAP_M
    ElseIf span < textWidth + GOST_ARROW_HEAD_LENGTH_M * 2 Then
        CoordinateMarkerPadding = GOST_ARROW_LENGTH_M + DIMENSION_GEOMETRY_GAP_M
    Else
        CoordinateMarkerPadding = DIMENSION_GEOMETRY_GAP_M
    End If
End Function

Private Function CoordinateTextCenter(ByVal coordinates As Object, ByVal keys As Variant, ByVal first As Double, ByVal last As Double, ByVal textWidth As Double, ByVal minimum As Double, ByVal maximum As Double) As Double
    Dim result As Double
    If Not FindCoordinateTextCenter(coordinates, keys, first, last, textWidth, minimum, maximum, result) Then
        Err.Raise vbObjectError + ERR_DIMENSION_LAYOUT, , "No clear dimension text position away from the extension lines."
    End If
    CoordinateTextCenter = result
End Function

Private Function FindCoordinateTextCenter(ByVal coordinates As Object, ByVal keys As Variant, ByVal first As Double, ByVal last As Double, ByVal textWidth As Double, ByVal minimum As Double, ByVal maximum As Double, ByRef result As Double) As Boolean
    Dim i As Long, start As Double, finish As Double, candidate As Double, desired As Double, best As Double
    Dim halfWidth As Double, padding As Double, distance As Double, valid As Boolean, found As Boolean
    desired = (first + last) / 2
    halfWidth = textWidth / 2
    padding = CoordinateMarkerPadding(last - first, textWidth)
    best = SEARCH_INITIAL_LIMIT
    start = minimum
    For i = LBound(keys) To UBound(keys) + 1
        finish = maximum
        If i <= UBound(keys) Then
            finish = MinD(maximum, CDbl(coordinates(keys(i))("coordinate")) - COORDINATE_LABEL_GAP_M)
        End If
        If finish - start >= textWidth Then
            candidate = -MinD(-start - halfWidth, -MinD(desired, finish - halfWidth))
            If candidate + halfWidth <= first Then
                candidate = MinD(candidate, first - padding - DIMENSION_GEOMETRY_GAP_M - halfWidth)
            ElseIf candidate - halfWidth >= last Then
                candidate = -MinD(-candidate, -last - padding - DIMENSION_GEOMETRY_GAP_M - halfWidth)
            End If
            valid = False
            If candidate + halfWidth <= first - padding - DIMENSION_GEOMETRY_GAP_M Then
                valid = True
            ElseIf candidate - halfWidth >= last + padding + DIMENSION_GEOMETRY_GAP_M Then
                valid = True
            ElseIf candidate - halfWidth >= first + DIMENSION_GEOMETRY_GAP_M And candidate + halfWidth <= last - DIMENSION_GEOMETRY_GAP_M Then
                valid = True
            End If
            valid = valid And candidate - halfWidth >= start And candidate + halfWidth <= finish
            distance = Abs(candidate - desired)
            If valid And distance < best Then
                result = candidate
                best = distance
                found = True
            End If
        End If
        If i <= UBound(keys) Then
            start = -MinD(-minimum, -CDbl(coordinates(keys(i))("coordinate")) - COORDINATE_LABEL_GAP_M)
        End If
    Next i
    FindCoordinateTextCenter = found
End Function

Private Function SortCoordinatePlans(ByVal plans As Collection) As Collection
    Dim sorted As New Collection, plan As Variant, i As Long, inserted As Boolean
    For Each plan In plans
        inserted = False
        For i = 1 To sorted.Count
            If CDbl(plan("lineStart")) < CDbl(sorted(i)("lineStart")) Then
                sorted.Add plan, Before:=i
                inserted = True
                Exit For
            End If
        Next i
        If Not inserted Then
            sorted.Add plan
        End If
    Next plan
    Set SortCoordinatePlans = sorted
End Function


Private Function DimensionLabel(ByVal value As Double) As String
    Dim text As String
    text = Format$(value, DIMENSION_NUMBER_FORMAT)
    ' Keep the configured precision; remove only insignificant trailing zeroes.
    Do While Right$(text, 1) = "0"
        text = Left$(text, Len(text) - 1)
    Loop
    If Right$(text, 1) = "." Or Right$(text, 1) = "," Then
        text = Left$(text, Len(text) - 1)
    End If
    DimensionLabel = text
End Function





