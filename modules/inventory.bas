Attribute VB_Name = "inventory"
Option Explicit
Private Sub Inventory(ByVal asm As SldWorks.ModelDoc2)
    Dim comps As Variant, i As Long, c As Object, doc As Object, j As Object, k As String
    comps = AsAssembly(asm).GetComponents(False)
    If IsEmpty(comps) Then
        Exit Sub
    End If
    For i = LBound(comps) To UBound(comps)
        Set c = comps(i)
        If c.IsSuppressed = False Then
            Set doc = c.GetModelDoc2
            If doc Is Nothing Then
                Warn "Unresolved component: " & c.Name2
            Else
                If InStr(1, doc.GetPathName, runDir & MODELS_RELATIVE_DIRECTORY, vbTextCompare) <> 1 Then
                    Dim localCopy As String, openError As Long, openWarning As Long
                    localCopy = runDir & MODELS_RELATIVE_DIRECTORY & prefix & fs.GetFileName(doc.GetPathName)
                    Log "LIBRARY using copied model: " & localCopy
                    If fs.FileExists(localCopy) = False Then
                        Err.Raise vbObjectError + ERR_LIBRARY_COMPONENT, , "Local copy absent: " & localCopy
                    End If
                    Set doc = sw.OpenDoc6(localCopy, doc.GetType, 1, c.ReferencedConfiguration, openError, openWarning)
                    If doc Is Nothing Then
                        Err.Raise vbObjectError + ERR_LIBRARY_COMPONENT, , "Cannot open copied library component: " & localCopy
                    End If
                End If
                k = LCase$(doc.GetPathName & "|" & c.ReferencedConfiguration)
                If quantities.Exists(k) Then
                    Set j = quantities(k): j("qty") = CLng(j("qty")) + 1
                Else
                    Set j = CreateObject("Scripting.Dictionary")
                    j.Add "doc", doc: j.Add "cfg", c.ReferencedConfiguration: j.Add "qty", 1
                    quantities.Add k, j
                End If
            End If
        End If
    Next i
End Sub

Private Sub AddModelJobs(ByVal doc As SldWorks.ModelDoc2, ByVal cfg As String, ByVal qty As Long)
    Dim bodies As Variant, i As Long, b As Object, k As String
    If doc.GetConfigurationByName(cfg) Is Nothing Then
        Err.Raise vbObjectError + ERR_CONFIGURATION, , "Configuration absent: " & cfg
    End If
    doc.ShowConfiguration2 cfg
    If doc.ConfigurationManager.ActiveConfiguration.Name <> cfg Then
        Err.Raise vbObjectError + ERR_CONFIGURATION, , "Configuration could not be activated: " & cfg
    End If
    If doc.GetType = SW_DOC_ASSEMBLY Then
        AddJob doc, cfg, "", "", qty, False
        Exit Sub
    End If
    bodies = AsPart(doc).GetBodies2(SW_SOLID_BODY, True)
    If IsEmpty(bodies) Then
        Warn "No visible solid bodies: " & doc.GetPathName
        Exit Sub
    End If
    If UBound(bodies) = LBound(bodies) Then
        Set b = bodies(LBound(bodies))
        AddJob doc, cfg, b.Name, "", qty, b.IsSheetMetal
    Else
        For i = LBound(bodies) To UBound(bodies)
            Set b = bodies(i)
            AddJob doc, cfg, b.Name, CutListName(doc, b.Name), qty, b.IsSheetMetal, True
        Next i
    End If
End Sub

Private Sub AddJob(ByVal doc As SldWorks.ModelDoc2, ByVal cfg As String, ByVal bodyName As String, ByVal cutName As String, ByVal qty As Long, ByVal sheetMetal As Boolean, Optional ByVal multiBody As Boolean = False)
    Dim j As Object, k As String
    k = LCase$(doc.GetPathName & "|" & cfg & "|" & bodyName)
    If jobs.Exists(k) Then
        Exit Sub
    End If
    Set j = CreateObject("Scripting.Dictionary")
    j.Add "doc", doc: j.Add "cfg", cfg: j.Add "body", bodyName
    j.Add "path", doc.GetPathName: j.Add "type", doc.GetType
    j.Add "cut", cutName: j.Add "qty", qty: j.Add "sheet", sheetMetal: j.Add "multi", multiBody
    jobs.Add k, j
End Sub

Private Function CutListName(ByVal doc As SldWorks.ModelDoc2, ByVal bodyName As String) As String
    Dim f As Object, bf As Object, bodies As Variant, b As Variant
    Set f = doc.FirstFeature
    Do While Not f Is Nothing
        If f.GetTypeName2 = "CutListFolder" Then
            Set bf = f.GetSpecificFeature2
            bodies = bf.GetBodies
            If Not IsEmpty(bodies) Then
                For Each b In bodies
                    If b.Name = bodyName Then
                        CutListName = f.Name
                        Exit Function
                    End If
                Next b
            End If
        End If
        Set f = f.GetNextFeature
    Loop
End Function

Private Function FindBody(ByVal doc As SldWorks.ModelDoc2, ByVal bodyName As String) As Object
    Dim a As Variant, b As Variant
    a = AsPart(doc).GetBodies2(SW_SOLID_BODY, True)
    If IsEmpty(a) Then
        Exit Function
    End If
    For Each b In a
        If b.Name = bodyName Then
            Set FindBody = b
            Exit Function
        End If
    Next b
End Function

