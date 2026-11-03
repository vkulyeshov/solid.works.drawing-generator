Attribute VB_Name = "inventory"
Option Explicit
Private Function SuppressionFilterEnabled() As Boolean
    Dim value As String
    value = Setting("ExcludeSuppressed", SUPPRESSION_FILTER_DISABLED)
    If value <> SUPPRESSION_FILTER_DISABLED And value <> SUPPRESSION_FILTER_ENABLED Then
        Err.Raise vbObjectError + ERR_DRAWING_SETTINGS, , "ExcludeSuppressed must be 0 or 1."
    End If
    SuppressionFilterEnabled = (value = SUPPRESSION_FILTER_ENABLED)
End Function

Private Function CopiedInventoryPath(ByVal referencePath As String) As String
    Dim name As String, candidate As String
    name = fs.GetFileName(referencePath)
    candidate = runDir & MODELS_RELATIVE_DIRECTORY & name
    If fs.FileExists(candidate) = False Then
        candidate = runDir & MODELS_RELATIVE_DIRECTORY & prefix & name
    End If
    If Len(name) = 0 Or fs.FileExists(candidate) = False Then
        Err.Raise vbObjectError + ERR_INVENTORY_REFERENCE, , "Copied component is absent from Models: " & referencePath
    End If
    CopiedInventoryPath = fs.GetAbsolutePathName(candidate)
End Function

Private Sub RememberLoadedInventoryModel(ByVal component As Object, ByVal path As String, ByVal models As Object)
    Dim doc As Object, loadedPath As String
    ' Suppressed and lightweight occurrences may have no model document.
    ' An unavailable or invalidated optional pointer falls back to opening its file.
    On Error Resume Next
    Set doc = component.GetModelDoc2
    If Not doc Is Nothing Then
        loadedPath = doc.GetPathName
        If Err.Number = 0 Then
            If StrComp(loadedPath, path, vbTextCompare) = 0 Then
                Set models(path) = doc
            End If
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Function OpenInventoryModel(ByVal path As String, ByVal cfg As String, ByVal docType As Long, ByVal models As Object) As Object
    Dim doc As Object, error As Long, warning As Long, loadedPath As String
    If models.Exists(path) Then
        On Error Resume Next
        Set doc = models(path)
        loadedPath = doc.GetPathName
        If Err.Number <> 0 Then
            Set doc = Nothing
        End If
        Err.Clear
        On Error GoTo 0
        If StrComp(loadedPath, path, vbTextCompare) <> 0 Then
            Set doc = Nothing
        End If
    End If
    If doc Is Nothing Then
        Set doc = sw.GetOpenDocumentByName(path)
    End If
    If doc Is Nothing Then
        Log "INVENTORY opening copied reference: " & path & " | " & cfg
        Set doc = sw.OpenDoc6(path, docType, SW_OPEN_SILENT, cfg, error, warning)
    End If
    If doc Is Nothing Then
        Err.Raise vbObjectError + ERR_INVENTORY_REFERENCE, , "Cannot open copied component: " & path & "; error=" & error
    End If
    If StrComp(doc.GetPathName, path, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + ERR_INVENTORY_REFERENCE, , "SolidWorks opened a different component: " & doc.GetPathName
    End If
    Set models(path) = doc
    Set OpenInventoryModel = doc
End Function

Private Function RememberInventoryModel(ByVal doc As Object, ByVal cfg As String) As Object
    Dim key As String, job As Object
    key = LCase$(doc.GetPathName & "|" & cfg)
    If quantities.Exists(key) Then
        Set job = quantities(key)
        job("qty") = CLng(job("qty")) + 1
    Else
        Set job = CreateObject("Scripting.Dictionary")
        job.Add "doc", doc
        job.Add "cfg", cfg
        job.Add "qty", 1
        quantities.Add key, job
    End If
    Set RememberInventoryModel = job
End Function

Private Function SavedInventoryReferences(ByVal doc As Object, ByVal cfg As String, ByVal plans As Object, ByVal models As Object) As Collection
    Dim key As String, components As Variant, item As Variant, component As Object, model As Object
    Dim references As New Collection, reference As Object, path As String, extension As String, docType As Long
    Dim stage As String, errorNumber As Long, errorText As String
    On Error GoTo InventoryFailed
    stage = "Reading assembly configuration"
    Log "INVENTORY reading saved assembly: " & doc.GetPathName & " | " & cfg
    key = LCase$(doc.GetPathName & "|" & cfg)
    If plans.Exists(key) Then
        Set SavedInventoryReferences = plans(key)
        Exit Function
    End If
    If doc.GetConfigurationByName(cfg) Is Nothing Then
        Err.Raise vbObjectError + ERR_CONFIGURATION, , "Configuration absent: " & cfg & " | " & doc.GetPathName
    End If
    If doc.ConfigurationManager.ActiveConfiguration.Name <> cfg Then
        If doc.ShowConfiguration2(cfg) = False Then
            Err.Raise vbObjectError + ERR_CONFIGURATION, , "Cannot activate inventory configuration: " & cfg
        End If
    End If
    If doc.ConfigurationManager.ActiveConfiguration.Name <> cfg Then
        Err.Raise vbObjectError + ERR_CONFIGURATION, , "Wrong inventory configuration: " & cfg
    End If
    ' Copy plain reference data before opening children or changing any other
    ' document's configuration, which can invalidate component COM arrays.
    stage = "Reading top-level components"
    components = AsAssembly(doc).GetComponents(True)
    If Not IsEmpty(components) Then
        For Each item In components
            Set component = item
            stage = "Mapping copied component path"
            path = CopiedInventoryPath(component.GetPathName)
            RememberLoadedInventoryModel component, path, models
            stage = "Reading component type: " & path
            extension = LCase$(fs.GetExtensionName(path))
            Select Case extension
                Case ASSEMBLY_EXTENSION
                    docType = SW_DOC_ASSEMBLY
                Case PART_EXTENSION
                    docType = SW_DOC_PART
                Case Else
                    Set model = component.GetModelDoc2
                    If model Is Nothing Then
                        Err.Raise vbObjectError + ERR_INVENTORY_REFERENCE, , "Cannot determine copied component type: " & path
                    End If
                    docType = model.GetType
            End Select
            Set reference = CreateObject("Scripting.Dictionary")
            reference.Add "path", path
            stage = "Reading referenced configuration: " & path
            reference.Add "cfg", CStr(component.ReferencedConfiguration)
            reference.Add "type", docType
            references.Add reference
        Next item
    End If
    stage = "Caching assembly references"
    plans.Add key, references
    Log "INVENTORY saved references=" & references.Count & " | " & doc.GetPathName & " | " & cfg
    Set SavedInventoryReferences = references
    Exit Function
InventoryFailed:
    errorNumber = Err.Number
    errorText = Err.Description
    Err.Raise errorNumber, "SavedInventoryReferences", stage & " | " & doc.GetPathName & " | " & cfg & ": " & errorText
End Function

Private Sub InventorySavedAssembly(ByVal asm As Object, ByVal cfg As String, ByVal plans As Object, ByVal stack As Object, ByVal models As Object)
    Dim key As String, references As Collection, reference As Object, doc As Object, job As Object, childKey As String
    key = LCase$(asm.GetPathName & "|" & cfg)
    If stack.Exists(key) Then
        Err.Raise vbObjectError + ERR_INVENTORY_REFERENCE, , "Cyclic assembly reference: " & key
    End If
    stack.Add key, True
    Set references = SavedInventoryReferences(asm, cfg, plans, models)
    For Each reference In references
        childKey = LCase$(CStr(reference("path")) & "|" & CStr(reference("cfg")))
        If quantities.Exists(childKey) Then
            Set job = quantities(childKey)
            Set doc = job("doc")
        Else
            Set doc = OpenInventoryModel(CStr(reference("path")), CStr(reference("cfg")), CLng(reference("type")), models)
        End If
        Set job = RememberInventoryModel(doc, CStr(reference("cfg")))
        If CLng(reference("type")) = SW_DOC_ASSEMBLY Then
            InventorySavedAssembly doc, CStr(reference("cfg")), plans, stack, models
        End If
    Next reference
    stack.Remove key
End Sub

Private Sub InventoryAllSavedComponents(ByVal asm As Object, ByVal cfg As String)
    Dim plans As Object, stack As Object, models As Object
    If InStr(1, asm.GetPathName, runDir & MODELS_RELATIVE_DIRECTORY, vbTextCompare) <> 1 Then
        Err.Raise vbObjectError + ERR_INVENTORY_REFERENCE, , "Saved-reference inventory requires a copied assembly in Models."
    End If
    Set plans = CreateObject("Scripting.Dictionary")
    plans.CompareMode = vbTextCompare
    Set stack = CreateObject("Scripting.Dictionary")
    stack.CompareMode = vbTextCompare
    Set models = CreateObject("Scripting.Dictionary")
    models.CompareMode = vbTextCompare
    InventorySavedAssembly asm, cfg, plans, stack, models
    Log "INVENTORY all saved configurations=" & plans.Count & "; unique component configurations=" & quantities.Count
End Sub

Private Function ComponentOrParentSuppressed(ByVal component As Object) As Boolean
    Dim current As Object
    If excludeSuppressed = False Then
        Exit Function
    End If
    Set current = component
    Do While Not current Is Nothing
        ' Test the occurrence in the active assembly context, not its document.
        ' Lightweight states are distinct from suppression.
        If current.GetSuppression2 = SW_COMPONENT_SUPPRESSED Then
            ComponentOrParentSuppressed = True
            Exit Function
        End If
        Set current = current.GetParent
    Loop
End Function

Private Sub Inventory(ByVal asm As SldWorks.ModelDoc2)
    Dim comps As Variant, i As Long, c As Object, doc As Object, j As Object, k As String
    comps = AsAssembly(asm).GetComponents(False)
    If IsEmpty(comps) Then
        Exit Sub
    End If
    For i = LBound(comps) To UBound(comps)
        Set c = comps(i)
        If ComponentOrParentSuppressed(c) Then
            Log "INVENTORY skipped suppressed component or descendant: " & c.Name2
        Else
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
                Set j = RememberInventoryModel(doc, CStr(c.ReferencedConfiguration))
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

