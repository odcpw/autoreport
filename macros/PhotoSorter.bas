Attribute VB_Name = "PhotoSorter"
Option Explicit

Public rootPath As String

Public Sub ShowPhotoSorter()
    PhotoSorterForm.Show vbModeless
End Sub

Public Function ChooseImageDirectory() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Ordner mit Fotos auswählen"
        .AllowMultiSelect = False
        If .Show = -1 Then
            ChooseImageDirectory = .SelectedItems(1)
        Else
            ChooseImageDirectory = ""
        End If
    End With
End Function

Public Sub ScanImagesIntoSheet(ByVal baseDirectory As String)
    If Len(baseDirectory) = 0 Then Exit Sub
    Dim images As Collection
    Set images = EnumerateImages(baseDirectory)

    Dim folderTagMap As Scripting.Dictionary
    Set folderTagMap = modABPhotosRepository.BuildFolderTagLookup()
    Dim entryMap As New Scripting.Dictionary
    entryMap.CompareMode = TextCompare

    Dim imageItem As Scripting.Dictionary
    Dim baseName As String
    Dim currentEntry As Scripting.Dictionary
    Dim relativePath As String

    For Each imageItem In images
        baseName = NzString(imageItem("baseName"))
        If Len(baseName) = 0 Then GoTo ContinueLoop

        If entryMap.Exists(baseName) Then
            Set currentEntry = entryMap(baseName)
        Else
            Set currentEntry = modABPhotosRepository.GetPhotoEntry(baseName)
            If currentEntry Is Nothing Then
                Set currentEntry = New Scripting.Dictionary
                currentEntry.CompareMode = TextCompare
                currentEntry("fileName") = baseName
                currentEntry("displayName") = baseName
                currentEntry("notes") = ""
                currentEntry(modABPhotoConstants.PHOTO_TAG_BERICHT) = ""
                currentEntry(modABPhotoConstants.PHOTO_TAG_SEMINAR) = ""
                currentEntry(modABPhotoConstants.PHOTO_TAG_TOPIC) = ""
                currentEntry("preferredLocale") = ""
                currentEntry("capturedAt") = imageItem("capturedAt")
            Else
                currentEntry.CompareMode = TextCompare
                If Len(NzString(currentEntry("displayName"))) = 0 Then currentEntry("displayName") = baseName
                currentEntry(modABPhotoConstants.PHOTO_TAG_BERICHT) = ""
                currentEntry(modABPhotoConstants.PHOTO_TAG_SEMINAR) = ""
                currentEntry(modABPhotoConstants.PHOTO_TAG_TOPIC) = ""
            End If
            Set entryMap(baseName) = currentEntry
        End If

        relativePath = NzString(imageItem("relativePath"))
        If Len(relativePath) = 0 Then relativePath = baseName

        If Len(NzString(currentEntry("capturedAt"))) = 0 Then
            currentEntry("capturedAt") = imageItem("capturedAt")
        End If

        modABPhotosRepository.ApplyFolderTags currentEntry, relativePath, folderTagMap
ContinueLoop:
    Next imageItem

    Dim key As Variant
    For Each key In entryMap.Keys
        Set currentEntry = entryMap(key)
        currentEntry.CompareMode = TextCompare
        If Len(NzString(currentEntry("displayName"))) = 0 Then
            currentEntry("displayName") = currentEntry("fileName")
        End If
        modABPhotosRepository.UpsertPhoto currentEntry
    Next key

    modABPhotosRepository.RemoveMissingPhotos entryMap
End Sub

Public Sub CreateFoldersForList(ByVal baseDirectory As String, ByVal listName As String, ByVal locale As String)
    If Len(baseDirectory) = 0 Then Exit Sub
    Dim buttonEntries As Collection
    Set buttonEntries = GetButtonList(listName, locale)
    If buttonEntries Is Nothing Then Exit Sub
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim btnEntry As Scripting.Dictionary
    For Each btnEntry In buttonEntries
        Dim folderName As String
        folderName = NzString(btnEntry("label"))
        If Len(folderName) = 0 Then folderName = NzString(btnEntry("value"))
        If Len(folderName) = 0 Then GoTo ContinueLoop
        Dim targetPath As String
        targetPath = BuildPath(baseDirectory, folderName)
        If Not fso.FolderExists(targetPath) Then
            fso.CreateFolder targetPath
        End If
ContinueLoop:
    Next btnEntry
    MsgBox "Ordnerstruktur aktualisiert.", vbInformation
End Sub

Public Sub RemoveEmptyFolders(ByVal baseDirectory As String)
    If Len(baseDirectory) = 0 Then Exit Sub
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(baseDirectory) Then Exit Sub
    Dim rootFld As Object
    Set rootFld = fso.GetFolder(baseDirectory)

    Dim deleteList As Collection
    Set deleteList = New Collection

    Dim subFld As Object
    For Each subFld In rootFld.SubFolders
        If subFld.Files.Count = 0 And subFld.SubFolders.Count = 0 Then
            deleteList.Add subFld.Path
        End If
    Next subFld

    Dim target As Variant
    For Each target In deleteList
        fso.DeleteFolder CStr(target), True
    Next target
    MsgBox "Leere Ordner entfernt.", vbInformation
End Sub

Private Function ParseTagValues(ByVal csv As String) As Collection
    Dim result As New Collection
    Dim textValue As String
    textValue = Trim$(NzString(csv))
    If Len(textValue) = 0 Then
        Set ParseTagValues = result
        Exit Function
    End If

    Dim parts() As String
    parts = Split(textValue, ",")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim token As String
        token = Trim$(parts(i))
        If Len(token) > 0 Then result.Add token
    Next i
    Set ParseTagValues = result
End Function

Private Function CollectionToSortedArray(items As Collection) As Variant
    If items Is Nothing Or items.Count = 0 Then
        CollectionToSortedArray = Array()
        Exit Function
    End If
    Dim arr() As String
    ReDim arr(1 To items.Count)
    Dim i As Long
    For i = 1 To items.Count
        arr(i) = CStr(items(i))
    Next i
    Dim j As Long
    Dim tmp As String
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If StrComp(arr(i), arr(j), vbTextCompare) > 0 Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i
    CollectionToSortedArray = arr
End Function

Public Function ResolvePhotoPath(ByVal baseDirectory As String, ByVal record As Scripting.Dictionary, ByVal locale As String) As String
    If record Is Nothing Then Exit Function
    If Len(baseDirectory) = 0 Then Exit Function

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim map As Scripting.Dictionary
    Set map = BuildDesiredPathMap(baseDirectory, record, locale)

    Dim baseName As String
    baseName = NzString(record("fileName"))
    Dim preferredRoot As String
    preferredRoot = BuildPath(baseDirectory, baseName)

    If map Is Nothing Then
        If Len(baseName) > 0 And fso.FileExists(preferredRoot) Then ResolvePhotoPath = preferredRoot
        Exit Function
    End If

    Dim canonical As String
    canonical = ""
    If Len(baseName) > 0 And map.Exists(preferredRoot) Then
        canonical = preferredRoot
    Else
        Dim pathKey As Variant
        For Each pathKey In map.Keys
            canonical = CStr(pathKey)
            Exit For
        Next pathKey
    End If

    If Len(canonical) > 0 And fso.FileExists(canonical) Then
        ResolvePhotoPath = canonical
        Exit Function
    End If

    Dim candidate As String
    Dim pathKey As Variant
    For Each pathKey In map.Keys
        candidate = CStr(pathKey)
        If fso.FileExists(candidate) Then
            ResolvePhotoPath = candidate
            Exit Function
        End If
    Next pathKey

    If Len(baseName) > 0 And fso.FileExists(preferredRoot) Then
        ResolvePhotoPath = preferredRoot
    End If
End Function

Private Function GetFolderLabelForTag(ByVal tagField As String, ByVal tagValue As String, ByVal locale As String) As String
    Dim listName As String
    Select Case tagField
        Case modABPhotoConstants.PHOTO_TAG_BERICHT
            listName = modABPhotoConstants.PHOTO_LIST_BERICHT
        Case modABPhotoConstants.PHOTO_TAG_SEMINAR
            listName = modABPhotoConstants.PHOTO_LIST_SEMINAR
        Case modABPhotoConstants.PHOTO_TAG_TOPIC
            listName = modABPhotoConstants.PHOTO_LIST_TOPIC
        Case Else
            GetFolderLabelForTag = tagValue
            Exit Function
    End Select

    Dim buttons As Collection
    Set buttons = GetButtonList(listName, locale)
    If Not buttons Is Nothing Then
        Dim buttonEntry As Scripting.Dictionary
        For Each buttonEntry In buttons
            If StrComp(NzString(buttonEntry("value")), tagValue, vbTextCompare) = 0 Then
                Dim labelValue As String
                labelValue = NzString(buttonEntry("label"))
                If Len(labelValue) = 0 Then labelValue = NzString(buttonEntry("value"))
                GetFolderLabelForTag = labelValue
                Exit Function
            End If
        Next buttonEntry
    End If

    GetFolderLabelForTag = tagValue
End Function

Public Function BuildDesiredPathMap(ByVal baseDirectory As String, ByVal record As Scripting.Dictionary, ByVal locale As String) As Scripting.Dictionary
    Dim result As New Scripting.Dictionary
    result.CompareMode = TextCompare

    If record Is Nothing Then
        Set BuildDesiredPathMap = result
        Exit Function
    End If

    Dim baseName As String
    baseName = NzString(record("fileName"))
    If Len(baseName) = 0 Then
        Set BuildDesiredPathMap = result
        Exit Function
    End If

    Dim fields As Variant
    fields = Array(modABPhotoConstants.PHOTO_TAG_BERICHT, modABPhotoConstants.PHOTO_TAG_SEMINAR, modABPhotoConstants.PHOTO_TAG_TOPIC)

    Dim anyTag As Boolean
    Dim field As Variant
    For Each field In fields
        Dim tags As Collection
        Set tags = ParseTagValues(NzString(record(CStr(field))))
        Dim sortedValues As Variant
        sortedValues = CollectionToSortedArray(tags)
        Dim idx As Long
        For idx = LBound(sortedValues) To UBound(sortedValues)
            Dim folderLabel As String
            folderLabel = GetFolderLabelForTag(CStr(field), CStr(sortedValues(idx)), locale)
            folderLabel = modABPhotosRepository.NormalizeFolderName(folderLabel)
            If Len(folderLabel) = 0 Then folderLabel = CStr(sortedValues(idx))
            Dim relativePath As String
            relativePath = folderLabel & "\" & baseName
            Dim absolutePath As String
            absolutePath = BuildPath(baseDirectory, relativePath)
            If Not result.Exists(absolutePath) Then
                result.Add absolutePath, relativePath
            End If
            anyTag = True
        Next idx
    Next field

    If Not anyTag Then
        Dim relativeRoot As String
        relativeRoot = baseName
        result(BuildPath(baseDirectory, relativeRoot)) = relativeRoot
    End If

    Set BuildDesiredPathMap = result
End Function

Private Sub EnsureDirectoryExists(fso As Object, ByVal folderPath As String)
    If Len(folderPath) = 0 Then Exit Sub
    If fso.FolderExists(folderPath) Then Exit Sub
    Dim parent As String
    parent = fso.GetParentFolderName(folderPath)
    If Len(parent) > 0 Then
        EnsureDirectoryExists fso, parent
    End If
    On Error Resume Next
    fso.CreateFolder folderPath
    On Error GoTo 0
End Sub

Public Sub SyncPhotoFiles(ByVal baseDirectory As String, oldRecord As Scripting.Dictionary, newRecord As Scripting.Dictionary, ByVal locale As String)
    If Len(baseDirectory) = 0 Then Exit Sub
    If newRecord Is Nothing Then Exit Sub

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim oldMap As Scripting.Dictionary
    If oldRecord Is Nothing Then
        Set oldMap = New Scripting.Dictionary
        oldMap.CompareMode = TextCompare
    Else
        Set oldMap = BuildDesiredPathMap(baseDirectory, oldRecord, locale)
    End If

    Dim newMap As Scripting.Dictionary
    Set newMap = BuildDesiredPathMap(baseDirectory, newRecord, locale)
    If newMap.Count = 0 Then Exit Sub

    Dim baseName As String
    baseName = NzString(newRecord("fileName"))
    Dim pathKey As Variant

    Dim canonicalAbsolute As String
    canonicalAbsolute = ""

    Dim preferredRoot As String
    preferredRoot = BuildPath(baseDirectory, baseName)
    If Len(baseName) > 0 And newMap.Exists(preferredRoot) Then
        canonicalAbsolute = preferredRoot
    Else
        For Each pathKey In newMap.Keys
            canonicalAbsolute = CStr(pathKey)
            Exit For
        Next pathKey
    End If

    If Len(canonicalAbsolute) = 0 Then Exit Sub

    Dim sourcePath As String
    sourcePath = ""

    For Each pathKey In oldMap.Keys
        If fso.FileExists(CStr(pathKey)) Then
            sourcePath = CStr(pathKey)
            Exit For
        End If
    Next pathKey

    If Len(sourcePath) = 0 Then
        For Each pathKey In newMap.Keys
            If fso.FileExists(CStr(pathKey)) Then
                sourcePath = CStr(pathKey)
                Exit For
            End If
        Next pathKey
    End If

    If Len(sourcePath) = 0 Then
        Dim rootCandidate As String
        rootCandidate = BuildPath(baseDirectory, baseName)
        If fso.FileExists(rootCandidate) Then sourcePath = rootCandidate
    End If

    Dim canonicalFolder As String
    canonicalFolder = fso.GetParentFolderName(canonicalAbsolute)
    If Len(canonicalFolder) > 0 Then EnsureDirectoryExists fso, canonicalFolder

    If Len(sourcePath) > 0 Then
        If StrComp(sourcePath, canonicalAbsolute, vbTextCompare) <> 0 Then
            If fso.FileExists(canonicalAbsolute) Then
                On Error Resume Next
                fso.DeleteFile canonicalAbsolute, True
                On Error GoTo 0
            End If
            On Error Resume Next
            fso.MoveFile sourcePath, canonicalAbsolute
            If Err.Number <> 0 Then
                Err.Clear
                fso.CopyFile sourcePath, canonicalAbsolute, True
                fso.DeleteFile sourcePath, True
            End If
            On Error GoTo 0
        End If
    End If

    If Not fso.FileExists(canonicalAbsolute) Then Exit Sub

    For Each pathKey In newMap.Keys
        Dim targetAbsolute As String
        targetAbsolute = CStr(pathKey)
        If StrComp(targetAbsolute, canonicalAbsolute, vbTextCompare) <> 0 Then
            Dim targetFolder As String
            targetFolder = fso.GetParentFolderName(targetAbsolute)
            If Len(targetFolder) > 0 Then EnsureDirectoryExists fso, targetFolder
            If Not fso.FileExists(targetAbsolute) Then
                On Error Resume Next
                fso.CopyFile canonicalAbsolute, targetAbsolute, True
                On Error GoTo 0
            End If
        End If
    Next pathKey

    For Each pathKey In oldMap.Keys
        Dim oldAbsolute As String
        oldAbsolute = CStr(pathKey)
        If Not newMap.Exists(oldAbsolute) Then
            If fso.FileExists(oldAbsolute) Then
                On Error Resume Next
                fso.DeleteFile oldAbsolute, True
                On Error GoTo 0
            End If
        End If
    Next pathKey

    modABPhotosRepository.UpsertPhoto newRecord
End Sub

Private Function EnumerateImages(ByVal baseDirectory As String) As Collection
    Dim results As New Collection
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(baseDirectory) Then
        Set EnumerateImages = results
        Exit Function
    End If
    Dim rootFolder As Object
    Set rootFolder = fso.GetFolder(baseDirectory)
    TraverseFolder rootFolder, baseDirectory, results
    Set EnumerateImages = results
End Function

Private Sub TraverseFolder(ByVal folder As Object, ByVal baseDirectory As String, ByRef results As Collection)
    Dim file As Object
    For Each file In folder.Files
        If IsImageFile(file.Name) Then
            Dim relativePath As String
            relativePath = Mid$(file.Path, Len(baseDirectory) + 2)
            Dim item As New Scripting.Dictionary
            item.CompareMode = TextCompare
            item("fullPath") = file.Path
            item("relativePath") = relativePath
            item("baseName") = file.Name
            item("capturedAt") = file.DateCreated
            results.Add item
        End If
    Next file
    Dim subFolder As Object
    For Each subFolder In folder.SubFolders
        TraverseFolder subFolder, baseDirectory, results
    Next subFolder
End Sub

Private Function IsImageFile(fileName As String) As Boolean
    Dim lowered As String
    lowered = LCase$(fileName)
    IsImageFile = (Right$(lowered, 4) = ".jpg") _
        Or (Right$(lowered, 5) = ".jpeg") _
        Or (Right$(lowered, 4) = ".png") _
        Or (Right$(lowered, 4) = ".bmp")
End Function

Private Function BuildPath(ByVal baseDirectory As String, ByVal segment As String) As String
    If Len(baseDirectory) = 0 Then
        BuildPath = segment
    ElseIf Right$(baseDirectory, 1) = "\" Or Right$(baseDirectory, 1) = "/" Then
        BuildPath = baseDirectory & segment
    Else
        BuildPath = baseDirectory & "\" & segment
    End If
End Function
