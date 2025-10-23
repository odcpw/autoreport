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
    Dim item As Variant
    For Each item In images
        Dim relativePath As String
        relativePath = item("relativePath")
        Dim record As Scripting.Dictionary
        Set record = GetPhotoEntry(relativePath)
        If record Is Nothing Then
            Set record = New Scripting.Dictionary
            record.CompareMode = TextCompare
        Else
            record.CompareMode = TextCompare
        End If
        Dim headers As Variant
        headers = HeaderPhotos()
        Dim header As Variant
        For Each header In headers
            Select Case header
                Case "fileName"
                    record(header) = relativePath
                Case "displayName"
                    If Not record.Exists(header) Or Len(NzString(record(header))) = 0 Then
                        record(header) = relativePath
                    End If
                Case "capturedAt"
                    record(header) = item("capturedAt")
                Case Else
                    If Not record.Exists(header) Then
                        record(header) = ""
                    End If
            End Select
        Next header
        UpsertPhoto record
    Next item
End Sub

Public Sub CreateFoldersForList(ByVal baseDirectory As String, ByVal listName As String, ByVal locale As String)
    If Len(baseDirectory) = 0 Then Exit Sub
    Dim entries As Collection
    Set entries = GetButtonList(listName, locale)
    If entries Is Nothing Then Exit Sub
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim entry As Scripting.Dictionary
    For Each entry In entries
        Dim folderName As String
        folderName = NzString(entry("value"))
        If Len(folderName) = 0 Then folderName = NzString(entry("label"))
        If Len(folderName) = 0 Then GoTo ContinueLoop
        Dim targetPath As String
        targetPath = BuildPath(baseDirectory, folderName)
        If Not fso.FolderExists(targetPath) Then
            fso.CreateFolder targetPath
        End If
ContinueLoop:
    Next entry
    MsgBox "Ordnerstruktur aktualisiert.", vbInformation
End Sub

Public Sub RemoveEmptyFolders(ByVal baseDirectory As String)
    If Len(baseDirectory) = 0 Then Exit Sub
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(baseDirectory) Then Exit Sub
    Dim rootFld As Object
    Set rootFld = fso.GetFolder(baseDirectory)
    Dim i As Long
    For i = rootFld.SubFolders.Count To 1 Step -1
        Dim subFld As Object
        Set subFld = rootFld.SubFolders(i)
        If subFld.Files.Count = 0 And subFld.SubFolders.Count = 0 Then
            fso.DeleteFolder subFld.Path, True
        End If
    Next i
    MsgBox "Leere Ordner entfernt.", vbInformation
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
    If Right$(baseDirectory, 1) = "" Or Right$(baseDirectory, 1) = "/" Then
        BuildPath = baseDirectory & segment
    Else
        BuildPath = baseDirectory & "" & segment
    End If
End Function
