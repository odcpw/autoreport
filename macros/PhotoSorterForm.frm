VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PhotoSorterForm 
   Caption         =   "PhotoSorter"
   ClientHeight    =   12255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   23055
   OleObjectBlob   =   "PhotoSorterForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PhotoSorterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Private Const FILTER_ALL As String = "Show All"
Private Const FILTER_UNSORTED As String = "Show Only Unsorted"
Private Const FILTER_SORTED As String = "Show Only Sorted"

Private Const LIST_BERICHT As String = "photo.berichtsbuttons"
Private Const LIST_AUDIT As String = "photo.auditbuttons"
Private Const LIST_TRAINING As String = "photo.trainingbuttons"
Private Const LIST_SUBFOLDERS As String = "photo.subfolders"

Private Const TAG_FIELD_CHAPTERS As String = "tagChapters"
Private Const TAG_FIELD_CATEGORIES As String = "tagCategories"
Private Const TAG_FIELD_TRAINING As String = "tagTraining"
Private Const TAG_FIELD_SUBFOLDERS As String = "tagSubfolders"

Private imageFiles As Collection
Private photoCache As Scripting.Dictionary
Private buttonCollection As Collection
Private localeCache As String
Private currentIndex As Long
Private isLoaded As Boolean

Private Sub UserForm_Initialize()
    Set imageFiles = New Collection
    Set photoCache = New Scripting.Dictionary
    photoCache.CompareMode = TextCompare
    Set buttonCollection = New Collection
    InitializeFilterOptions
    LoadTagButtons
    RefreshPhotoList FILTER_ALL
    isLoaded = True
End Sub

Private Sub InitializeFilterOptions()
    cmbFilter.Clear
    cmbFilter.AddItem FILTER_ALL
    cmbFilter.AddItem FILTER_UNSORTED
    cmbFilter.AddItem FILTER_SORTED
    cmbFilter.ListIndex = 0
End Sub

Public Sub RefreshPhotoList(Optional ByVal preferredFilter As String = "")
    LoadPhotoCache
    Dim activeFilter As String
    If Len(preferredFilter) > 0 Then
        activeFilter = preferredFilter
    Else
        activeFilter = cmbFilter.Value
    End If
    BuildFilteredPhotoList activeFilter
    If imageFiles.Count = 0 Then
        currentIndex = 0
    ElseIf currentIndex = 0 Or currentIndex > imageFiles.Count Then
        currentIndex = 1
    End If
    UpdateImageDisplay
    UpdateImageCounts
End Sub

Private Sub LoadPhotoCache()
    Set photoCache = New Scripting.Dictionary
    photoCache.CompareMode = TextCompare
    Dim data As Collection
    Set data = ReadTableAsCollection(PhotosSheet)
    Dim entry As Scripting.Dictionary
    For Each entry In data
        Dim fileKey As String
        fileKey = NzString(entry("fileName"))
        If Len(fileKey) > 0 Then
            photoCache(fileKey) = entry
        End If
    Next entry
End Sub

Private Sub BuildFilteredPhotoList(ByVal filterOption As String)
    Set imageFiles = New Collection
    Dim key As Variant
    For Each key In photoCache.Keys
        If ShouldIncludeRecord(photoCache(key), filterOption) Then
            imageFiles.Add CStr(key)
        End If
    Next key
    SortImageFiles
    cmbFilter.Value = filterOption
End Sub

Private Function ShouldIncludeRecord(ByVal record As Scripting.Dictionary, ByVal filterOption As String) As Boolean
    Select Case filterOption
        Case FILTER_UNSORTED
            ShouldIncludeRecord = Not RecordHasAnyTag(record)
        Case FILTER_SORTED
            ShouldIncludeRecord = RecordHasAnyTag(record)
        Case Else
            ShouldIncludeRecord = True
    End Select
End Function

Private Function RecordHasAnyTag(ByVal record As Scripting.Dictionary) As Boolean
    RecordHasAnyTag = (Len(NzString(record(TAG_FIELD_CHAPTERS))) > 0) _
        Or (Len(NzString(record(TAG_FIELD_CATEGORIES))) > 0) _
        Or (Len(NzString(record(TAG_FIELD_TRAINING))) > 0) _
        Or (Len(NzString(record(TAG_FIELD_SUBFOLDERS))) > 0)
End Function

Private Sub SortImageFiles()
    If imageFiles.Count <= 1 Then Exit Sub
    Dim temp() As String
    ReDim temp(1 To imageFiles.Count)
    Dim i As Long
    For i = 1 To imageFiles.Count
        temp(i) = imageFiles(i)
    Next i
    Dim swapped As Boolean
    Do
        swapped = False
        For i = LBound(temp) To UBound(temp) - 1
            If StrComp(temp(i), temp(i + 1), vbTextCompare) > 0 Then
                Dim hold As String
                hold = temp(i)
                temp(i) = temp(i + 1)
                temp(i + 1) = hold
                swapped = True
            End If
        Next i
    Loop While swapped
    Set imageFiles = New Collection
    For i = LBound(temp) To UBound(temp)
        imageFiles.Add temp(i)
    Next i
End Sub

Private Sub LoadTagButtons()
    ClearTagButtonFrame ButtonsBericht
    ClearTagButtonFrame ButtonsVGSeminar
    ClearTagButtonFrame ButtonsSubfolders
    Set buttonCollection = New Collection
    CreateTagButtons ButtonsBericht, LIST_BERICHT, TAG_FIELD_CHAPTERS
    CreateTagButtons ButtonsVGSeminar, LIST_AUDIT, TAG_FIELD_CATEGORIES
    CreateTagButtons ButtonsSubfolders, LIST_SUBFOLDERS, TAG_FIELD_SUBFOLDERS
End Sub

Private Sub ClearTagButtonFrame(ByVal targetFrame As MSForms.Frame)
    Dim index As Long
    For index = targetFrame.Controls.Count - 1 To 0 Step -1
        If TypeName(targetFrame.Controls(index)) = "CommandButton" Then
            targetFrame.Controls.Remove targetFrame.Controls(index).Name
        End If
    Next index
End Sub

Private Sub CreateTagButtons(ByVal targetFrame As MSForms.Frame, ByVal listName As String, ByVal tagField As String)
    Dim buttons As Collection
    Set buttons = GetButtonList(listName, GetActiveLocale())
    If buttons Is Nothing Then Exit Sub
    If buttons.Count = 0 Then Exit Sub

    Const BTN_WIDTH As Integer = 110
    Const BTN_HEIGHT As Integer = 22
    Const BTN_SPACING As Integer = 6
    Const BTN_COLUMNS As Integer = 4

    Dim entry As Scripting.Dictionary
    Dim i As Long
    For i = 1 To buttons.Count
        Set entry = buttons(i)
        Dim caption As String
        caption = NzString(entry("label"))
        Dim value As String
        value = NzString(entry("value"))
        If Len(value) = 0 Then value = caption
        Dim cmd As MSForms.CommandButton
        Set cmd = targetFrame.Controls.Add("Forms.CommandButton.1")
        cmd.Name = listName & "_" & i
        cmd.Caption = caption
        cmd.Width = BTN_WIDTH
        cmd.Height = BTN_HEIGHT
        Dim colIndex As Integer
        Dim rowIndex As Integer
        colIndex = (i - 1) Mod BTN_COLUMNS
        rowIndex = (i - 1) \ BTN_COLUMNS
        cmd.Left = 6 + colIndex * (BTN_WIDTH + BTN_SPACING)
        cmd.Top = 6 + rowIndex * (BTN_HEIGHT + BTN_SPACING)
        Dim wrapper As CPhotoTagButton
        Set wrapper = New CPhotoTagButton
        wrapper.Initialize Me, cmd, tagField, value, listName, caption
        buttonCollection.Add wrapper
    Next i
    targetFrame.ScrollBars = fmScrollBarsVertical
    targetFrame.ScrollHeight = (buttons.Count \ BTN_COLUMNS + 1) * (BTN_HEIGHT + BTN_SPACING) + BTN_SPACING
End Sub

Private Function GetActiveLocale() As String
    If Len(localeCache) = 0 Then
        On Error GoTo fallback
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(SHEET_META)
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim r As Long
        For r = ROW_HEADER_ROW + 1 To lastRow
            If StrComp(CStr(ws.Cells(r, 1).Value), "locale", vbTextCompare) = 0 Then
                localeCache = NzString(ws.Cells(r, 2).Value)
                Exit For
            End If
        Next r
fallback:
        If Len(localeCache) = 0 Then localeCache = DefaultLocale()
    End If
    GetActiveLocale = localeCache
End Function

Private Sub cmbFilter_Change()
    If Not isLoaded Then Exit Sub
    RefreshPhotoList cmbFilter.Value
    InvisibleTextBox.SetFocus
End Sub

Private Sub btnPrevious_Click()
    If currentIndex > 1 Then
        currentIndex = currentIndex - 1
        UpdateImageDisplay
        UpdateImageCounts
    End If
    InvisibleTextBox.SetFocus
End Sub

Private Sub btnNext_Click()
    If currentIndex < imageFiles.Count Then
        currentIndex = currentIndex + 1
        UpdateImageDisplay
        UpdateImageCounts
    End If
    InvisibleTextBox.SetFocus
End Sub

Private Sub InvisibleTextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case vbKeyA
            btnPrevious_Click
        Case vbKeyD
            btnNext_Click
    End Select
End Sub

Private Sub btnShowCounts_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    ShowCategoryCounts
End Sub

Private Sub btnShowCounts_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    RestoreCategoryCaptions
    UpdateButtonStates GetCurrentPhotoRecord
    InvisibleTextBox.SetFocus
End Sub

Private Sub btnClearSheets_Click()
    If currentIndex = 0 Then Exit Sub
    Dim fileName As String
    fileName = imageFiles(currentIndex)
    SetPhotoTags fileName, TAG_FIELD_CHAPTERS, Array()
    SetPhotoTags fileName, TAG_FIELD_CATEGORIES, Array()
    SetPhotoTags fileName, TAG_FIELD_TRAINING, Array()
    SetPhotoTags fileName, TAG_FIELD_SUBFOLDERS, Array()
    photoCache(fileName) = GetPhotoEntry(fileName)
    UpdateButtonStates photoCache(fileName)
    UpdateImageCounts
End Sub

Private Sub cmdCreateAllFolders_Click()
    If Len(rootPath) = 0 Then
        MsgBox "Bitte zuerst einen Wurzelordner auswählen.", vbExclamation
        Exit Sub
    End If
    CreateFoldersForList rootPath, LIST_SUBFOLDERS, GetActiveLocale()
End Sub

Private Sub cmdRemoveEmptyFolders_Click()
    If Len(rootPath) = 0 Then
        MsgBox "Bitte zuerst einen Wurzelordner auswählen.", vbExclamation
        Exit Sub
    End If
    RemoveEmptyFolders rootPath
End Sub

Private Sub cmdChooseDirectory_Click()
    Dim selectedPath As String
    selectedPath = ChooseImageDirectory()
    If Len(selectedPath) = 0 Then
        lblDirectoryPath.Caption = "No directory selected"
        Exit Sub
    End If
    rootPath = selectedPath
    lblDirectoryPath.Caption = rootPath
    ScanImagesIntoSheet rootPath
    RefreshPhotoList cmbFilter.Value
    InvisibleTextBox.SetFocus
End Sub

Private Sub UpdateImageDisplay()
    If currentIndex <= 0 Or currentIndex > imageFiles.Count Then
        ImageControl.Picture = Nothing
        lblCurrentImageName.Caption = ""
        UpdateButtonStates Nothing
        Exit Sub
    End If

    Dim fileName As String
    fileName = imageFiles(currentIndex)
    Dim record As Scripting.Dictionary
    Set record = GetPhotoEntry(fileName)
    If record Is Nothing Then
        ImageControl.Picture = Nothing
        lblCurrentImageName.Caption = fileName
        UpdateButtonStates Nothing
        Exit Sub
    End If

    photoCache(fileName) = record
    Dim displayName As String
    displayName = NzString(record("displayName"))
    If Len(displayName) = 0 Then displayName = fileName
    lblCurrentImageName.Caption = displayName

    Dim fullPath As String
    fullPath = BuildPhotoPath(fileName)
    If Len(fullPath) > 0 And Dir(fullPath) <> "" Then
        ImageControl.Picture = LoadPicture(fullPath)
        ImageControl.PictureSizeMode = fmPictureSizeModeZoom
    Else
        ImageControl.Picture = Nothing
    End If
    UpdateButtonStates record
End Sub

Private Function BuildPhotoPath(ByVal fileName As String) As String
    If Len(fileName) = 0 Then Exit Function
    If InStr(fileName, ":") > 0 Or Left$(fileName, 2) = "\" Then
        BuildPhotoPath = fileName
    ElseIf Len(rootPath) > 0 Then
        If Right$(rootPath, 1) = "" Or Right$(rootPath, 1) = "/" Then
            BuildPhotoPath = rootPath & fileName
        Else
            BuildPhotoPath = rootPath & "" & fileName
        End If
    Else
        BuildPhotoPath = fileName
    End If
End Function

Private Sub UpdateButtonStates(ByVal record As Scripting.Dictionary)
    Dim buttonObj As CPhotoTagButton
    For Each buttonObj In buttonCollection
        Dim isActive As Boolean
        If record Is Nothing Then
            isActive = False
        Else
            isActive = TagExistsInRecord(record, buttonObj.TagField, buttonObj.TagValue)
        End If
        buttonObj.SetActive isActive
    Next buttonObj
End Sub

Private Function TagExistsInRecord(ByVal record As Scripting.Dictionary, ByVal tagField As String, ByVal tagValue As String) As Boolean
    Dim values As Collection
    Set values = ParseTagCollection(NzString(record(tagField)))
    Dim item As Variant
    For Each item In values
        If StrComp(CStr(item), tagValue, vbTextCompare) = 0 Then
            TagExistsInRecord = True
            Exit Function
        End If
    Next item
    TagExistsInRecord = False
End Function

Public Sub ToggleTag(ByVal tagField As String, ByVal tagValue As String)
    If currentIndex = 0 Then Exit Sub
    Dim fileName As String
    fileName = imageFiles(currentIndex)
    TogglePhotoTag fileName, tagField, tagValue
    photoCache(fileName) = GetPhotoEntry(fileName)
    UpdateButtonStates photoCache(fileName)
    UpdateImageCounts
End Sub

Private Sub ShowCategoryCounts()
    Dim counts As Scripting.Dictionary
    Set counts = ComputeTagCounts()
    Dim buttonObj As CPhotoTagButton
    For Each buttonObj In buttonCollection
        Dim key As String
        key = buttonObj.TagField & "|" & buttonObj.TagValue
        Dim countValue As Long
        If counts.Exists(key) Then countValue = counts(key) Else countValue = 0
        buttonObj.ShowCount countValue
    Next buttonObj
End Sub

Private Sub RestoreCategoryCaptions()
    Dim buttonObj As CPhotoTagButton
    For Each buttonObj In buttonCollection
        buttonObj.ResetCaption
    Next buttonObj
End Sub

Private Function ComputeTagCounts() As Scripting.Dictionary
    Dim counts As New Scripting.Dictionary
    counts.CompareMode = TextCompare
    Dim key As Variant
    For Each key In photoCache.Keys
        Dim record As Scripting.Dictionary
        Set record = photoCache(key)
        AddCounts counts, TAG_FIELD_CHAPTERS, NzString(record(TAG_FIELD_CHAPTERS))
        AddCounts counts, TAG_FIELD_CATEGORIES, NzString(record(TAG_FIELD_CATEGORIES))
        AddCounts counts, TAG_FIELD_TRAINING, NzString(record(TAG_FIELD_TRAINING))
        AddCounts counts, TAG_FIELD_SUBFOLDERS, NzString(record(TAG_FIELD_SUBFOLDERS))
    Next key
    Set ComputeTagCounts = counts
End Function

Private Sub AddCounts(ByRef counts As Scripting.Dictionary, ByVal tagField As String, ByVal csv As String)
    Dim items As Collection
    Set items = ParseTagCollection(csv)
    Dim item As Variant
    For Each item In items
        Dim key As String
        key = tagField & "|" & CStr(item)
        If counts.Exists(key) Then
            counts(key) = counts(key) + 1
        Else
            counts.Add key, 1
        End If
    Next item
End Sub

Private Function ParseTagCollection(ByVal csv As String) As Collection
    Dim result As New Collection
    If Len(csv) = 0 Then
        Set ParseTagCollection = result
        Exit Function
    End If
    Dim parts() As String
    parts = Split(csv, ",")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim value As String
        value = Trim$(parts(i))
        If Len(value) > 0 Then result.Add value
    Next i
    Set ParseTagCollection = result
End Function

Private Function GetCurrentPhotoRecord() As Scripting.Dictionary
    If currentIndex = 0 Or currentIndex > imageFiles.Count Then
        Set GetCurrentPhotoRecord = Nothing
    Else
        Dim fileName As String
        fileName = imageFiles(currentIndex)
        If photoCache.Exists(fileName) Then
            Set GetCurrentPhotoRecord = photoCache(fileName)
        Else
            Set GetCurrentPhotoRecord = GetPhotoEntry(fileName)
        End If
    End If
End Function

Private Sub UpdateImageCounts()
    Dim totalPhotos As Long
    totalPhotos = photoCache.Count
    Dim unsorted As Long
    unsorted = CountUnsortedPhotos()
    Dim current As Long
    If imageFiles.Count = 0 Then
        current = 0
    Else
        current = currentIndex
    End If
    lblImageCounts.Caption = "Image " & current & " of " & imageFiles.Count & " — Total Photos " & totalPhotos & " — Unsorted Photos " & unsorted
End Sub

Private Function CountUnsortedPhotos() As Long
    Dim key As Variant
    For Each key In photoCache.Keys
        If Not RecordHasAnyTag(photoCache(key)) Then
            CountUnsortedPhotos = CountUnsortedPhotos + 1
        End If
    Next key
End Function

Private Sub FilterAndReloadImages()
    RefreshPhotoList cmbFilter.Value
End Sub
