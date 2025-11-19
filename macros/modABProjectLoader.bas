Attribute VB_Name = "modABProjectLoader"
Option Explicit

'=============================================================
' Load project.json into structured sheets
'=============================================================

Public Sub LoadProjectJson(ByVal filePath As String)
    Dim jsonText As String
    jsonText = ReadTextFile(filePath)

    Dim project As Dictionary
    Set project = JsonConverter.ParseJson(jsonText)

    PopulateProjectTables project
    MsgBox "project.json loaded", vbInformation
End Sub

Public Sub PopulateProjectTables(project As Dictionary)
    EnsureAutoBerichtSheets True

    WriteMeta GetDict(project, "meta")
    WriteLists GetDict(project, "lists")
    WritePhotos GetDict(project, "photos")
    WriteChaptersAndRows project
    WriteHistory GetDict(project, "history")
End Sub

Private Sub WriteMeta(meta As Variant)
    Dim ws As Worksheet
    Set ws = EnsureSheetWithHeaders(SHEET_META, HeaderMeta(), False)
    ws.Cells.ClearContents
    WriteHeaderRow ws, HeaderMeta()

    If TypeName(meta) <> "Dictionary" Then Exit Sub
    Dim rowIndex As Long
    rowIndex = ROW_HEADER_ROW + 1
    Dim key As Variant
    For Each key In meta.Keys
        ws.Cells(rowIndex, 1).Value = CStr(key)
        ws.Cells(rowIndex, 2).Value = meta(key)
        rowIndex = rowIndex + 1
    Next key
End Sub

Private Sub WriteChaptersAndRows(project As Dictionary)
    Dim chapters As Variant
    chapters = GetDict(project, "chapters")

    Dim wsChapters As Worksheet
    Set wsChapters = EnsureSheetWithHeaders(SHEET_CHAPTERS, HeaderChapters(), False)
    wsChapters.Cells.ClearContents
    WriteHeaderRow wsChapters, HeaderChapters()

    Dim wsRows As Worksheet
    Set wsRows = EnsureSheetWithHeaders(SHEET_ROWS, HeaderRows(), False)
    wsRows.Cells.ClearContents
    WriteHeaderRow wsRows, HeaderRows()

    If TypeName(chapters) <> "Collection" Then Exit Sub

    Dim chapterQueue As New Collection
    Dim item As Variant
    For Each item In chapters
        chapterQueue.Add item
    Next item

    Dim chapterPosition As Long
    chapterPosition = ROW_HEADER_ROW + 1
    Dim rowPosition As Long
    rowPosition = ROW_HEADER_ROW + 1

    Do While chapterQueue.Count > 0
        Dim chapter As Dictionary
        Set chapter = chapterQueue(1)
        chapterQueue.Remove 1

        Dim chapterId As String
        chapterId = NzString(GetDictValue(chapter, "id", ""))
        If Len(chapterId) = 0 Then GoTo NextChapter

        Dim parentId As String
        parentId = NzString(GetDictValue(chapter, "parentId", GetParentChapterId(chapterId)))

        wsChapters.Cells(chapterPosition, HeaderIndex(wsChapters, "chapterId")).Value = chapterId
        wsChapters.Cells(chapterPosition, HeaderIndex(wsChapters, "parentId")).Value = parentId
        wsChapters.Cells(chapterPosition, HeaderIndex(wsChapters, "orderIndex")).Value = NzNumber(GetDictValue(chapter, "orderIndex", chapterPosition - ROW_HEADER_ROW))
        WriteChapterTitles wsChapters, chapterPosition, GetDictValue(chapter, "title", Nothing)
        wsChapters.Cells(chapterPosition, HeaderIndex(wsChapters, "pageSize")).Value = NzNumber(GetDictValue(chapter, "pageSize", 5))
        wsChapters.Cells(chapterPosition, HeaderIndex(wsChapters, "isActive")).Value = NzBool(GetDictValue(chapter, "isActive", True))
        chapterPosition = chapterPosition + 1

        Dim rowsVariant As Variant
        rowsVariant = GetDictValue(chapter, "rows", Nothing)
        If TypeName(rowsVariant) = "Collection" Then
            Dim rowEntry As Variant
            For Each rowEntry In rowsVariant
                rowPosition = WriteRow(wsRows, rowEntry, chapterId, rowPosition)
            Next rowEntry
        End If

        Dim children As Variant
        children = GetDictValue(chapter, "children", Nothing)
        If TypeName(children) = "Collection" Then
            Dim child As Variant
            For Each child In children
                If TypeName(child) = "Dictionary" Then
                    If Not child.Exists("parentId") Then child("parentId") = chapterId
                    chapterQueue.Add child
                End If
            Next child
        End If
NextChapter:
    Loop
End Sub

Private Function WriteRow(ws As Worksheet, rowData As Variant, chapterId As String, rowPosition As Long) As Long
    If TypeName(rowData) <> "Dictionary" Then
        WriteRow = rowPosition
        Exit Function
    End If

    Dim rowId As String
    rowId = NzString(GetDictValue(rowData, "id", ""))
    If Len(rowId) = 0 Then
        WriteRow = rowPosition
        Exit Function
    End If

    Dim workChapter As String
    workChapter = NzString(GetDictValue(rowData, "chapterId", chapterId))

    Dim master As Dictionary
    Set master = GetDictValue(rowData, "master", Nothing)

    Dim workstate As Dictionary
    Set workstate = GetDictValue(rowData, "workstate", Nothing)

    Dim overrides As Dictionary
    Set overrides = GetDictValue(rowData, "overrides", GetDictValue(rowData, "workstate", Nothing))

    Dim customer As Dictionary
    Set customer = GetDictValue(rowData, "customer", Nothing)

    ws.Cells(rowPosition, HeaderIndex(ws, "rowId")).Value = rowId
    ws.Cells(rowPosition, HeaderIndex(ws, "chapterId")).Value = workChapter
    ws.Cells(rowPosition, HeaderIndex(ws, "titleOverride")).Value = NzString(GetDictValue(rowData, "titleOverride", ""))

    ws.Cells(rowPosition, HeaderIndex(ws, "masterFinding")).Value = NzString(GetDictValue(master, "finding", ""))
    Dim levels As Dictionary
    Set levels = GetDictValue(master, "levels", Nothing)
    ws.Cells(rowPosition, HeaderIndex(ws, "masterLevel1")).Value = NzString(GetDictValue(levels, "1", GetDictValue(master, "level1", "")))
    ws.Cells(rowPosition, HeaderIndex(ws, "masterLevel2")).Value = NzString(GetDictValue(levels, "2", GetDictValue(master, "level2", "")))
    ws.Cells(rowPosition, HeaderIndex(ws, "masterLevel3")).Value = NzString(GetDictValue(levels, "3", GetDictValue(master, "level3", "")))
    ws.Cells(rowPosition, HeaderIndex(ws, "masterLevel4")).Value = NzString(GetDictValue(levels, "4", GetDictValue(master, "level4", "")))

    Dim overrideFinding As String
    Dim findingEnabled As Boolean
    overrideFinding = NzString(GetDictValue(GetDictValue(overrides, "finding", Nothing), "text", GetDictValue(workstate, "findingOverride", "")))
    findingEnabled = NzBool(GetDictValue(GetDictValue(overrides, "finding", Nothing), "enabled", GetDictValue(workstate, "useFindingOverride", False)))
    ws.Cells(rowPosition, HeaderIndex(ws, "overrideFinding")).Value = overrideFinding
    ws.Cells(rowPosition, HeaderIndex(ws, "useOverrideFinding")).Value = findingEnabled

    Dim levelOverrides As Dictionary
    Set levelOverrides = GetDictValue(overrides, "levels", Nothing)
    Dim workLevelOverrides As Dictionary
    Set workLevelOverrides = GetDictValue(workstate, "levelOverrides", Nothing)
    Dim useMap As Dictionary
    Set useMap = GetDictValue(workstate, "useLevelOverride", Nothing)
    Dim levelOverrideText As New Dictionary
    levelOverrideText.CompareMode = TextCompare
    Dim levelOverrideUse As New Dictionary
    levelOverrideUse.CompareMode = TextCompare
    Dim i As Long
    For i = 1 To 4
        Dim levelPayload As Variant
        levelPayload = GetDictValue(levelOverrides, CStr(i), Empty)
        Dim levelText As String
        Dim levelEnabled As Boolean
        If TypeName(levelPayload) = "Dictionary" Then
            levelText = NzString(GetDictValue(levelPayload, "text", ""))
            levelEnabled = NzBool(GetDictValue(levelPayload, "enabled", False))
        Else
            levelText = NzString(levelPayload)
            levelEnabled = False
        End If
        If TypeName(workLevelOverrides) = "Dictionary" Then
            If levelText = "" Then levelText = NzString(GetDictValue(workLevelOverrides, CStr(i), ""))
        End If
        If TypeName(useMap) = "Dictionary" Then
            levelEnabled = NzBool(GetDictValue(useMap, CStr(i), levelEnabled))
        End If
        ws.Cells(rowPosition, HeaderIndex(ws, "overrideLevel" & i)).Value = levelText
        ws.Cells(rowPosition, HeaderIndex(ws, "useOverrideLevel" & i)).Value = levelEnabled
        levelOverrideText(CStr(i)) = levelText
        levelOverrideUse(CStr(i)) = levelEnabled
    Next i

    ws.Cells(rowPosition, HeaderIndex(ws, "customerAnswer")).Value = GetDictValue(customer, "answer", Null)
    ws.Cells(rowPosition, HeaderIndex(ws, "customerRemark")).Value = NzString(GetDictValue(customer, "remark", ""))
    ws.Cells(rowPosition, HeaderIndex(ws, "customerPriority")).Value = GetDictValue(customer, "priority", Null)

    ws.Cells(rowPosition, HeaderIndex(ws, "selectedLevel")).Value = GetDictValue(workstate, "selectedLevel", 2)
    ws.Cells(rowPosition, HeaderIndex(ws, "includeFinding")).Value = NzBool(GetDictValue(workstate, "includeFinding", True))
    ws.Cells(rowPosition, HeaderIndex(ws, "includeRecommendation")).Value = NzBool(GetDictValue(workstate, "includeRecommendation", True))
    ws.Cells(rowPosition, HeaderIndex(ws, "overwriteMode")).Value = NzString(GetDictValue(workstate, "overwriteMode", "append"))
    ws.Cells(rowPosition, HeaderIndex(ws, "done")).Value = NzBool(GetDictValue(workstate, "done", False))
    ws.Cells(rowPosition, HeaderIndex(ws, "notes")).Value = NzString(GetDictValue(workstate, "notes", ""))
    ws.Cells(rowPosition, HeaderIndex(ws, "lastEditedBy")).Value = NzString(GetDictValue(workstate, "lastEditedBy", ""))
    ws.Cells(rowPosition, HeaderIndex(ws, "lastEditedAt")).Value = NzString(GetDictValue(workstate, "lastEditedAt", ""))

    WriteRow = rowPosition + 1
End Function

Private Sub WriteChapterTitles(ws As Worksheet, rowIndex As Long, titles As Variant)
    Dim titleValue As String
    titleValue = ""
    If TypeName(titles) = "Dictionary" Then
        ws.Cells(rowIndex, HeaderIndex(ws, "defaultTitle_de")).Value = NzString(GetDictValue(titles, "de", ""))
        ws.Cells(rowIndex, HeaderIndex(ws, "defaultTitle_fr")).Value = NzString(GetDictValue(titles, "fr", GetDictValue(titles, "de", "")))
        ws.Cells(rowIndex, HeaderIndex(ws, "defaultTitle_it")).Value = NzString(GetDictValue(titles, "it", GetDictValue(titles, "de", "")))
        ws.Cells(rowIndex, HeaderIndex(ws, "defaultTitle_en")).Value = NzString(GetDictValue(titles, "en", GetDictValue(titles, "de", "")))
    Else
        titleValue = NzString(titles)
        ws.Cells(rowIndex, HeaderIndex(ws, "defaultTitle_de")).Value = titleValue
        ws.Cells(rowIndex, HeaderIndex(ws, "defaultTitle_fr")).Value = titleValue
        ws.Cells(rowIndex, HeaderIndex(ws, "defaultTitle_it")).Value = titleValue
        ws.Cells(rowIndex, HeaderIndex(ws, "defaultTitle_en")).Value = titleValue
    End If
End Sub

Private Sub WritePhotos(photoDict As Variant)
    Dim ws As Worksheet
    Set ws = EnsureSheetWithHeaders(SHEET_PHOTOS, HeaderPhotos(), False)
    ws.Cells.ClearContents
    WriteHeaderRow ws, HeaderPhotos()

    Dim wsTags As Worksheet
    Set wsTags = EnsureSheetWithHeaders(SHEET_PHOTO_TAGS, HeaderPhotoTags(), False)
    wsTags.Cells.ClearContents
    WriteHeaderRow wsTags, HeaderPhotoTags()

    If TypeName(photoDict) <> "Dictionary" Then Exit Sub

    Dim rowIndex As Long
    rowIndex = ROW_HEADER_ROW + 1
    Dim key As Variant
    For Each key In photoDict.Keys
        Dim payload As Dictionary
        Set payload = photoDict(key)
        ws.Cells(rowIndex, HeaderIndex(ws, "fileName")).Value = CStr(key)
        ws.Cells(rowIndex, HeaderIndex(ws, "notes")).Value = NzString(GetDictValue(payload, "notes", ""))
        Dim tags As Dictionary
        Set tags = GetDictValue(payload, "tags", Nothing)
        ws.Cells(rowIndex, HeaderIndex(ws, "preferredLocale")).Value = NzString(GetDictValue(payload, "preferredLocale", ""))
        WritePhotoTags wsTags, CStr(key), modABPhotoConstants.PHOTO_LIST_BERICHT, GetDictValue(tags, "bericht", Array())
        WritePhotoTags wsTags, CStr(key), modABPhotoConstants.PHOTO_LIST_SEMINAR, GetDictValue(tags, "seminar", Array())
        WritePhotoTags wsTags, CStr(key), modABPhotoConstants.PHOTO_LIST_TOPIC, GetDictValue(tags, "topic", Array())
        rowIndex = rowIndex + 1
    Next key
End Sub

Private Sub WritePhotoTags(wsTags As Worksheet, ByVal fileName As String, ByVal listName As String, tagArr As Variant)
    Dim asCollection As Collection
    If TypeName(tagArr) = "Collection" Then
        Set asCollection = tagArr
    Else
        Set asCollection = New Collection
        If IsArray(tagArr) Then
            Dim lb As Long, ub As Long, i As Long
            lb = LBound(tagArr): ub = UBound(tagArr)
            For i = lb To ub
                If Len(NzString(tagArr(i))) > 0 Then asCollection.Add NzString(tagArr(i))
            Next i
        ElseIf Len(NzString(tagArr)) > 0 Then
            asCollection.Add NzString(tagArr)
        End If
    End If

    Dim nextRow As Long
    nextRow = wsTags.Cells(wsTags.Rows.Count, 1).End(xlUp).Row + 1
    Dim item As Variant
    For Each item In asCollection
        wsTags.Cells(nextRow, HeaderIndex(wsTags, "fileName")).Value = fileName
        wsTags.Cells(nextRow, HeaderIndex(wsTags, "listName")).Value = listName
        wsTags.Cells(nextRow, HeaderIndex(wsTags, "tagValue")).Value = NzString(item)
        nextRow = nextRow + 1
    Next item
End Sub

Private Sub WriteLists(listDict As Variant)
    Dim ws As Worksheet
    Set ws = EnsureSheetWithHeaders(SHEET_LISTS, HeaderLists(), False)
    ws.Cells.ClearContents
    WriteHeaderRow ws, HeaderLists()

    If TypeName(listDict) <> "Dictionary" Then Exit Sub

    Dim rowIndex As Long
    rowIndex = ROW_HEADER_ROW + 1
    Dim listName As Variant
    For Each listName In listDict.Keys
        Dim entries As Variant
        entries = listDict(listName)
        If TypeName(entries) = "Collection" Then
            Dim entry As Variant
            For Each entry In entries
                rowIndex = WriteListEntry(ws, listName, entry, rowIndex)
            Next entry
        ElseIf TypeName(entries) = "Dictionary" Then
            rowIndex = WriteListEntry(ws, listName, entries, rowIndex)
        End If
    Next listName
End Sub

Private Function WriteListEntry(ws As Worksheet, listName As String, entry As Variant, rowIndex As Long) As Long
    If TypeName(entry) <> "Dictionary" Then
        WriteListEntry = rowIndex
        Exit Function
    End If

    ws.Cells(rowIndex, HeaderIndex(ws, "listName")).Value = listName
    ws.Cells(rowIndex, HeaderIndex(ws, "value")).Value = NzString(GetDictValue(entry, "value", ""))
    Dim labels As Dictionary
    Set labels = GetDictValue(entry, "labels", Nothing)
    ws.Cells(rowIndex, HeaderIndex(ws, "label_de")).Value = NzString(GetDictValue(labels, "de", GetDictValue(entry, "label", "")))
    ws.Cells(rowIndex, HeaderIndex(ws, "label_fr")).Value = NzString(GetDictValue(labels, "fr", GetDictValue(entry, "label", "")))
    ws.Cells(rowIndex, HeaderIndex(ws, "label_it")).Value = NzString(GetDictValue(labels, "it", GetDictValue(entry, "label", "")))
    ws.Cells(rowIndex, HeaderIndex(ws, "label_en")).Value = NzString(GetDictValue(labels, "en", GetDictValue(entry, "label", "")))
    ws.Cells(rowIndex, HeaderIndex(ws, "group")).Value = NzString(GetDictValue(entry, "group", listName))
    ws.Cells(rowIndex, HeaderIndex(ws, "sortOrder")).Value = NzNumber(GetDictValue(entry, "sortOrder", rowIndex - ROW_HEADER_ROW))
    ws.Cells(rowIndex, HeaderIndex(ws, "chapterId")).Value = NzString(GetDictValue(entry, "chapterId", ""))

    WriteListEntry = rowIndex + 1
End Function

Private Sub WriteHistory(history As Variant)
    Dim ws As Worksheet
    Set ws = EnsureSheetWithHeaders(SHEET_OVERRIDES_HISTORY, HeaderOverridesHistory(), False)
    ws.Cells.ClearContents
    WriteHeaderRow ws, HeaderOverridesHistory()

    If TypeName(history) <> "Collection" Then Exit Sub

    Dim rowIndex As Long
    rowIndex = ROW_HEADER_ROW + 1
    Dim entry As Variant
    For Each entry In history
        If TypeName(entry) <> "Dictionary" Then GoTo Continue
        ws.Cells(rowIndex, HeaderIndex(ws, "timestamp")).Value = GetDictValue(entry, "timestamp", Now)
        ws.Cells(rowIndex, HeaderIndex(ws, "rowId")).Value = NzString(GetDictValue(entry, "rowId", ""))
        ws.Cells(rowIndex, HeaderIndex(ws, "fieldName")).Value = NzString(GetDictValue(entry, "field", ""))
        ws.Cells(rowIndex, HeaderIndex(ws, "oldValue")).Value = NzString(GetDictValue(entry, "oldValue", ""))
        ws.Cells(rowIndex, HeaderIndex(ws, "newValue")).Value = NzString(GetDictValue(entry, "newValue", ""))
        ws.Cells(rowIndex, HeaderIndex(ws, "user")).Value = NzString(GetDictValue(entry, "user", ""))
        rowIndex = rowIndex + 1
Continue:
    Next entry
End Sub

Private Function GetDict(container As Dictionary, key As String) As Variant
    If container Is Nothing Then
        GetDict = Nothing
    ElseIf container.Exists(key) Then
        GetDict = container(key)
    Else
        GetDict = Nothing
    End If
End Function

Private Function JoinVariant(value As Variant) As String
    If IsEmpty(value) Or IsNull(value) Then
        JoinVariant = ""
    ElseIf TypeName(value) = "Collection" Then
        Dim parts() As String
        ReDim parts(1 To value.Count)
        Dim i As Long
        For i = 1 To value.Count
            parts(i) = CStr(value(i))
        Next i
        JoinVariant = Join(parts, ", ")
    ElseIf IsArray(value) Then
        Dim lb As Long, ub As Long
        lb = LBound(value)
        ub = UBound(value)
        If ub - lb + 1 <= 0 Then
            JoinVariant = ""
        Else
            Dim arr() As String
            ReDim arr(lb To ub)
            Dim idx As Long
            For idx = lb To ub
                arr(idx) = CStr(value(idx))
            Next idx
            JoinVariant = Join(arr, ", ")
        End If
    ElseIf TypeName(value) = "Dictionary" Then
        JoinVariant = Join(value.Keys, ", ")
    Else
        JoinVariant = CStr(value)
    End If
End Function

Private Function ReadTextFile(ByVal filePath As String) As String
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    ReadTextFile = Input$(LOF(fileNum), fileNum)
    Close #fileNum
End Function
