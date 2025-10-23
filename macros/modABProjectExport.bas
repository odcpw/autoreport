Attribute VB_Name = "modABProjectExport"
Option Explicit

'=============================================================
' Export structured sheets to project.json
'=============================================================

Public Sub ExportProjectJson(ByVal outputPath As String)
    EnsureAutoBerichtSheets
    Dim project As Dictionary
    Set project = BuildProjectSnapshot()

    Dim jsonText As String
    jsonText = JsonConverter.ConvertToJson(project, Whitespace:=2)
    WriteTextFile outputPath, jsonText
    MsgBox "project.json exported to " & outputPath, vbInformation
End Sub

Private Function BuildProjectSnapshot() As Dictionary
    Dim project As New Dictionary
    project.CompareMode = TextCompare

    project("version") = 1
    project("meta") = ReadMeta()
    project("chapters") = ReadChaptersWithRows()
    project("photos") = ReadPhotos()
    project("lists") = ReadLists()
    project("history") = ReadOverrideHistory()

    Set BuildProjectSnapshot = project
End Function

Private Function ReadMeta() As Dictionary
    Dim result As New Dictionary
    result.CompareMode = TextCompare

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_META)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = ROW_HEADER_ROW + 1 To lastRow
        Dim key As String
        key = NzString(ws.Cells(r, 1).Value)
        If Len(key) = 0 Then GoTo Continue
        result(key) = ws.Cells(r, 2).Value
Continue:
    Next r
    If Not result.Exists("locale") Then result("locale") = DefaultLocale()
    If Not result.Exists("createdAt") Then result("createdAt") = Format$(Now, "yyyy-mm-dd\THH:nn:ss")
    Set ReadMeta = result
End Function

Private Function ReadChaptersWithRows() As Collection
    Dim wsChapters As Worksheet
    Set wsChapters = ThisWorkbook.Worksheets(SHEET_CHAPTERS)
    Dim chapterMap As New Dictionary
    chapterMap.CompareMode = TextCompare

    Dim lastRow As Long
    lastRow = wsChapters.Cells(wsChapters.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = ROW_HEADER_ROW + 1 To lastRow
        Dim chapterId As String
        chapterId = NzString(wsChapters.Cells(r, HeaderIndex(wsChapters, "chapterId")).Value)
        If Len(chapterId) = 0 Then GoTo NextChapter
        Dim chapter As New Dictionary
        chapter.CompareMode = TextCompare
        chapter("id") = chapterId
        chapter("parentId") = NzString(wsChapters.Cells(r, HeaderIndex(wsChapters, "parentId")).Value)
        chapter("orderIndex") = wsChapters.Cells(r, HeaderIndex(wsChapters, "orderIndex")).Value
        Dim titles As New Dictionary
        titles.CompareMode = TextCompare
        titles("de") = NzString(wsChapters.Cells(r, HeaderIndex(wsChapters, "defaultTitle_de")).Value)
        titles("fr") = NzString(wsChapters.Cells(r, HeaderIndex(wsChapters, "defaultTitle_fr")).Value)
        titles("it") = NzString(wsChapters.Cells(r, HeaderIndex(wsChapters, "defaultTitle_it")).Value)
        titles("en") = NzString(wsChapters.Cells(r, HeaderIndex(wsChapters, "defaultTitle_en")).Value)
        chapter("title") = titles
        chapter("pageSize") = wsChapters.Cells(r, HeaderIndex(wsChapters, "pageSize")).Value
        chapter("isActive") = wsChapters.Cells(r, HeaderIndex(wsChapters, "isActive")).Value
        chapter("rows") = New Collection
        chapterMap(chapterId) = chapter
NextChapter:
    Next r

    Dim wsRows As Worksheet
    Set wsRows = ThisWorkbook.Worksheets(SHEET_ROWS)
    Dim rowsLast As Long
    rowsLast = wsRows.Cells(wsRows.Rows.Count, 1).End(xlUp).Row

    For r = ROW_HEADER_ROW + 1 To rowsLast
        Dim rowId As String
        rowId = NzString(wsRows.Cells(r, HeaderIndex(wsRows, "rowId")).Value)
        If Len(rowId) = 0 Then GoTo NextRow
        Dim chapterId As String
        chapterId = NzString(wsRows.Cells(r, HeaderIndex(wsRows, "chapterId")).Value)
        If Not chapterMap.Exists(chapterId) Then
            Dim placeholder As New Dictionary
            placeholder.CompareMode = TextCompare
            placeholder("id") = chapterId
            placeholder("parentId") = GetParentChapterId(chapterId)
            placeholder("orderIndex") = chapterMap.Count + 1
            Dim t As New Dictionary
            t.CompareMode = TextCompare
            t("de") = chapterId
            placeholder("title") = t
            placeholder("pageSize") = 5
            placeholder("isActive") = True
            placeholder("rows") = New Collection
            chapterMap(chapterId) = placeholder
        End If
        Dim rowEntry As Dictionary
        Set rowEntry = RowToDictionary(wsRows, r)
        chapterMap(chapterId)("rows").Add rowEntry
NextRow:
    Next r

    Dim ordered As New Collection
    Dim chapterIdKey As Variant
    Dim temp As New Collection
    For Each chapterIdKey In chapterMap.Keys
        temp.Add chapterMap(chapterIdKey)
    Next chapterIdKey
    SortChapterCollection temp
    For Each chapterIdKey In temp
        ordered.Add chapterIdKey
    Next chapterIdKey
    Set ReadChaptersWithRows = ordered
End Function

Private Function RowToDictionary(ws As Worksheet, rowIndex As Long) As Dictionary
    Dim entry As New Dictionary
    entry.CompareMode = TextCompare

    entry("id") = NzString(ws.Cells(rowIndex, HeaderIndex(ws, "rowId")).Value)
    entry("chapterId") = NzString(ws.Cells(rowIndex, HeaderIndex(ws, "chapterId")).Value)
    entry("titleOverride") = NzString(ws.Cells(rowIndex, HeaderIndex(ws, "titleOverride")).Value)

    Dim master As New Dictionary
    master.CompareMode = TextCompare
    master("finding") = NzString(ws.Cells(rowIndex, HeaderIndex(ws, "masterFinding")).Value)
    Dim levels As New Dictionary
    levels.CompareMode = TextCompare
    levels("1") = NzString(ws.Cells(rowIndex, HeaderIndex(ws, "masterLevel1")).Value)
    levels("2") = NzString(ws.Cells(rowIndex, HeaderIndex(ws, "masterLevel2")).Value)
    levels("3") = NzString(ws.Cells(rowIndex, HeaderIndex(ws, "masterLevel3")).Value)
    levels("4") = NzString(ws.Cells(rowIndex, HeaderIndex(ws, "masterLevel4")).Value)
    master("levels") = levels
    entry("master") = master

    Dim overrides As New Dictionary
    overrides.CompareMode = TextCompare
    Dim findingOverride As New Dictionary
    findingOverride.CompareMode = TextCompare
    findingOverride("text") = NzString(ws.Cells(rowIndex, HeaderIndex(ws, "overrideFinding")).Value)
    Dim findingEnabled As Boolean
    findingEnabled = ws.Cells(rowIndex, HeaderIndex(ws, "useOverrideFinding")).Value
    findingOverride("enabled") = findingEnabled
    overrides("finding") = findingOverride

    Dim levelOverrides As New Dictionary
    levelOverrides.CompareMode = TextCompare
    Dim levelOverrideText As New Dictionary
    levelOverrideText.CompareMode = TextCompare
    Dim levelOverrideUse As New Dictionary
    levelOverrideUse.CompareMode = TextCompare
    Dim i As Long
    For i = 1 To 4
        Dim levelDict As New Dictionary
        levelDict.CompareMode = TextCompare
        Dim levelText As String
        levelText = NzString(ws.Cells(rowIndex, HeaderIndex(ws, "overrideLevel" & i)).Value)
        Dim levelEnabled As Boolean
        levelEnabled = ws.Cells(rowIndex, HeaderIndex(ws, "useOverrideLevel" & i)).Value
        levelDict("text") = levelText
        levelDict("enabled") = levelEnabled
        levelOverrides(CStr(i)) = levelDict
        levelOverrideText(CStr(i)) = levelText
        levelOverrideUse(CStr(i)) = levelEnabled
    Next i
    overrides("levels") = levelOverrides
    entry("overrides") = overrides

    Dim customer As New Dictionary
    customer.CompareMode = TextCompare
    customer("answer") = ws.Cells(rowIndex, HeaderIndex(ws, "customerAnswer")).Value
    customer("remark") = NzString(ws.Cells(rowIndex, HeaderIndex(ws, "customerRemark")).Value)
    customer("priority") = ws.Cells(rowIndex, HeaderIndex(ws, "customerPriority")).Value
    entry("customer") = customer

    Dim workstate As New Dictionary
    workstate.CompareMode = TextCompare
    workstate("selectedLevel") = ws.Cells(rowIndex, HeaderIndex(ws, "selectedLevel")).Value
    workstate("includeFinding") = ws.Cells(rowIndex, HeaderIndex(ws, "includeFinding")).Value
    workstate("includeRecommendation") = ws.Cells(rowIndex, HeaderIndex(ws, "includeRecommendation")).Value
    workstate("overwriteMode") = NzString(ws.Cells(rowIndex, HeaderIndex(ws, "overwriteMode")).Value)
    workstate("done") = ws.Cells(rowIndex, HeaderIndex(ws, "done")).Value
    workstate("notes") = NzString(ws.Cells(rowIndex, HeaderIndex(ws, "notes")).Value)
    workstate("lastEditedBy") = NzString(ws.Cells(rowIndex, HeaderIndex(ws, "lastEditedBy")).Value)
    workstate("lastEditedAt") = NzString(ws.Cells(rowIndex, HeaderIndex(ws, "lastEditedAt")).Value)
    workstate("findingOverride") = NzString(ws.Cells(rowIndex, HeaderIndex(ws, "overrideFinding")).Value)
    workstate("useFindingOverride") = findingEnabled
    workstate("levelOverrides") = levelOverrideText
    workstate("useLevelOverride") = levelOverrideUse
    entry("workstate") = workstate

    Set RowToDictionary = entry
End Function

Private Sub SortChapterCollection(col As Collection)
    Dim i As Long, j As Long
    For i = 1 To col.Count - 1
        For j = i + 1 To col.Count
            Dim a As Dictionary, b As Dictionary
            Set a = col(i)
            Set b = col(j)
            If NzNumber(a("orderIndex")) > NzNumber(b("orderIndex")) Then
                col.Remove j
                col.Add b, , i
            End If
        Next j
    Next i
End Sub

Private Function ReadPhotos() As Dictionary
    Dim result As New Dictionary
    result.CompareMode = TextCompare

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_PHOTOS)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim r As Long
    For r = ROW_HEADER_ROW + 1 To lastRow
        Dim fileName As String
        fileName = NzString(ws.Cells(r, HeaderIndex(ws, "fileName")).Value)
        If Len(fileName) = 0 Then GoTo Continue
        Dim photo As New Dictionary
        photo.CompareMode = TextCompare
        photo("displayName") = NzString(ws.Cells(r, HeaderIndex(ws, "displayName")).Value)
        photo("notes") = NzString(ws.Cells(r, HeaderIndex(ws, "notes")).Value)
        Dim tags As New Dictionary
        tags.CompareMode = TextCompare
        tags("chapters") = SplitTags(ws.Cells(r, HeaderIndex(ws, "tagChapters")).Value)
        tags("categories") = SplitTags(ws.Cells(r, HeaderIndex(ws, "tagCategories")).Value)
        tags("training") = SplitTags(ws.Cells(r, HeaderIndex(ws, "tagTraining")).Value)
        tags("subfolders") = SplitTags(ws.Cells(r, HeaderIndex(ws, "tagSubfolders")).Value)
        photo("tags") = tags
        photo("preferredLocale") = NzString(ws.Cells(r, HeaderIndex(ws, "preferredLocale")).Value)
        photo("capturedAt") = NzString(ws.Cells(r, HeaderIndex(ws, "capturedAt")).Value)
        result(fileName) = photo
Continue:
    Next r
    Set ReadPhotos = result
End Function

Private Function SplitTags(value As Variant) As Variant
    Dim result() As String
    If IsEmpty(value) Or IsNull(value) Then
        SplitTags = Array()
        Exit Function
    End If
    Dim text As String
    text = Trim$(CStr(value))
    If Len(text) = 0 Then
        SplitTags = Array()
        Exit Function
    End If
    Dim parts() As String
    parts = Split(text, ",")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        parts(i) = Trim$(parts(i))
    Next i
    SplitTags = parts
End Function

Private Function ReadLists() As Dictionary
    Dim result As New Dictionary
    result.CompareMode = TextCompare

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_LISTS)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = ROW_HEADER_ROW + 1 To lastRow
        Dim listName As String
        listName = NzString(ws.Cells(r, HeaderIndex(ws, "listName")).Value)
        If Len(listName) = 0 Then GoTo Continue
        Dim item As New Dictionary
        item.CompareMode = TextCompare
        item("value") = NzString(ws.Cells(r, HeaderIndex(ws, "value")).Value)
        item("label") = NzString(ws.Cells(r, HeaderIndex(ws, "label_de")).Value)
        item("labels") = BuildLabelMap(ws, r)
        item("group") = NzString(ws.Cells(r, HeaderIndex(ws, "group")).Value)
        item("sortOrder") = ws.Cells(r, HeaderIndex(ws, "sortOrder")).Value
        item("chapterId") = NzString(ws.Cells(r, HeaderIndex(ws, "chapterId")).Value)
        If Not result.Exists(listName) Then
            Dim bucket As New Collection
            result(listName) = bucket
        End If
        result(listName).Add item
Continue:
    Next r
    Set ReadLists = result
End Function

Private Function BuildLabelMap(ws As Worksheet, rowIndex As Long) As Dictionary
    Dim labels As New Dictionary
    labels.CompareMode = TextCompare
    labels("de") = NzString(ws.Cells(rowIndex, HeaderIndex(ws, "label_de")).Value)
    labels("fr") = NzString(ws.Cells(rowIndex, HeaderIndex(ws, "label_fr")).Value)
    labels("it") = NzString(ws.Cells(rowIndex, HeaderIndex(ws, "label_it")).Value)
    labels("en") = NzString(ws.Cells(rowIndex, HeaderIndex(ws, "label_en")).Value)
    Set BuildLabelMap = labels
End Function

Private Function ReadOverrideHistory() As Collection
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_OVERRIDES_HISTORY)
    Dim result As New Collection
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = ROW_HEADER_ROW + 1 To lastRow
        Dim entry As New Dictionary
        entry.CompareMode = TextCompare
        entry("timestamp") = ws.Cells(r, HeaderIndex(ws, "timestamp")).Value
        entry("rowId") = NzString(ws.Cells(r, HeaderIndex(ws, "rowId")).Value)
        entry("field") = NzString(ws.Cells(r, HeaderIndex(ws, "fieldName")).Value)
        entry("oldValue") = NzString(ws.Cells(r, HeaderIndex(ws, "oldValue")).Value)
        entry("newValue") = NzString(ws.Cells(r, HeaderIndex(ws, "newValue")).Value)
        entry("user") = NzString(ws.Cells(r, HeaderIndex(ws, "user")).Value)
        result.Add entry
    Next r
    Set ReadOverrideHistory = result
End Function

Private Sub WriteTextFile(ByVal filePath As String, ByVal content As String)
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, content
    Close #fileNum
End Sub

Private Function NzNumber(value As Variant) As Double
    If IsNumeric(value) Then
        NzNumber = CDbl(value)
    Else
        NzNumber = 99999
    End If
End Function


Private Function GetParentChapterId(childId As String) As String
    Dim trimmed As String
    trimmed = Trim$(childId)
    Do While Len(trimmed) > 0 And Right$(trimmed, 1) = "."
        trimmed = Left$(trimmed, Len(trimmed) - 1)
    Loop
    Dim parts() As String
    parts = Split(trimmed, ".")
    If UBound(parts) <= 0 Then
        GetParentChapterId = ""
    Else
        Dim i As Long
        Dim result As String
        result = parts(0)
        For i = 1 To UBound(parts) - 1
            result = result & "." & parts(i)
        Next i
        GetParentChapterId = result
    End If
End Function
