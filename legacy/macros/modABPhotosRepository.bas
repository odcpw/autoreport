Attribute VB_Name = "modABPhotosRepository"
Option Explicit

'=============================================================
' Photo metadata and button catalog helpers
'=============================================================

Private Sub EnsurePhotoEntryDefaults(ByRef entry As Scripting.Dictionary)
    If entry Is Nothing Then Exit Sub
    entry.CompareMode = TextCompare
    If Not entry.Exists("fileName") Then entry("fileName") = ""
    If Not entry.Exists("notes") Then entry("notes") = ""
    If Not entry.Exists(modABPhotoConstants.PHOTO_TAG_BERICHT) Then entry(modABPhotoConstants.PHOTO_TAG_BERICHT) = ""
    If Not entry.Exists(modABPhotoConstants.PHOTO_TAG_SEMINAR) Then entry(modABPhotoConstants.PHOTO_TAG_SEMINAR) = ""
    If Not entry.Exists(modABPhotoConstants.PHOTO_TAG_TOPIC) Then entry(modABPhotoConstants.PHOTO_TAG_TOPIC) = ""
    If Not entry.Exists("preferredLocale") Then entry("preferredLocale") = ""
End Sub

Public Function PhotosSheet() As Worksheet
    EnsureAutoBerichtSheets
    Set PhotosSheet = ThisWorkbook.Worksheets(SHEET_PHOTOS)
End Function

Public Function PhotoTagsSheet() As Worksheet
    EnsureAutoBerichtSheets
    Set PhotoTagsSheet = ThisWorkbook.Worksheets(SHEET_PHOTO_TAGS)
End Function

Public Function ListsSheet() As Worksheet
    EnsureAutoBerichtSheets
    Set ListsSheet = ThisWorkbook.Worksheets(SHEET_LISTS)
End Function

Public Sub EnsurePhotoRecord(fileName As String)
    Dim entry As Scripting.Dictionary
    Set entry = GetPhotoEntry(fileName)
    If entry Is Nothing Then
        Dim newEntry As New Scripting.Dictionary
        newEntry.CompareMode = TextCompare
        newEntry("fileName") = fileName
        newEntry("notes") = ""
        newEntry(modABPhotoConstants.PHOTO_TAG_BERICHT) = ""
        newEntry(modABPhotoConstants.PHOTO_TAG_SEMINAR) = ""
        newEntry(modABPhotoConstants.PHOTO_TAG_TOPIC) = ""
        newEntry("preferredLocale") = ""
        EnsurePhotoEntryDefaults newEntry
        UpsertPhoto newEntry
    End If
End Sub

Public Function GetPhotoEntry(fileName As String) As Scripting.Dictionary
    Dim ws As Worksheet
    Set ws = PhotosSheet()
    Dim rowIndex As Long
    rowIndex = FindRowIndex(ws, "fileName", fileName)
    If rowIndex = 0 Then
        Set GetPhotoEntry = Nothing
        Exit Function
    End If

    Dim headers As Variant
    headers = ws.Range(ws.Cells(ROW_HEADER_ROW, 1), ws.Cells(ROW_HEADER_ROW, ws.UsedRange.Columns.Count)).Value
    Dim entry As New Scripting.Dictionary
    entry.CompareMode = TextCompare

    Dim c As Long
    For c = LBound(headers, 2) To UBound(headers, 2)
        entry(CStr(headers(1, c))) = ws.Cells(rowIndex, c).Value
    Next c
    EnsurePhotoEntryDefaults entry
    ' Hydrate tag fields for UI convenience from PhotoTags table
    Dim tags As Dictionary
    Set tags = GetPhotoTagsDict(NzString(entry("fileName")))
    entry(modABPhotoConstants.PHOTO_TAG_BERICHT) = JoinTags(GetDictValue(tags, modABPhotoConstants.PHOTO_LIST_BERICHT, Array()))
    entry(modABPhotoConstants.PHOTO_TAG_SEMINAR) = JoinTags(GetDictValue(tags, modABPhotoConstants.PHOTO_LIST_SEMINAR, Array()))
    entry(modABPhotoConstants.PHOTO_TAG_TOPIC) = JoinTags(GetDictValue(tags, modABPhotoConstants.PHOTO_LIST_TOPIC, Array()))
    Set GetPhotoEntry = entry
End Function

Public Sub UpsertPhoto(entry As Scripting.Dictionary)
    Dim ws As Worksheet
    Set ws = PhotosSheet()
    EnsurePhotoEntryDefaults entry
    UpsertRow ws, "fileName", entry
End Sub

Public Sub RemoveMissingPhotos(currentEntries As Scripting.Dictionary)
    Dim ws As Worksheet
    Set ws = PhotosSheet()
    Dim nameCol As Long
    nameCol = HeaderIndex(ws, "fileName")
    If nameCol = 0 Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, nameCol).End(xlUp).Row

    Dim r As Long
    For r = lastRow To ROW_HEADER_ROW + 1 Step -1
        Dim key As String
        key = NzString(ws.Cells(r, nameCol).Value)
        If Len(key) > 0 Then
            If currentEntries Is Nothing Or Not currentEntries.Exists(key) Then
                RemoveTagsForFile key
                ws.Rows(r).Delete
            End If
        End If
    Next r
End Sub

Public Function GetButtonList(listName As String, locale As String) As Collection
    Dim ws As Worksheet
    Set ws = ListsSheet()

    Dim listColumn As Long
    listColumn = HeaderIndex(ws, "listName")
    If listColumn = 0 Then
        Set GetButtonList = New Collection
        Exit Function
    End If

    Dim valueCol As Long
    valueCol = HeaderIndex(ws, "value")
    Dim groupCol As Long
    groupCol = HeaderIndex(ws, "group")
    Dim sortCol As Long
    sortCol = HeaderIndex(ws, "sortOrder")
    Dim chapterCol As Long
    chapterCol = HeaderIndex(ws, "chapterId")

    Dim labelCol As Long
    labelCol = HeaderIndex(ws, "label_" & Replace(locale, "-", "_"))
    If labelCol = 0 Then labelCol = HeaderIndex(ws, "label_de")

    Dim result As New Collection
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, listColumn).End(xlUp).Row

    Dim r As Long
    For r = ROW_HEADER_ROW + 1 To lastRow
        If StrComp(NzString(ws.Cells(r, listColumn).Value), listName, vbTextCompare) = 0 Then
            Dim item As Scripting.Dictionary
            Set item = New Scripting.Dictionary
            item.CompareMode = TextCompare
            item("value") = NzString(ws.Cells(r, valueCol).Value)
            item("label") = NzString(ws.Cells(r, labelCol).Value)
            If groupCol > 0 Then item("group") = NzString(ws.Cells(r, groupCol).Value)
            If sortCol > 0 Then item("sortOrder") = ws.Cells(r, sortCol).Value
            If chapterCol > 0 Then item("chapterId") = NzString(ws.Cells(r, chapterCol).Value)
            result.Add item
        End If
    Next r

    SortButtonCollection result
    Set GetButtonList = result
End Function

Private Sub SortButtonCollection(items As Collection)
    Dim i As Long, j As Long
    For i = 1 To items.Count - 1
        For j = i + 1 To items.Count
            Dim a As Scripting.Dictionary, b As Scripting.Dictionary
            Set a = items(i)
            Set b = items(j)
            Dim aOrder As Double, bOrder As Double
            aOrder = NzNumber(GetDictValue(a, "sortOrder"), 99999)
            bOrder = NzNumber(GetDictValue(b, "sortOrder"), 99999)
            If aOrder > bOrder Then
                items.Remove j
                items.Add b, , i
            End If
        Next j
    Next i
End Sub

Public Function JoinTags(tags As Variant) As String
    Dim result As String
    result = ""

    If TypeName(tags) = "Collection" Then
        Dim count As Long
        count = tags.Count
        If count = 0 Then
            JoinTags = ""
            Exit Function
        End If
        Dim temp() As String
        ReDim temp(1 To count)
        Dim i As Long
        For i = 1 To count
            temp(i) = CStr(tags(i))
        Next i
        result = Join(temp, ", ")
    ElseIf IsArray(tags) Then
        Dim lb As Long, ub As Long
        lb = LBound(tags)
        ub = UBound(tags)
        If ub - lb + 1 <= 0 Then
            JoinTags = ""
            Exit Function
        End If
        Dim parts() As String
        ReDim parts(lb To ub)
        Dim idx As Long
        For idx = lb To ub
            parts(idx) = CStr(tags(idx))
        Next idx
        result = Join(parts, ", ")
    Else
        result = NzString(tags)
    End If

    JoinTags = result
End Function

Public Function BuildFolderTagLookup() As Scripting.Dictionary
    Dim map As New Scripting.Dictionary
    map.CompareMode = TextCompare

    Dim ws As Worksheet
    Set ws = ListsSheet()

    Dim colListName As Long
    Dim colValue As Long
    Dim colLabelDe As Long
    Dim colLabelFr As Long
    Dim colLabelIt As Long
    Dim colLabelEn As Long

    colListName = HeaderIndex(ws, "listName")
    colValue = HeaderIndex(ws, "value")
    colLabelDe = HeaderIndex(ws, "label_de")
    colLabelFr = HeaderIndex(ws, "label_fr")
    colLabelIt = HeaderIndex(ws, "label_it")
    colLabelEn = HeaderIndex(ws, "label_en")

    If colListName = 0 Then
        Set BuildFolderTagLookup = map
        Exit Function
    End If

    Dim listToField As New Scripting.Dictionary
    listToField.CompareMode = TextCompare
    listToField(PHOTO_LIST_BERICHT) = PHOTO_TAG_BERICHT
    listToField(PHOTO_LIST_SEMINAR) = PHOTO_TAG_SEMINAR
    listToField(PHOTO_LIST_TOPIC) = PHOTO_TAG_TOPIC

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, colListName).End(xlUp).Row

    Dim r As Long
    For r = ROW_HEADER_ROW + 1 To lastRow
        Dim listName As String
        listName = NzString(ws.Cells(r, colListName).Value)
        If Not listToField.Exists(listName) Then GoTo ContinueRow

        Dim tagField As String
        tagField = CStr(listToField(listName))

        Dim tagValue As String
        If colValue > 0 Then
            tagValue = NzString(ws.Cells(r, colValue).Value)
        Else
            tagValue = ""
        End If
        If Len(tagValue) = 0 Then GoTo ContinueRow

        AddFolderMapping map, NzString(ws.Cells(r, colLabelDe).Value), tagField, tagValue
        AddFolderMapping map, NzString(ws.Cells(r, colLabelFr).Value), tagField, tagValue
        AddFolderMapping map, NzString(ws.Cells(r, colLabelIt).Value), tagField, tagValue
        AddFolderMapping map, NzString(ws.Cells(r, colLabelEn).Value), tagField, tagValue
ContinueRow:
    Next r

    Set BuildFolderTagLookup = map
End Function

Public Function CollectFolderTags(ByVal relativePath As String, folderMap As Scripting.Dictionary) As Scripting.Dictionary
    Dim result As New Scripting.Dictionary
    result.CompareMode = TextCompare
    result(PHOTO_LIST_BERICHT) = New Scripting.Dictionary
    result(PHOTO_LIST_SEMINAR) = New Scripting.Dictionary
    result(PHOTO_LIST_TOPIC) = New Scripting.Dictionary

    If folderMap Is Nothing Or folderMap.Count = 0 Then
        Set CollectFolderTags = result
        Exit Function
    End If

    Dim normalizedPath As String
    normalizedPath = Replace(relativePath, "/", "\")

    Dim segments() As String
    segments = Split(normalizedPath, "\")
    If UBound(segments) < 0 Then
        Set CollectFolderTags = result
        Exit Function
    End If

    Dim i As Long
    For i = LBound(segments) To UBound(segments) - 1
        Dim folderToken As String
        folderToken = NormalizeFolderName(segments(i))
        If Len(folderToken) = 0 Then GoTo ContinueSegment
        Dim lookupKey As String
        lookupKey = LCase$(folderToken)
        If folderMap.Exists(lookupKey) Then
            Dim entries As Collection
            Set entries = folderMap(lookupKey)
            Dim desc As Scripting.Dictionary
            For Each desc In entries
                If result.Exists(desc("listName")) Then
                    Dim bucket As Scripting.Dictionary
                    Set bucket = result(desc("listName"))
                    bucket(desc("value")) = True
                End If
            Next desc
        End If
ContinueSegment:
    Next i

    Set CollectFolderTags = result
End Function

Private Sub AddFolderMapping(ByRef map As Scripting.Dictionary, ByVal labelValue As Variant, ByVal tagField As String, ByVal tagValue As String)
    Dim token As String
    token = NormalizeFolderName(labelValue)
    If Len(token) = 0 Then Exit Sub

    Dim key As String
    key = LCase$(token)

    Dim entries As Collection
    If map.Exists(key) Then
        Set entries = map(key)
    Else
        Set entries = New Collection
        Set map(key) = entries
    End If

    Dim existing As Scripting.Dictionary
    For Each existing In entries
        If StrComp(existing("field"), tagField, vbTextCompare) = 0 _
            And StrComp(existing("value"), tagValue, vbTextCompare) = 0 Then
            Exit Sub
        End If
    Next existing

    Dim descriptor As New Scripting.Dictionary
    descriptor.CompareMode = TextCompare
    descriptor("field") = tagField
    descriptor("value") = tagValue
    entries.Add descriptor
End Sub

Public Function NormalizeFolderName(ByVal rawValue As Variant) As String
    Dim textValue As String
    textValue = NzString(rawValue)
    textValue = Replace(textValue, Chr$(160), " ")
    textValue = Trim$(textValue)
    Do While InStr(textValue, "  ") > 0
        textValue = Replace(textValue, "  ", " ")
    Loop
    NormalizeFolderName = textValue
End Function

Private Function ExistingTagDictionary(ByVal record As Scripting.Dictionary, ByVal fieldName As String) As Scripting.Dictionary
    Dim dict As New Scripting.Dictionary
    dict.CompareMode = TextCompare

    If Not record Is Nothing Then
        If record.Exists(fieldName) Then
            Dim raw As String
            raw = NzString(record(fieldName))
            If Len(raw) > 0 Then
                Dim parts() As String
                parts = Split(raw, ",")
                Dim idx As Long
                For idx = LBound(parts) To UBound(parts)
                    Dim trimmed As String
                    trimmed = Trim$(parts(idx))
                    If Len(trimmed) > 0 Then dict(trimmed) = True
                Next idx
            End If
        End If
    End If

    Set ExistingTagDictionary = dict
End Function

Public Function JoinDictionaryKeys(ByVal dict As Scripting.Dictionary) As String
    If dict Is Nothing Then Exit Function
    If dict.Count = 0 Then
        JoinDictionaryKeys = ""
    Else
        Dim keys As Variant
        keys = dict.Keys
        Dim i As Long, j As Long
        For i = LBound(keys) To UBound(keys) - 1
            For j = i + 1 To UBound(keys)
                If StrComp(keys(i), keys(j), vbTextCompare) > 0 Then
                    Dim tmp As Variant
                    tmp = keys(i)
                    keys(i) = keys(j)
                    keys(j) = tmp
                End If
            Next j
        Next i
        JoinDictionaryKeys = Join(keys, ", ")
    End If
End Function

'=====================
' PhotoTags table I/O
'=====================
Public Function GetTagList(fileName As String, tagField As String) As Collection
    Dim listName As String
    listName = FieldToListName(tagField)
    Set GetTagList = GetTagsForFileAndList(fileName, listName)
End Function

Public Function GetTagsForFileAndList(ByVal fileName As String, ByVal listName As String) As Collection
    Dim result As New Collection
    If Len(fileName) = 0 Or Len(listName) = 0 Then
        Set GetTagsForFileAndList = result
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = PhotoTagsSheet()
    Dim colFile As Long, colList As Long, colTag As Long
    colFile = HeaderIndex(ws, "fileName")
    colList = HeaderIndex(ws, "listName")
    colTag = HeaderIndex(ws, "tagValue")
    If colFile = 0 Or colList = 0 Or colTag = 0 Then
        Set GetTagsForFileAndList = result
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, colFile).End(xlUp).Row
    Dim r As Long
    For r = ROW_HEADER_ROW + 1 To lastRow
        If StrComp(NzString(ws.Cells(r, colFile).Value), fileName, vbTextCompare) = 0 _
            And StrComp(NzString(ws.Cells(r, colList).Value), listName, vbTextCompare) = 0 Then
            Dim val As String
            val = NzString(ws.Cells(r, colTag).Value)
            If Len(val) > 0 Then result.Add val
        End If
    Next r
    Set GetTagsForFileAndList = result
End Function

Public Sub SetPhotoTags(fileName As String, tagField As String, tags As Variant)
    Dim listName As String
    listName = FieldToListName(tagField)
    If Len(listName) = 0 Then Exit Sub
    SetTagList fileName, listName, tags
End Sub

Public Sub TogglePhotoTag(fileName As String, tagField As String, tagValue As String)
    EnsurePhotoRecord fileName
    Dim listName As String
    listName = FieldToListName(tagField)
    If Len(listName) = 0 Then Exit Sub

    Dim current As Collection
    Set current = GetTagsForFileAndList(fileName, listName)

    Dim dict As New Scripting.Dictionary
    dict.CompareMode = TextCompare

    Dim i As Long
    For i = 1 To current.Count
        dict(current(i)) = True
    Next i

    If dict.Exists(tagValue) Then
        dict.Remove tagValue
    Else
        dict(tagValue) = True
    End If

    SetTagList fileName, listName, dict.Keys
End Sub

Public Sub SetTagList(ByVal fileName As String, ByVal listName As String, tags As Variant)
    If Len(fileName) = 0 Or Len(listName) = 0 Then Exit Sub
    Dim ws As Worksheet
    Set ws = PhotoTagsSheet()
    Dim colFile As Long, colList As Long, colTag As Long
    colFile = HeaderIndex(ws, "fileName")
    colList = HeaderIndex(ws, "listName")
    colTag = HeaderIndex(ws, "tagValue")
    If colFile = 0 Or colList = 0 Or colTag = 0 Then Exit Sub

    ' Remove existing
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, colFile).End(xlUp).Row
    Dim r As Long
    For r = lastRow To ROW_HEADER_ROW + 1 Step -1
        If StrComp(NzString(ws.Cells(r, colFile).Value), fileName, vbTextCompare) = 0 _
            And StrComp(NzString(ws.Cells(r, colList).Value), listName, vbTextCompare) = 0 Then
            ws.Rows(r).Delete
        End If
    Next r

    Dim values As Collection
    Set values = NormalizeToCollection(tags)
    If values.Count = 0 Then Exit Sub

    Dim insertRow As Long
    insertRow = ws.Cells(ws.Rows.Count, colFile).End(xlUp).Row + 1
    Dim item As Variant
    For Each item In values
        ws.Cells(insertRow, colFile).Value = fileName
        ws.Cells(insertRow, colList).Value = listName
        ws.Cells(insertRow, colTag).Value = item
        insertRow = insertRow + 1
    Next item
End Sub

Public Sub SetAllTagsForFile(ByVal fileName As String, tagsDict As Dictionary)
    If Len(fileName) = 0 Then Exit Sub
    Dim listName As Variant
    For Each listName In tagsDict.Keys
        Dim bucket As Scripting.Dictionary
        Set bucket = tagsDict(listName)
        SetTagList fileName, CStr(listName), bucket.Keys
    Next listName
End Sub

Public Sub RemoveTagsForFile(ByVal fileName As String)
    If Len(fileName) = 0 Then Exit Sub
    Dim ws As Worksheet
    Set ws = PhotoTagsSheet()
    Dim colFile As Long
    colFile = HeaderIndex(ws, "fileName")
    If colFile = 0 Then Exit Sub
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, colFile).End(xlUp).Row
    Dim r As Long
    For r = lastRow To ROW_HEADER_ROW + 1 Step -1
        If StrComp(NzString(ws.Cells(r, colFile).Value), fileName, vbTextCompare) = 0 Then
            ws.Rows(r).Delete
        End If
    Next r
End Sub

Public Function GetPhotoTagsDict(ByVal fileName As String) As Dictionary
    Dim result As New Dictionary
    result.CompareMode = TextCompare
    result(modABPhotoConstants.PHOTO_LIST_BERICHT) = Array()
    result(modABPhotoConstants.PHOTO_LIST_SEMINAR) = Array()
    result(modABPhotoConstants.PHOTO_LIST_TOPIC) = Array()

    If Len(fileName) = 0 Then
        Set GetPhotoTagsDict = result
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = PhotoTagsSheet()
    Dim colFile As Long, colList As Long, colTag As Long
    colFile = HeaderIndex(ws, "fileName")
    colList = HeaderIndex(ws, "listName")
    colTag = HeaderIndex(ws, "tagValue")
    If colFile = 0 Or colList = 0 Or colTag = 0 Then
        Set GetPhotoTagsDict = result
        Exit Function
    End If

    Dim temp As New Scripting.Dictionary
    temp.CompareMode = TextCompare
    temp(modABPhotoConstants.PHOTO_LIST_BERICHT) = New Collection
    temp(modABPhotoConstants.PHOTO_LIST_SEMINAR) = New Collection
    temp(modABPhotoConstants.PHOTO_LIST_TOPIC) = New Collection

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, colFile).End(xlUp).Row
    Dim r As Long
    For r = ROW_HEADER_ROW + 1 To lastRow
        If StrComp(NzString(ws.Cells(r, colFile).Value), fileName, vbTextCompare) = 0 Then
            Dim ln As String
            ln = NzString(ws.Cells(r, colList).Value)
            Dim tv As String
            tv = NzString(ws.Cells(r, colTag).Value)
            If temp.Exists(ln) And Len(tv) > 0 Then temp(ln).Add tv
        End If
    Next r

    Dim lnKey As Variant
    For Each lnKey In temp.Keys
        result(lnKey) = temp(lnKey)
    Next lnKey
    Set GetPhotoTagsDict = result
End Function

Private Function NormalizeToCollection(tags As Variant) As Collection
    Dim result As New Collection
    If TypeName(tags) = "Collection" Then
        Dim item As Variant
        For Each item In tags
            If Len(NzString(item)) > 0 Then result.Add NzString(item)
        Next item
    ElseIf IsArray(tags) Then
        Dim lb As Long, ub As Long, i As Long
        lb = LBound(tags): ub = UBound(tags)
        For i = lb To ub
            If Len(NzString(tags(i))) > 0 Then result.Add NzString(tags(i))
        Next i
    ElseIf Len(NzString(tags)) > 0 Then
        result.Add NzString(tags)
    End If
    Set NormalizeToCollection = result
End Function

Private Function FieldToListName(ByVal tagField As String) As String
    Select Case tagField
        Case modABPhotoConstants.PHOTO_TAG_BERICHT: FieldToListName = modABPhotoConstants.PHOTO_LIST_BERICHT
        Case modABPhotoConstants.PHOTO_TAG_SEMINAR: FieldToListName = modABPhotoConstants.PHOTO_LIST_SEMINAR
        Case modABPhotoConstants.PHOTO_TAG_TOPIC: FieldToListName = modABPhotoConstants.PHOTO_LIST_TOPIC
        Case Else: FieldToListName = ""
    End Select
End Function
