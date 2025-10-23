'===============================================================
' AutoBericht Project JSON Export Scaffold
'---------------------------------------------------------------
' Requires:
'   - VBA-JSON (https://github.com/VBA-tools/VBA-JSON)
'   - Microsoft Scripting Runtime reference (Dictionary/Collection)
'---------------------------------------------------------------
' This module loads master/self-evaluation data from Excel ranges
' or existing JSON files, assembles the unified project structure,
' and writes a project.json ready for AutoBericht.
'===============================================================

Option Explicit

Private Const SCHEMA_VERSION As Long = 1
Private Const DEFAULT_PAGE_SIZE As Long = 5

Public Sub ExportProjectJson( _
    ByVal masterPath As String, _
    ByVal selfEvalPath As String, _
    ByVal outputPath As String)

    Dim master As Dictionary
    Dim selfEval As Dictionary
    Dim project As Dictionary

    Set master = LoadJsonFile(masterPath)
    Set selfEval = LoadJsonFile(selfEvalPath)

    Set project = BuildProjectSnapshot(master, selfEval)

    Dim jsonText As String
    jsonText = JsonConverter.ConvertToJson(project, Whitespace:=2)

    WriteTextFile outputPath, jsonText
    MsgBox "Project JSON exported to " & outputPath, vbInformation, "AutoBericht Export"
End Sub

Private Function BuildProjectSnapshot( _
    master As Dictionary, _
    selfEval As Dictionary) As Dictionary

    Dim project As New Dictionary
    project.CompareMode = TextCompare

    project("version") = SCHEMA_VERSION
    project("meta") = BuildMetaInfo(selfEval)
    project("lists") = BuildTagLists()
    project("photos") = BuildPhotoMetadata()
    project("chapters") = BuildChapterList(master, selfEval)

    Set BuildProjectSnapshot = project
End Function

Private Function BuildMetaInfo(selfEval As Dictionary) As Dictionary
    Dim meta As New Dictionary
    meta.CompareMode = TextCompare

    meta("projectId") = GetProjectId()
    meta("company") = NzString(selfEval("company"))
    meta("createdAt") = Format$(Now, "yyyy-mm-dd\THH:nn:ss\Z")
    meta("locale") = GetMetaLocale()
    meta("author") = GetCurrentUser()

    Set BuildMetaInfo = meta
End Function

Private Function BuildTagLists() As Dictionary
    Dim lists As New Dictionary
    lists.CompareMode = TextCompare

    lists("categoryList") = BuildCategoryList()
    lists("trainingList") = BuildTrainingList()

    Set BuildTagLists = lists
End Function

Private Function BuildPhotoMetadata() As Dictionary
    Dim photos As New Dictionary
    photos.CompareMode = TextCompare

    ' TODO: Populate from Excel table/range.
    ' Example:
    '   AddPhoto photos, "site/leitbild.jpg", "Neue Leitbild-Tafel",
    '            SplitTags("Gefährdung"), SplitTags("1.1.3"), SplitTags("Führungsgrundsätze")

    Set BuildPhotoMetadata = photos
End Function

Private Sub AddPhoto( _
    photos As Dictionary, _
    ByVal path As String, _
    ByVal notes As String, _
    categories As Collection, _
    chapters As Collection, _
    training As Collection)

    Dim info As New Dictionary
    info.CompareMode = TextCompare

    Dim tags As New Dictionary
    tags.CompareMode = TextCompare
    tags("categories") = CollectionToArray(categories)
    tags("chapters") = CollectionToArray(chapters)
    tags("training") = CollectionToArray(training)

    info("notes") = notes
    info("tags") = tags

    photos(path) = info
End Sub

Private Function BuildChapterList( _
    master As Dictionary, _
    selfEval As Dictionary) As Collection

    Dim output As New Collection
    Dim selfMap As Dictionary
    Set selfMap = IndexSelfEval(selfEval)

    Dim chapters As Collection
    Set chapters = master("chapters")

    Dim chapterNode As Variant
    For Each chapterNode In chapters
        Dim chapter As Dictionary
        Set chapter = New Dictionary
        chapter.CompareMode = TextCompare

        chapter("id") = chapterNode("id")
        chapter("title") = chapterNode("title")
        chapter("pageSize") = DEFAULT_PAGE_SIZE
        chapter("rows") = BuildRows(chapterNode, selfMap)

        output.Add chapter
    Next chapterNode

    Set BuildChapterList = output
End Function

Private Function BuildRows( _
    chapterNode As Dictionary, _
    selfMap As Dictionary) As Collection

    Dim rows As New Collection
    Dim children As Variant

    If chapterNode.Exists("children") Then
        For Each children In chapterNode("children")
            Dim isFinding As Boolean
            isFinding = children.Exists("findingTemplate")

            If isFinding Then
                rows.Add BuildRow(children, selfMap)
            End If

            If children.Exists("children") Then
                AppendCollections rows, BuildRows(children, selfMap)
            End If
        Next children
    End If

    Set BuildRows = rows
End Function

Private Function BuildRow( _
    node As Dictionary, _
    selfMap As Dictionary) As Dictionary

    Dim row As New Dictionary
    row.CompareMode = TextCompare

    Dim rowId As String
    rowId = node("id")

    row("id") = rowId
    row("title") = node("title")

    Dim masterBlock As New Dictionary
    masterBlock.CompareMode = TextCompare
    masterBlock("finding") = NzString(node("findingTemplate"))
    masterBlock("recommendations") = node("recommendations")
    row("master") = masterBlock

    Dim customer As New Dictionary
    customer.CompareMode = TextCompare
    If selfMap.Exists(rowId) Then
        Dim client As Dictionary
        Set client = selfMap(rowId)
        customer("answer") = client("answer")
        customer("remark") = NzString(client("remark"))
        customer("priority") = GetOptional(client, "priority")
    Else
        customer("answer") = Null
        customer("remark") = ""
        customer("priority") = Null
    End If
    row("customer") = customer

    Dim work As New Dictionary
    work.CompareMode = TextCompare
    work("selectedLevel") = 2
    work("findingOverride") = ""
    work("recommendationOverride") = ""
    work("includeFinding") = True
    work("includeRecommendation") = True
    work("overwriteMode") = "append"
    work("done") = False
    work("notes") = ""
    row("workstate") = work

    Dim assets As New Dictionary
    assets.CompareMode = TextCompare
    assets("photos") = New Collection
    assets("slides") = New Collection
    assets("documents") = New Collection
    row("assets") = assets

    row("exportHints") = New Dictionary

    Set BuildRow = row
End Function

'---------------------------------------------------------------
' Helper Routines
'---------------------------------------------------------------

Private Function LoadJsonFile(ByVal filePath As String) As Dictionary
    Dim text As String
    text = ReadTextFile(filePath)
    Set LoadJsonFile = JsonConverter.ParseJson(text)
End Function

Private Function IndexSelfEval(selfEval As Dictionary) As Dictionary
    Dim map As New Dictionary
    map.CompareMode = TextCompare

    Dim responses As Collection
    If selfEval.Exists("responses") Then
        Set responses = selfEval("responses")
    Else
        Set responses = New Collection
    End If

    Dim item As Variant
    For Each item In responses
        Dim id As String
        id = item("id")
        Dim entry As New Dictionary
        entry.CompareMode = TextCompare
        entry("answer") = GetOptional(item, "answer")
        entry("remark") = GetOptional(item, "remark")
        entry("priority") = GetOptional(item, "priority")
        map(id) = entry
    Next item

    Set IndexSelfEval = map
End Function

Private Sub AppendCollections(target As Collection, source As Collection)
    Dim item As Variant
    For Each item In source
        target.Add item
    Next item
End Sub

Private Function ReadTextFile(ByVal filePath As String) As String
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    ReadTextFile = Input$(LOF(fileNum), fileNum)
    Close #fileNum
End Function

Private Sub WriteTextFile(ByVal filePath As String, ByVal content As String)
    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, content
    Close #fileNum
End Sub

Private Function CollectionToArray(col As Collection) As Variant
    Dim arr() As Variant
    Dim i As Long
    ReDim arr(0 To col.Count - 1)
    For i = 1 To col.Count
        arr(i - 1) = col(i)
    Next i
    CollectionToArray = arr
End Function

Private Function SplitTags(ByVal value As String) As Collection
    Dim col As New Collection
    Dim parts() As String
    Dim i As Long
    parts = Split(value, ",")
    For i = LBound(parts) To UBound(parts)
        Dim trimmed As String
        trimmed = Trim$(parts(i))
        If Len(trimmed) > 0 Then
            col.Add trimmed
        End If
    Next i
    Set SplitTags = col
End Function

Private Function NzString(ByVal value As Variant) As String
    If IsNull(value) Or IsEmpty(value) Then
        NzString = ""
    Else
        NzString = CStr(value)
    End If
End Function

Private Function GetOptional(dict As Dictionary, key As String) As Variant
    If dict.Exists(key) Then
        GetOptional = dict(key)
    Else
        GetOptional = Null
    End If
End Function

Private Function GetProjectId() As String
    ' TODO: derive from workbook context.
    GetProjectId = "PROJECT-" & Format$(Now, "yyyymmdd-hhnnss")
End Function

Private Function GetMetaLocale() As String
    ' TODO: read from workbook setting.
    GetMetaLocale = "de-CH"
End Function

Private Function GetCurrentUser() As String
    GetCurrentUser = Environ$("USERNAME")
End Function

