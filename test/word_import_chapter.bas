Attribute VB_Name = "modWordImportChapter"
Option Explicit

' Requires JsonConverter.bas (VBA-JSON) in the Word project.

' === STYLE CONFIG (edit these to match your template) ===
Private Const STYLE_BODY As String = "BodyText"
Private Const STYLE_SECTION As String = "BodyText"

Public Sub ImportChapter1Table()
    Dim jsonPath As String
    jsonPath = ResolveSidecarPath()
    If Len(jsonPath) = 0 Then Exit Sub

    Dim jsonText As String
    jsonText = ReadAllText(jsonPath)
    If Len(jsonText) = 0 Then
        MsgBox "Sidecar JSON is empty.", vbExclamation
        Exit Sub
    End If

    Dim root As Object
    Set root = JsonConverter.ParseJson(jsonText)

    Dim report As Object
    Set report = GetObject(root, "report")
    If report Is Nothing Then
        MsgBox "Missing report section in JSON.", vbExclamation
        Exit Sub
    End If

    Dim project As Object
    Set project = GetObject(report, "project")
    If project Is Nothing Then
        MsgBox "Missing report.project in JSON.", vbExclamation
        Exit Sub
    End If

    Dim chapters As Object
    Set chapters = GetObject(project, "chapters")
    If chapters Is Nothing Or chapters.Count = 0 Then
        MsgBox "No chapters found in JSON.", vbExclamation
        Exit Sub
    End If

    Dim chapter As Object
    Set chapter = chapters.Item(1)

    Dim insertRng As Range
    Set insertRng = ResolveInsertRange("Chapter1")
    If insertRng Is Nothing Then
        MsgBox "Content control or bookmark 'Chapter1' not found.", vbExclamation
        Exit Sub
    End If

    Dim rows As Object
    Set rows = GetObject(chapter, "rows")
    If rows Is Nothing Then
        MsgBox "No rows in Chapter 1.", vbExclamation
        Exit Sub
    End If

    Dim includedSections As Object
    Set includedSections = BuildIncludedSections(rows)

    Dim tableRowCount As Long
    tableRowCount = CountDataRows(rows, includedSections)
    If tableRowCount = 0 Then
        MsgBox "No data rows found in Chapter 1.", vbExclamation
        Exit Sub
    End If

    insertRng.Text = ""
    insertRng.Collapse wdCollapseStart

    Dim tbl As Table
    Set tbl = ActiveDocument.Tables.Add(insertRng, tableRowCount + 1, 4)
    tbl.Borders.Enable = True
    tbl.Rows(1).Range.Font.Bold = True
    tbl.Cell(1, 1).Range.Text = "ID"
    tbl.Cell(1, 2).Range.Text = "Finding"
    tbl.Cell(1, 3).Range.Text = "Recommendation"
    tbl.Cell(1, 4).Range.Text = "Priority"

    Dim row As Variant
    Dim targetRow As Long
    targetRow = 2

    For Each row In rows
        If IsSectionRow(row) Then
            If ShouldIncludeSection(row, includedSections) Then
                On Error Resume Next
                tbl.Cell(targetRow, 1).Merge tbl.Cell(targetRow, 4)
                On Error GoTo 0
                tbl.Cell(targetRow, 1).Range.Text = SafeSectionTitle(row)
                tbl.Rows(targetRow).Range.Font.Bold = True
                tbl.Rows(targetRow).Range.Style = STYLE_SECTION
                targetRow = targetRow + 1
            End If
        ElseIf IsIncludedRow(row) Then
            tbl.Cell(targetRow, 1).Range.Text = SafeText(row, "id")
            tbl.Cell(targetRow, 2).Range.Text = ResolveFinding(row)
            tbl.Cell(targetRow, 3).Range.Text = ResolveRecommendation(row)
            tbl.Cell(targetRow, 4).Range.Text = ""
            tbl.Rows(targetRow).Range.Style = STYLE_BODY
            targetRow = targetRow + 1
        End If
    Next row

    On Error Resume Next
    tbl.Columns(4).PreferredWidthType = wdPreferredWidthPoints
    tbl.Columns(4).PreferredWidth = CentimetersToPoints(1.2)
    On Error GoTo 0
    tbl.AutoFitBehavior wdAutoFitContent

    MsgBox "Chapter 1 table imported.", vbInformation
End Sub

Private Function ResolveSidecarPath() As String
    Dim defaultPath As String
    If Len(ActiveDocument.Path) > 0 Then
        defaultPath = ActiveDocument.Path & "\\project_sidecar.json"
        If FileExists(defaultPath) Then
            ResolveSidecarPath = defaultPath
            Exit Function
        End If
    End If

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select project_sidecar.json"
    fd.Filters.Clear
    fd.Filters.Add "JSON", "*.json"
    fd.AllowMultiSelect = False
    If fd.Show <> -1 Then Exit Function
    ResolveSidecarPath = fd.SelectedItems(1)
End Function

Private Function ReadAllText(ByVal path As String) As String
    Dim text As String
    text = ReadAllTextUtf8(path)
    If Len(text) > 0 Then
        ReadAllText = text
        Exit Function
    End If
    Dim f As Integer
    f = FreeFile
    On Error GoTo CleanFail
    Open path For Input As #f
    ReadAllText = Input$(LOF(f), f)
CleanExit:
    Close #f
    Exit Function
CleanFail:
    ReadAllText = ""
    Resume CleanExit
End Function

Private Function ReadAllTextUtf8(ByVal path As String) As String
    On Error GoTo CleanFail
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' text
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile path
    ReadAllTextUtf8 = stream.ReadText(-1)
    stream.Close
    Exit Function
CleanFail:
    ReadAllTextUtf8 = ""
End Function

Private Function FileExists(ByVal path As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(path) <> "")
End Function

Private Function FindContentControl(ByVal title As String) As ContentControl
    Dim cc As ContentControl
    For Each cc In ActiveDocument.ContentControls
        If LCase$(cc.Title) = LCase$(title) Or LCase$(cc.Tag) = LCase$(title) Then
            Set FindContentControl = cc
            Exit Function
        End If
    Next cc
End Function

Private Function FindBookmark(ByVal name As String) As Bookmark
    On Error Resume Next
    Set FindBookmark = ActiveDocument.Bookmarks(name)
    On Error GoTo 0
End Function

Private Function ResolveInsertRange(ByVal anchorName As String) As Range
    Dim cc As ContentControl
    Set cc = FindContentControl(anchorName)
    If Not cc Is Nothing Then
        Set ResolveInsertRange = cc.Range
        Exit Function
    End If

    Dim bm As Bookmark
    Set bm = FindBookmark(anchorName)
    If bm Is Nothing Then Exit Function

    Dim anchorPara As Paragraph
    Set anchorPara = bm.Range.Paragraphs(1)

    If anchorPara.Next Is Nothing Then
        anchorPara.Range.InsertParagraphAfter
    End If

    Dim blankPara As Paragraph
    Set blankPara = anchorPara.Next
    If Len(Trim$(Replace(blankPara.Range.Text, vbCr, ""))) > 0 Then
        blankPara.Range.InsertParagraphBefore
        Set blankPara = anchorPara.Next
    End If

    If blankPara.Next Is Nothing Then
        blankPara.Range.InsertParagraphAfter
    End If

    Dim targetPara As Paragraph
    Set targetPara = blankPara.Next
    Set ResolveInsertRange = targetPara.Range
End Function

Private Function IsSectionRow(ByVal row As Object) As Boolean
    On Error GoTo SafeExit
    If row.Exists("kind") Then
        IsSectionRow = (LCase$(CStr(row("kind"))) = "section")
    End If
SafeExit:
End Function

Private Function CountDataRows(ByVal rows As Object, ByVal includedSections As Object) As Long
    Dim row As Variant
    Dim count As Long
    For Each row In rows
        If IsSectionRow(row) Then
            If ShouldIncludeSection(row, includedSections) Then count = count + 1
        ElseIf IsIncludedRow(row) Then
            count = count + 1
        End If
    Next row
    CountDataRows = count
End Function

Private Function BuildIncludedSections(ByVal rows As Object) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim row As Variant
    For Each row In rows
        If Not IsSectionRow(row) Then
            If IsIncludedRow(row) Then
                Dim sectionKey As String
                sectionKey = ""
                On Error Resume Next
                If row.Exists("sectionId") Then sectionKey = CStr(row("sectionId"))
                On Error GoTo 0
                If Len(sectionKey) > 0 Then dict(sectionKey) = True
            End If
        End If
    Next row
    Set BuildIncludedSections = dict
End Function

Private Function ShouldIncludeSection(ByVal row As Object, ByVal includedSections As Object) As Boolean
    Dim key As String
    key = ""
    On Error Resume Next
    If row.Exists("id") Then key = CStr(row("id"))
    On Error GoTo 0
    If Len(key) = 0 Then Exit Function
    If includedSections.Exists(key) Then ShouldIncludeSection = True
End Function

Private Function SafeSectionTitle(ByVal row As Object) As String
    Dim title As String
    title = ""
    On Error Resume Next
    If row.Exists("title") Then title = CStr(row("title"))
    On Error GoTo 0
    If Len(title) = 0 Then
        SafeSectionTitle = SafeText(row, "id")
    Else
        SafeSectionTitle = title
    End If
End Function

Private Function IsIncludedRow(ByVal row As Object) As Boolean
    Dim ws As Object
    Set ws = GetObject(row, "workstate")
    If ws Is Nothing Then
        IsIncludedRow = True
        Exit Function
    End If
    On Error Resume Next
    If ws.Exists("includeFinding") Then
        IsIncludedRow = CBool(ws("includeFinding"))
    Else
        IsIncludedRow = True
    End If
    On Error GoTo 0
End Function

Private Function ResolveFinding(ByVal row As Object) As String
    Dim ws As Object
    Set ws = GetObject(row, "workstate")
    If Not ws Is Nothing Then
        If GetBool(ws, "useFindingOverride") Then
            ResolveFinding = SafeText(ws, "findingOverride")
            Exit Function
        End If
    End If

    Dim master As Object
    Set master = GetObject(row, "master")
    If Not master Is Nothing Then
        ResolveFinding = ToPlainText(master("finding"))
    End If
End Function

Private Function ResolveRecommendation(ByVal row As Object) As String
    Dim levelKey As String
    levelKey = "1"

    Dim ws As Object
    Set ws = GetObject(row, "workstate")
    If Not ws Is Nothing Then
        If ws.Exists("includeRecommendation") Then
            If Not CBool(ws("includeRecommendation")) Then
                ResolveRecommendation = ""
                Exit Function
            End If
        End If
        If ws.Exists("selectedLevel") Then
            levelKey = CStr(ws("selectedLevel"))
        End If
        If ws.Exists("useLevelOverride") Then
            Dim overrides As Object
            Set overrides = GetObject(ws, "levelOverrides")
            If Not overrides Is Nothing Then
                If GetBoolFromDict(ws("useLevelOverride"), levelKey) Then
                    ResolveRecommendation = ToPlainText(overrides(levelKey))
                    Exit Function
                End If
            End If
        End If
    End If

    Dim master As Object
    Set master = GetObject(row, "master")
    If Not master Is Nothing Then
        Dim levels As Object
        Set levels = GetObject(master, "levels")
        If Not levels Is Nothing Then
            If levels.Exists(levelKey) Then
                ResolveRecommendation = ToPlainText(levels(levelKey))
                Exit Function
            End If
        End If
    End If
End Function

Private Function GetObject(ByVal dict As Object, ByVal key As String) As Object
    On Error GoTo SafeExit
    If dict Is Nothing Then Exit Function
    If dict.Exists(key) Then
        If IsObject(dict(key)) Then
            Set GetObject = dict(key)
        End If
    End If
SafeExit:
End Function

Private Function SafeText(ByVal dict As Object, ByVal key As String) As String
    On Error GoTo SafeExit
    If dict.Exists(key) Then
        SafeText = ToPlainText(dict(key))
    End If
SafeExit:
End Function

Private Function ToPlainText(ByVal value As Variant) As String
    If IsObject(value) Then
        If TypeName(value) = "Collection" Then
            Dim parts As Collection
            Set parts = value
            Dim item As Variant
            Dim buff As String
            For Each item In parts
                If Len(buff) > 0 Then buff = buff & vbCrLf
                buff = buff & CStr(item)
            Next item
            ToPlainText = buff
            Exit Function
        End If
    End If
    If IsNull(value) Then
        ToPlainText = ""
    Else
        ToPlainText = CStr(value)
    End If
End Function

Private Function GetBool(ByVal dict As Object, ByVal key As String) As Boolean
    On Error GoTo SafeExit
    If dict.Exists(key) Then
        GetBool = CBool(dict(key))
    End If
SafeExit:
End Function

Private Function GetBoolFromDict(ByVal dict As Variant, ByVal key As String) As Boolean
    On Error GoTo SafeExit
    If IsObject(dict) Then
        If dict.Exists(key) Then
            GetBoolFromDict = CBool(dict(key))
        End If
    End If
SafeExit:
End Function
