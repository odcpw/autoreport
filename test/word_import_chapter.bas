Attribute VB_Name = "modWordImportChapter"
Option Explicit

' Requires JsonConverter.bas (VBA-JSON) in the Word project.

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

    Dim cc As ContentControl
    Set cc = FindContentControl("Chapter1")
    If cc Is Nothing Then
        MsgBox "Content control 'Chapter1' not found.", vbExclamation
        Exit Sub
    End If

    Dim rows As Object
    Set rows = GetObject(chapter, "rows")
    If rows Is Nothing Then
        MsgBox "No rows in Chapter 1.", vbExclamation
        Exit Sub
    End If

    Dim tableRowCount As Long
    tableRowCount = CountDataRows(rows)
    If tableRowCount = 0 Then
        MsgBox "No data rows found in Chapter 1.", vbExclamation
        Exit Sub
    End If

    Dim rng As Range
    Set rng = cc.Range
    rng.Text = ""
    rng.Collapse wdCollapseStart

    Dim tbl As Table
    Set tbl = ActiveDocument.Tables.Add(rng, tableRowCount + 1, 3)
    tbl.Borders.Enable = True
    tbl.Rows(1).Range.Font.Bold = True
    tbl.Cell(1, 1).Range.Text = "ID"
    tbl.Cell(1, 2).Range.Text = "Finding"
    tbl.Cell(1, 3).Range.Text = "Recommendation"

    Dim row As Variant
    Dim targetRow As Long
    targetRow = 2

    For Each row In rows
        If IsSectionRow(row) Then
            ' skip section headers
        Else
            tbl.Cell(targetRow, 1).Range.Text = SafeText(row, "id")
            tbl.Cell(targetRow, 2).Range.Text = ResolveFinding(row)
            tbl.Cell(targetRow, 3).Range.Text = ResolveRecommendation(row)
            targetRow = targetRow + 1
        End If
    Next row

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

Private Function IsSectionRow(ByVal row As Object) As Boolean
    On Error GoTo SafeExit
    If row.Exists("kind") Then
        IsSectionRow = (LCase$(CStr(row("kind"))) = "section")
    End If
SafeExit:
End Function

Private Function CountDataRows(ByVal rows As Object) As Long
    Dim row As Variant
    Dim count As Long
    For Each row In rows
        If Not IsSectionRow(row) Then count = count + 1
    Next row
    CountDataRows = count
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
