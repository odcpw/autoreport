Attribute VB_Name = "modWordMarkdown"
Option Explicit

' Minimal markdown converter for Word:
' - Lines starting with "- " become bullet list items
' - **bold** and *italic* toggles within a line
' - Blank lines become paragraph breaks

' Config constants live in modAutoBerichtConfig.

Public Sub ConvertMarkdownInSelection()
    LogDebug "ConvertMarkdownInSelection: start"
    Dim rng As Range
    Set rng = Selection.Range
    Dim cc As ContentControl
    If rng.ContentControls.Count > 0 Then
        Set cc = rng.ContentControls(1)
        ConvertMarkdownWithUnlock cc, rng
    Else
        ConvertMarkdownRange rng
    End If
End Sub

Private Sub ConvertMarkdownWithUnlock(ByVal cc As ContentControl, ByVal rng As Range)
    Dim lockContents As Boolean
    Dim lockControl As Boolean
    lockContents = cc.LockContents
    lockControl = cc.LockContentControl
    On Error Resume Next
    If lockContents Then cc.LockContents = False
    If lockControl Then cc.LockContentControl = False
    On Error GoTo 0
    ConvertMarkdownRange rng
    On Error Resume Next
    If lockContents Then cc.LockContents = True
    If lockControl Then cc.LockContentControl = True
    On Error GoTo 0
End Sub

Private Sub ConvertMarkdownRange(ByVal rng As Range)
    LogDebug "ConvertMarkdownRange: " & rng.Start & "-" & rng.End
    If rng.Tables.Count > 0 Then
        Dim cell As Cell
        For Each cell In rng.Tables(1).Range.Cells
            ConvertMarkdownCell cell
        Next cell
        Exit Sub
    End If
    ConvertMarkdownPlainRange rng
End Sub

Public Sub ConvertMarkdownChapterDialog()
    Dim chapterId As String
    chapterId = PromptChapterId(AB_PROMPT_MARKDOWN_CHAPTER)
    If Len(chapterId) = 0 Then Exit Sub
    ConvertMarkdownForChapter chapterId
End Sub

Public Sub ConvertMarkdownAll()
    Dim ids() As String
    ids = Split(AB_DEFAULT_CHAPTER_IDS, ",")
    Dim i As Long
    For i = LBound(ids) To UBound(ids)
        ConvertMarkdownForChapter Trim$(ids(i))
    Next i
    MsgBox "Markdown for all chapters completed.", vbInformation
End Sub

Public Sub ConvertMarkdownForChapter(ByVal chapterId As String)
    Dim startName As String
    Dim endName As String
    BuildChapterBookmarks chapterId, startName, endName

    Dim rng As Range
    Set rng = ResolveBookmarkRange(startName, endName)
    If rng Is Nothing Then
        LogDebug "Markdown: bookmark range missing for chapter " & chapterId
        Exit Sub
    End If
    ConvertMarkdownRange rng
End Sub

Private Sub BuildChapterBookmarks(ByVal chapterId As String, ByRef startName As String, ByRef endName As String)
    startName = "Chapter" & chapterId & "_start"
    endName = "Chapter" & chapterId & "_end"
    startName = Replace(startName, ".", "_")
    endName = Replace(endName, ".", "_")
End Sub

Private Function PromptChapterId(ByVal prompt As String) As String
    Dim input As String
    input = InputBox(prompt, AB_PROMPT_CHOOSE_CHAPTER_TITLE, AB_PROMPT_CHAPTER_DEFAULT)
    input = Trim$(input)
    If Len(input) = 0 Then Exit Function
    If Not IsValidChapterId(input) Then
        MsgBox "Invalid chapter ID: " & input, vbExclamation
        Exit Function
    End If
    PromptChapterId = input
End Function

Private Function IsValidChapterId(ByVal chapterId As String) As Boolean
    Dim ids() As String
    ids = Split(AB_DEFAULT_CHAPTER_IDS, ",")
    Dim i As Long
    For i = LBound(ids) To UBound(ids)
        If Trim$(ids(i)) = Trim$(chapterId) Then
            IsValidChapterId = True
            Exit Function
        End If
    Next i
End Function

Private Sub ConvertMarkdownCell(ByVal cell As Cell)
    LogDebug "ConvertMarkdownCell"
    On Error Resume Next
    If cell.ColumnIndex <> 2 Then Exit Sub
    Dim firstStyle As String
    firstStyle = CStr(cell.Range.Paragraphs(1).Style)
    If firstStyle = AB_STYLE_SKIP_HEADING2 Or firstStyle = AB_STYLE_SKIP_HEADING3 Then Exit Sub
    On Error GoTo 0
    Dim cellRange As Range
    Set cellRange = cell.Range
    If cellRange.End > cellRange.Start Then
        cellRange.End = cellRange.End - 1 ' remove end-of-cell marker
    End If
    ConvertMarkdownPlainRange cellRange
End Sub

Private Sub ConvertMarkdownPlainRange(ByVal rng As Range)
    LogDebug "ConvertMarkdownPlainRange"
    NormalizeRangeForCell rng
    Dim text As String
    text = rng.Text
    text = Replace(text, vbCrLf, vbLf)
    text = Replace(text, vbCr, vbLf)
    rng.Text = ""

    Dim lines() As String
    lines = Split(text, vbLf)

    Dim i As Long
    Dim writer As Range
    Set writer = rng.Duplicate
    writer.Collapse wdCollapseStart
    For i = LBound(lines) To UBound(lines)
        Dim line As String
        line = lines(i)

        If Len(Trim$(line)) = 0 Then
            If i < UBound(lines) Then
                writer.InsertParagraphAfter
                writer.Collapse wdCollapseEnd
            End If
        ElseIf Left$(line, 2) = "- " Then
            InsertMarkdownLine writer, Mid$(line, 3), True, i < UBound(lines)
        Else
            InsertMarkdownLine writer, line, False, i < UBound(lines)
        End If
    Next i
End Sub

Private Sub InsertMarkdownLine(ByVal rng As Range, ByVal line As String, ByVal asBullet As Boolean, ByVal appendBreak As Boolean)
    If asBullet Then LogDebug "InsertMarkdownLine: bullet"
    rng.Collapse wdCollapseEnd
    Dim lineStart As Long
    lineStart = rng.End

    ApplyLineStyle rng, lineStart, asBullet
    AppendFormattedText rng, line

    If appendBreak Then
        rng.InsertParagraphAfter
        rng.Collapse wdCollapseEnd
    End If
End Sub

Private Sub AppendFormattedText(ByVal rng As Range, ByVal line As String)
    Dim i As Long
    Dim boldOn As Boolean
    Dim italicOn As Boolean
    Dim buffer As String

    i = 1
    Do While i <= Len(line)
        If Mid$(line, i, 2) = "**" Then
            FlushBuffer rng, buffer, boldOn, italicOn
            buffer = ""
            boldOn = Not boldOn
            i = i + 2
        ElseIf Mid$(line, i, 1) = "*" Then
            FlushBuffer rng, buffer, boldOn, italicOn
            buffer = ""
            italicOn = Not italicOn
            i = i + 1
        Else
            buffer = buffer & Mid$(line, i, 1)
            i = i + 1
        End If
    Loop
    FlushBuffer rng, buffer, boldOn, italicOn
End Sub

Private Sub FlushBuffer(ByVal rng As Range, ByVal buffer As String, ByVal boldOn As Boolean, ByVal italicOn As Boolean)
    If Len(buffer) = 0 Then Exit Sub
    Dim part As Range
    Set part = rng.Duplicate
    part.Collapse wdCollapseEnd
    part.Text = buffer
    ApplyEmphasis part, boldOn, italicOn
    rng.SetRange part.End, part.End
End Sub

Private Function FindContentControl(ByVal title As String) As ContentControl
    Dim cc As ContentControl
    For Each cc In ActiveDocument.ContentControls
        If LCase$(cc.Title) = LCase$(title) Or LCase$(cc.Tag) = LCase$(title) Then
            Set FindContentControl = cc
            Exit Function
        End If
    Next cc
End Function

Private Function ResolveBookmarkRange(ByVal startName As String, ByVal endName As String) As Range
    Dim bmStart As Bookmark
    Dim bmEnd As Bookmark
    On Error Resume Next
    Set bmStart = ActiveDocument.Bookmarks(startName)
    Set bmEnd = ActiveDocument.Bookmarks(endName)
    On Error GoTo 0
    If bmStart Is Nothing Or bmEnd Is Nothing Then Exit Function
    Set ResolveBookmarkRange = ActiveDocument.Range(bmStart.Range.End, bmEnd.Range.Start)
End Function

Private Sub ApplyParagraphStyle(ByVal rng As Range, ByVal styleName As String, ByVal fallback As Variant)
    On Error Resume Next
    If Len(styleName) > 0 Then
        rng.Paragraphs(1).Range.Style = styleName
    Else
        rng.Paragraphs(1).Range.Style = fallback
    End If
    If Err.Number <> 0 Then
        rng.Paragraphs(1).Range.Style = fallback
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Sub ApplyEmphasis(ByVal rng As Range, ByVal boldOn As Boolean, ByVal italicOn As Boolean)
    On Error Resume Next
    If boldOn And italicOn And Len(AB_STYLE_BOLDITALIC) > 0 Then
        rng.Style = AB_STYLE_BOLDITALIC
    ElseIf boldOn And Len(AB_STYLE_BOLD) > 0 Then
        rng.Style = AB_STYLE_BOLD
    ElseIf italicOn And Len(AB_STYLE_ITALIC) > 0 Then
        rng.Style = AB_STYLE_ITALIC
    Else
        rng.Font.Bold = IIf(boldOn, True, False)
        rng.Font.Italic = IIf(italicOn, True, False)
    End If
    On Error GoTo 0
End Sub

Private Sub ApplyLineStyle(ByVal rng As Range, ByVal lineStart As Long, ByVal asBullet As Boolean)
    Dim lineRange As Range
    Set lineRange = rng.Duplicate
    lineRange.SetRange lineStart, rng.End
    If lineRange.Paragraphs.Count = 0 Then Exit Sub
    If asBullet Then
        lineRange.ListFormat.ApplyBulletDefault
        ApplyParagraphStyle lineRange, AB_STYLE_BULLET, wdStyleListBullet
    Else
        lineRange.ListFormat.RemoveNumbers NumberType:=wdNumberParagraph
        ApplyParagraphStyle lineRange, AB_STYLE_BODY, wdStyleNormal
    End If
    lineRange.ParagraphFormat.SpaceAfter = 0
End Sub

Private Sub NormalizeRangeForCell(ByVal rng As Range)
    If rng.End > rng.Start Then
        Dim tailChar As String
        tailChar = Right$(rng.Text, 1)
        If AscW(tailChar) = 7 Then
            rng.End = rng.End - 1
        End If
    End If
End Sub

Private Sub LogDebug(ByVal message As String)
    If Not AB_DEBUG_MARKDOWN Then Exit Sub
    Debug.Print Format$(Now, "hh:nn:ss") & " | " & message
End Sub
