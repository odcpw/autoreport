Attribute VB_Name = "modWordMarkdown"
Option Explicit

' Minimal markdown converter for Word:
' - Lines starting with "- " become bullet list items
' - **bold** and *italic* toggles within a line
' - Blank lines become paragraph breaks

' === STYLE CONFIG (edit these to match your template) ===
' Use custom style names if possible (same names across DE/FR/IT templates).
' Leave blank to use built-in styles via wdStyle* constants (language-safe).
Private Const STYLE_BODY As String = "Normal"
Private Const STYLE_BULLET As String = "List Paragraph"
' Optional character styles for markdown emphasis. Leave blank to use direct bold/italic.
Private Const STYLE_BOLD As String = ""
Private Const STYLE_ITALIC As String = ""
Private Const STYLE_BOLDITALIC As String = ""

Public Sub ConvertMarkdownInSelection()
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

Public Sub ConvertMarkdownInContentControl()
    Dim rng As Range
    Set rng = ResolveBookmarkRange("Chapter1_start", "Chapter1_end")
    If Not rng Is Nothing Then
        ConvertMarkdownRange rng
        Exit Sub
    End If

    Dim cc As ContentControl
    Set cc = FindContentControl("Chapter1")
    If cc Is Nothing Then
        MsgBox "Content control or bookmark for Chapter1 not found.", vbExclamation
        Exit Sub
    End If
    ConvertMarkdownWithUnlock cc, cc.Range
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
    If rng.Tables.Count > 0 Then
        Dim cell As Cell
        For Each cell In rng.Tables(1).Range.Cells
            ConvertMarkdownCell cell
        Next cell
        Exit Sub
    End If
    ConvertMarkdownPlainRange rng
End Sub

Private Sub ConvertMarkdownCell(ByVal cell As Cell)
    Dim cellRange As Range
    Set cellRange = cell.Range
    If cellRange.End > cellRange.Start Then
        cellRange.End = cellRange.End - 1 ' remove end-of-cell marker
    End If
    ConvertMarkdownPlainRange cellRange
End Sub

Private Sub ConvertMarkdownPlainRange(ByVal rng As Range)
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
    rng.Collapse wdCollapseEnd
    Dim lineStart As Long
    lineStart = rng.End

    AppendFormattedText rng, line
    ApplyLineStyle rng, lineStart, asBullet

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
    If boldOn And italicOn And Len(STYLE_BOLDITALIC) > 0 Then
        rng.Style = STYLE_BOLDITALIC
    ElseIf boldOn And Len(STYLE_BOLD) > 0 Then
        rng.Style = STYLE_BOLD
    ElseIf italicOn And Len(STYLE_ITALIC) > 0 Then
        rng.Style = STYLE_ITALIC
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
        ApplyParagraphStyle lineRange, STYLE_BULLET, wdStyleListBullet
    Else
        lineRange.ListFormat.RemoveNumbers NumberType:=wdNumberParagraph
        ApplyParagraphStyle lineRange, STYLE_BODY, wdStyleNormal
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
