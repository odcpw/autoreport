Attribute VB_Name = "modWordMarkdown"
Option Explicit

' Minimal markdown converter for Word:
' - Lines starting with "- " become bullet list items
' - **bold** and *italic* toggles within a line
' - Blank lines become paragraph breaks

' === STYLE CONFIG (edit these to match your template) ===
Private Const STYLE_BODY As String = "BodyText"
Private Const STYLE_BULLET As String = "ListBullet"

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
    Dim cc As ContentControl
    Set cc = FindContentControl("Chapter1")
    If cc Is Nothing Then
        MsgBox "Content control 'Chapter1' not found.", vbExclamation
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
            writer.InsertParagraphAfter
            writer.Collapse wdCollapseEnd
        ElseIf Left$(line, 2) = "- " Then
            InsertMarkdownLine writer, Mid$(line, 3), True
        Else
            InsertMarkdownLine writer, line, False
        End If
    Next i
End Sub

Private Sub InsertMarkdownLine(ByVal rng As Range, ByVal line As String, ByVal asBullet As Boolean)
    rng.Collapse wdCollapseEnd
    Dim lineStart As Long
    lineStart = rng.End

    AppendFormattedText rng, line

    Dim lineRange As Range
    Set lineRange = rng.Duplicate
    lineRange.SetRange lineStart, rng.End
    lineRange.ParagraphFormat.SpaceAfter = 0

    If asBullet Then
        lineRange.ListFormat.ApplyBulletDefault
        ApplyParagraphStyle lineRange, STYLE_BULLET
    Else
        lineRange.ListFormat.RemoveNumbers NumberType:=wdNumberParagraph
        ApplyParagraphStyle lineRange, STYLE_BODY
    End If

    rng.InsertParagraphAfter
    rng.Collapse wdCollapseEnd
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
    part.Font.Bold = IIf(boldOn, True, False)
    part.Font.Italic = IIf(italicOn, True, False)
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

Private Sub ApplyParagraphStyle(ByVal rng As Range, ByVal styleName As String)
    On Error Resume Next
    rng.Paragraphs(1).Range.Style = styleName
    On Error GoTo 0
End Sub
