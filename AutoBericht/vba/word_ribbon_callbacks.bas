Attribute VB_Name = "modWordRibbonCallbacks"
Option Explicit

' Ribbon callbacks for AutoBericht tab.

Public Sub AB_ImportReport(control As Object)
    On Error GoTo Fail
    ImportChapter0Summary
    ImportChapter1Table
    Exit Sub
Fail:
    MsgBox "Import failed: " & Err.Description, vbExclamation
End Sub

Public Sub AB_ConvertMarkdown(control As Object)
    On Error GoTo Fail
    ConvertMarkdownInContentControl
    Exit Sub
Fail:
    MsgBox "Markdown conversion failed: " & Err.Description, vbExclamation
End Sub
