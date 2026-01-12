Attribute VB_Name = "modWordRibbonCallbacks"
Option Explicit

' Ribbon callbacks for AutoBericht tab.

Public Sub AB_ImportReport(control As Object)
    On Error GoTo Fail
    ImportChapter0Summary
    ImportChapterAll
    Exit Sub
Fail:
    MsgBox "Import failed: " & Err.Description, vbExclamation
End Sub

Public Sub AB_ImportChapterDialog(control As Object)
    On Error GoTo Fail
    ImportChapterDialog
    Exit Sub
Fail:
    MsgBox "Import chapter failed: " & Err.Description, vbExclamation
End Sub

Public Sub AB_ImportAll(control As Object)
    On Error GoTo Fail
    ImportChapterAll
    Exit Sub
Fail:
    MsgBox "Import all failed: " & Err.Description, vbExclamation
End Sub

Public Sub AB_ConvertMarkdown(control As Object)
    On Error GoTo Fail
    ConvertMarkdownInContentControl
    Exit Sub
Fail:
    MsgBox "Markdown conversion failed: " & Err.Description, vbExclamation
End Sub

Public Sub AB_ConvertMarkdownChapter(control As Object)
    On Error GoTo Fail
    ConvertMarkdownChapterDialog
    Exit Sub
Fail:
    MsgBox "Markdown chapter failed: " & Err.Description, vbExclamation
End Sub

Public Sub AB_ConvertMarkdownAll(control As Object)
    On Error GoTo Fail
    ConvertMarkdownAll
    Exit Sub
Fail:
    MsgBox "Markdown all failed: " & Err.Description, vbExclamation
End Sub

Public Sub AB_InsertLogo(control As Object)
    On Error GoTo Fail
    InsertLogoAtToken
    Exit Sub
Fail:
    MsgBox "Logo insertion failed: " & Err.Description, vbExclamation
End Sub

Public Sub AB_ExportPptReport(control As Object)
    MsgBox "PPT Bericht export not implemented yet.", vbInformation
End Sub

Public Sub AB_ExportPptTraining(control As Object)
    On Error GoTo Fail
    ExportTrainingPptD
    Exit Sub
Fail:
    MsgBox "Training export failed: " & Err.Description, vbExclamation
End Sub

Public Sub AB_ExportPptTrainingF(control As Object)
    On Error GoTo Fail
    ExportTrainingPptF
    Exit Sub
Fail:
    MsgBox "Training export (FR) failed: " & Err.Description, vbExclamation
End Sub
