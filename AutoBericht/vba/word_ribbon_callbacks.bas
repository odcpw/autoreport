Attribute VB_Name = "modWordRibbonCallbacks"
Option Explicit

' Ribbon callbacks for AutoBericht tab.
' 8 buttons: Import (3) + Markdown (3) + Export (2)

' === IMPORT GROUP ===

Public Sub AB_ImportAll(control As Object)
    On Error GoTo Fail
    ImportChapter0Summary
    ImportChapterAll
    Exit Sub
Fail:
    MsgBox "Import all failed: " & Err.Description, vbExclamation
End Sub

Public Sub AB_ImportChapter(control As Object)
    On Error GoTo Fail
    ImportChapterDialog
    Exit Sub
Fail:
    MsgBox "Import chapter failed: " & Err.Description, vbExclamation
End Sub

Public Sub AB_ImportTextMarkers(control As Object)
    On Error GoTo Fail
    ImportTextMarkers
    Exit Sub
Fail:
    MsgBox "Text markers import failed: " & Err.Description, vbExclamation
End Sub

' === MARKDOWN GROUP ===

Public Sub AB_ConvertMarkdownAll(control As Object)
    On Error GoTo Fail
    ConvertMarkdownAll
    Exit Sub
Fail:
    MsgBox "Markdown all failed: " & Err.Description, vbExclamation
End Sub

Public Sub AB_ConvertMarkdownChapter(control As Object)
    On Error GoTo Fail
    ConvertMarkdownChapterDialog
    Exit Sub
Fail:
    MsgBox "Markdown chapter failed: " & Err.Description, vbExclamation
End Sub

Public Sub AB_ConvertMarkdownSelection(control As Object)
    On Error GoTo Fail
    ConvertMarkdownInSelection
    Exit Sub
Fail:
    MsgBox "Markdown selection failed: " & Err.Description, vbExclamation
End Sub

' === EXPORT GROUP ===

Public Sub AB_ExportPptTrainingD(control As Object)
    On Error GoTo Fail
    ExportTrainingPptD
    Exit Sub
Fail:
    MsgBox "Training export (DE) failed: " & Err.Description, vbExclamation
End Sub

Public Sub AB_ExportPptTrainingF(control As Object)
    On Error GoTo Fail
    ExportTrainingPptF
    Exit Sub
Fail:
    MsgBox "Training export (FR) failed: " & Err.Description, vbExclamation
End Sub
