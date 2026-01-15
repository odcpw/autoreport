Attribute VB_Name = "modWordRibbonCallbacks"
Option Explicit

' Ribbon callbacks for AutoBericht tab.
' 11 buttons: Import (2) + Markdown (3) + Textfelder (3) + Export (3)

' === IMPORT GROUP ===

Public Sub AB_ImportAll(control As Object)
    On Error GoTo Fail
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

' === TEXTFELDER GROUP ===

Public Sub AB_ImportTextFields(control As Object)
    On Error GoTo Fail
    ImportTextFields
    Exit Sub
Fail:
    MsgBox "Text fields import failed: " & Err.Description, vbExclamation
End Sub

Public Sub AB_InsertLogo(control As Object)
    On Error GoTo Fail
    InsertLogos
    Exit Sub
Fail:
    MsgBox "Logo insertion failed: " & Err.Description, vbExclamation
End Sub

Public Sub AB_InsertLogos(control As Object)
    On Error GoTo Fail
    InsertLogos
    Exit Sub
Fail:
    MsgBox "Logo insertion failed: " & Err.Description, vbExclamation
End Sub

Public Sub AB_InsertLogoMain(control As Object)
    On Error GoTo Fail
    InsertLogos
    Exit Sub
Fail:
    MsgBox "Main logo insertion failed: " & Err.Description, vbExclamation
End Sub

Public Sub AB_InsertLogoHeader(control As Object)
    On Error GoTo Fail
    InsertLogos
    Exit Sub
Fail:
    MsgBox "Header logo insertion failed: " & Err.Description, vbExclamation
End Sub

Public Sub AB_InsertSpider(control As Object)
    On Error GoTo Fail
    InsertSpiderChart
    Exit Sub
Fail:
    MsgBox "Spider chart insertion failed: " & Err.Description, vbExclamation
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

Public Sub AB_ExportPptBesprechung(control As Object)
    On Error GoTo Fail
    ExportBesprechungPpt
    Exit Sub
Fail:
    MsgBox "Besprechung export failed: " & Err.Description, vbExclamation
End Sub
