Attribute VB_Name = "modABProjectValidation"
Option Explicit

'=============================================================
' Pre-flight validation for structured sheets
'=============================================================

Public Function ValidateAutoBerichtWorkbook(Optional ByVal showDialog As Boolean = True) As Boolean
    Dim errors As New Collection

    ValidateSheetHeaders SHEET_META, HeaderMeta(), errors
    ValidateSheetHeaders SHEET_CHAPTERS, HeaderChapters(), errors
    ValidateSheetHeaders SHEET_ROWS, HeaderRows(), errors
    ValidateSheetHeaders SHEET_PHOTOS, HeaderPhotos(), errors
    ValidateSheetHeaders SHEET_PHOTO_TAGS, HeaderPhotoTags(), errors
    ValidateSheetHeaders SHEET_LISTS, HeaderLists(), errors
    ValidateSheetHeaders SHEET_OVERRIDES_HISTORY, HeaderOverridesHistory(), errors

    If errors.Count = 0 Then
        ValidateAutoBerichtWorkbook = True
        Exit Function
    End If

    ValidateAutoBerichtWorkbook = False
    If Not showDialog Then Exit Function

    Dim message As String
    message = "AutoBericht workbook validation failed:" & vbCrLf & vbCrLf & JoinCollection(errors, vbCrLf)
    MsgBox message, vbExclamation, "AutoBericht Validation"
End Function

Private Sub ValidateSheetHeaders(ByVal sheetName As String, ByVal headers As Variant, errors As Collection)
    If Not SheetExists(sheetName) Then
        errors.Add "Missing sheet: " & sheetName
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)

    Dim h As Variant
    For Each h In headers
        If HeaderIndex(ws, CStr(h)) = 0 Then
            errors.Add "Missing column in " & sheetName & ": " & CStr(h)
        End If
    Next h
End Sub

Private Function JoinCollection(values As Collection, separator As String) As String
    Dim parts() As String
    ReDim parts(1 To values.Count)
    Dim i As Long
    For i = 1 To values.Count
        parts(i) = CStr(values(i))
    Next i
    JoinCollection = Join(parts, separator)
End Function

