Attribute VB_Name = "modABIdUtils"
Option Explicit

'=============================================================
' Helpers for working with hierarchical report item identifiers
'=============================================================

Public Function NormalizeReportItemId(ByVal rawValue As String) As String
    Dim t As String
    t = Trim$(rawValue)
    Do While Len(t) > 0 And (Right$(t, 1) = "." Or Right$(t, 1) = " ")
        t = Left$(t, Len(t) - 1)
    Loop
    Do While InStr(t, "..") > 0
        t = Replace$(t, "..", ".")
    Loop
    NormalizeReportItemId = t
End Function

Public Function IsValidReportItemId(ByVal candidate As String) As Boolean
    Dim s As String
    s = NormalizeReportItemId(candidate)
    If Len(s) = 0 Then Exit Function

    Dim parts() As String
    parts = Split(s, ".")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim seg As String
        seg = Trim$(parts(i))
        If Len(seg) = 0 Then Exit Function
        If Not seg Like "*[0-9]*" Then Exit Function
    Next i
    IsValidReportItemId = True
End Function

Public Function ParentChapterId(ByVal childId As String) As String
    Dim cleaned As String
    cleaned = NormalizeReportItemId(childId)
    If Len(cleaned) = 0 Then Exit Function

    Dim pos As Long
    pos = InStrRev(cleaned, ".")
    If pos <= 0 Then
        ParentChapterId = ""
    Else
        ParentChapterId = Left$(cleaned, pos - 1)
    End If
End Function
