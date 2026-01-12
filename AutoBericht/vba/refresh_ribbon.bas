Attribute VB_Name = "modRefreshRibbon"
Option Explicit

' Refreshes customUI/customUI.xml in the active template from embedded ribbon XML.
' Close and reopen template after running to see changes.

Public Sub AB_RefreshRibbonFromEmbedded()
    Dim xmlBase64 As String
    xmlBase64 = _
        "PGN1c3RvbVVJIHhtbG5zPSJodHRwOi8vc2NoZW1hcy5taWNyb3NvZnQuY29tL29mZmljZS8yMDA2LzAxL2N1c3RvbXVpIj4KICA8cmliYm9uPgogICAgPHRhYnM+CiAgICAgIDx0YWIgaWQ9InRhYkF1dG9CZXJpY2h0IiBsYWJlbD0iQXV0b0JlcmljaHQiPgogICAgICAgIDxncm91cCBpZD0iZ3JwQXV0b0JlcmljaHRNYWluIiBsYWJlbD0i" _
        "QXV0b0JlcmljaHQiPgogICAgICAgICAgPGJ1dHRvbiBpZD0iYnRuQXV0b0JlcmljaHRJbXBvcnRDaGFwdGVyIiBsYWJlbD0iSW1wb3J0IGNoYXB0ZXLigKYiIHNpemU9ImxhcmdlIiBpbWFnZU1zbz0iVGFibGVJbnNlcnQiIG9uQWN0aW9uPSJBQl9JbXBvcnRDaGFwdGVyRGlhbG9nIiAvPgogICAgICAgICAgPGJ1dHRvbiBpZD0iYnRuQXV0" _
        "b0JlcmljaHRJbXBvcnRBbGwiIGxhYmVsPSJJbXBvcnQgYWxsIiBzaXplPSJsYXJnZSIgaW1hZ2VNc289IlRhYmxlSW5zZXJ0RGlhbG9nIiBvbkFjdGlvbj0iQUJfSW1wb3J0QWxsIiAvPgogICAgICAgICAgPGJ1dHRvbiBpZD0iYnRuQXV0b0JlcmljaHRNYXJrZG93bkNoYXB0ZXIiIGxhYmVsPSJNYXJrZG93biBjaGFwdGVy4oCmIiBzaXpl" _
        "PSJsYXJnZSIgaW1hZ2VNc289IlRleHRUb1RhYmxlIiBvbkFjdGlvbj0iQUJfQ29udmVydE1hcmtkb3duQ2hhcHRlciIgLz4KICAgICAgICAgIDxidXR0b24gaWQ9ImJ0bkF1dG9CZXJpY2h0TWFya2Rvd25BbGwiIGxhYmVsPSJNYXJrZG93biBhbGwiIHNpemU9ImxhcmdlIiBpbWFnZU1zbz0iVGV4dFRvVGFibGVEaWFsb2ciIG9uQWN0aW9u" _
        "PSJBQl9Db252ZXJ0TWFya2Rvd25BbGwiIC8+CiAgICAgICAgICA8YnV0dG9uIGlkPSJidG5BdXRvQmVyaWNodExvZ28iIGxhYmVsPSJJbnNlcnQgbG9nbyIgc2l6ZT0ibGFyZ2UiIGltYWdlTXNvPSJQaWN0dXJlSW5zZXJ0RnJvbUZpbGUiIG9uQWN0aW9uPSJBQl9JbnNlcnRMb2dvIiAvPgogICAgICAgIDwvZ3JvdXA+CiAgICAgICAgPGdy" _
        "b3VwIGlkPSJncnBBdXRvQmVyaWNodFBwdCIgbGFiZWw9IlBvd2VyUG9pbnQiPgogICAgICAgICAgPGJ1dHRvbiBpZD0iYnRuQXV0b0JlcmljaHRQcHRSZXBvcnQiIGxhYmVsPSJQUFQgQmVyaWNodCIgc2l6ZT0ibGFyZ2UiIGltYWdlTXNvPSJTbGlkZU5ldyIgb25BY3Rpb249IkFCX0V4cG9ydFBwdFJlcG9ydCIgLz4KICAgICAgICAgIDxi" _
        "dXR0b24gaWQ9ImJ0bkF1dG9CZXJpY2h0UHB0VHJhaW5pbmciIGxhYmVsPSJWRyBTZW1pbmFyIEQiIHNpemU9ImxhcmdlIiBpbWFnZU1zbz0iUHJlc2VudGF0aW9uIiBvbkFjdGlvbj0iQUJfRXhwb3J0UHB0VHJhaW5pbmciIC8+CiAgICAgICAgICA8YnV0dG9uIGlkPSJidG5BdXRvQmVyaWNodFBwdFRyYWluaW5nRiIgbGFiZWw9IlZHIFNl" _
        "bWluYXIgRiIgc2l6ZT0ibGFyZ2UiIGltYWdlTXNvPSJQcmVzZW50YXRpb24iIG9uQWN0aW9uPSJBQl9FeHBvcnRQcHRUcmFpbmluZ0YiIC8+CiAgICAgICAgPC9ncm91cD4KICAgICAgPC90YWI+CiAgICA8L3RhYnM+CiAgPC9yaWJib24+CjwvY3VzdG9tVUk+Cg=="
    Dim xmlText As String
    xmlText = Base64ToString(xmlBase64)
    ReplaceCustomUI xmlText
    MsgBox "Ribbon updated. Save the template and reopen to see changes.", vbInformation
End Sub

Private Sub ReplaceCustomUI(ByVal xmlText As String)
    Dim pkg As OfficeCustomUI
    Set pkg = New OfficeCustomUI
    pkg.ReplaceRibbon ActiveDocument, xmlText
End Sub

Private Function Base64ToString(ByVal b64 As String) As String
    Dim bytes() As Byte
    bytes = base64Decode(b64)
    Base64ToString = StrConv(bytes, vbUnicode)
End Function

' Minimal base64 decoder
Private Function base64Decode(ByVal strData As String) As Byte()
    Dim xmlObj As Object
    Set xmlObj = CreateObject("MSXML2.DOMDocument.6.0")
    Dim node As Object
    Set node = xmlObj.createElement("b64")
    node.DataType = "bin.base64"
    node.Text = strData
    base64Decode = node.nodeTypedValue
End Function

' Helper class to replace ribbon via CustomUI part
' (lightweight wrapper around OpenXML for Word)
Private Class OfficeCustomUI
    Public Sub ReplaceRibbon(ByVal doc As Document, ByVal xmlText As String)
        Dim relType As String
        relType = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
        Dim pkg As Office.Document
        Set pkg = doc
        Dim part As Office.CustomXMLPart
        Dim found As Boolean
        For Each part In pkg.CustomXMLParts
            If part.BuiltIn Then GoTo NextPart
            If part.NamespaceURI = relType Then
                part.Delete
                found = True
            End If
NextPart:
        Next part
        pkg.CustomXMLParts.Add(xmlText)
    End Sub
End Class
