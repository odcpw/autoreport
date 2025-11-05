VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPhotoTagButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public WithEvents btn As MSForms.CommandButton
Public ParentForm As PhotoSorterForm
Public TagField As String
Public TagValue As String
Public ListName As String
Public CaptionText As String
Private m_isActive As Boolean
Public Sub Initialize(ByVal formRef As PhotoSorterForm, _
                      ByVal buttonRef As MSForms.CommandButton, _
                      ByVal fieldName As String, _
                      ByVal value As String, _
                      ByVal listNameValue As String, _
                      ByVal captionValue As String)
    Set ParentForm = formRef
    Set btn = buttonRef
    TagField = fieldName
    TagValue = value
    ListName = listNameValue
    CaptionText = captionValue
    btn.Caption = captionValue
    btn.Tag = fieldName & "|" & value
    SetActive False
End Sub
Private Sub btn_Click()
    ParentForm.ToggleTag TagField, TagValue, CaptionText
    ParentForm.Controls("InvisibleTextBox").SetFocus
End Sub
Public Sub SetActive(ByVal isActive As Boolean)
    m_isActive = isActive
    If isActive Then
        btn.BackColor = RGB(220, 240, 255)
    Else
        btn.BackColor = vbButtonFace
    End If
End Sub
Public Property Get IsActive() As Boolean
    IsActive = m_isActive
End Property
Public Sub ShowCount(ByVal countValue As Long)
    btn.Caption = "(" & CStr(countValue) & ") " & CaptionText
    If countValue > 0 Then
        btn.Font.Bold = True
    ElseIf Not m_isActive Then
        btn.Font.Bold = False
    End If
End Sub
Public Sub ResetCaption()
    btn.Caption = CaptionText
    If m_isActive Then
        btn.Font.Bold = True
    Else
        btn.Font.Bold = False
    End If
    SetActive m_isActive
End Sub
