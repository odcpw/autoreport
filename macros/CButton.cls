VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1
Public ParentForm As UserForm
Public categoryName As String

Private Sub btn_Click()
    PhotoSorterForm.ToggleCategoryAssignment categoryName, btn.caption
    ParentForm.Controls("InvisibleTextBox").SetFocus
End Sub

