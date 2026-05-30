VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPhotoButtonEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1

Private mParentForm As frmItemEditor
Private mActionName As String
Private mPhotoIndex As Long

Public Sub Initialize(ByVal targetButton As MSForms.CommandButton, ByVal parentForm As frmItemEditor, ByVal actionName As String, Optional ByVal photoIndex As Long = 0)
    Set btn = targetButton
    Set mParentForm = parentForm
    mActionName = actionName
    mPhotoIndex = photoIndex
End Sub

Private Sub btn_Click()
    Select Case mActionName
        Case "SELECT"
            mParentForm.SelectPhotos
        Case "REMOVE_ALL"
            mParentForm.RemoveAllPhotos
        Case "REMOVE_ONE"
            mParentForm.RemovePhotoAt mPhotoIndex
    End Select
End Sub
