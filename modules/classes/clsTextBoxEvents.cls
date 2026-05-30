VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTextBoxEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents txt As MSForms.TextBox
Attribute txt.VB_VarHelpID = -1

Private mFieldName As String
Private mParentForm As frmItemEditor

Private mSuppressEvents As Boolean


Public Sub Initialize( _
    ByVal targetTextbox As MSForms.TextBox, _
    ByVal fieldName As String, _
    ByVal parentForm As frmItemEditor)

    Set txt = targetTextbox

    mFieldName = fieldName

    Set mParentForm = parentForm

    RestorePlaceholder

End Sub


Private Function placeholderText() As String

    placeholderText = _
        Trim$(GetPlaceholderText(mFieldName))

End Function


Private Function IsPlaceholderShowing() As Boolean

    IsPlaceholderShowing = _
        (txt.ForeColor = UI_PLACEHOLDER_FORECOLOR)

End Function


Private Sub RestorePlaceholder()

    If mSuppressEvents Then Exit Sub

    If placeholderText = "" Then Exit Sub

    If Trim$(txt.text) <> "" Then Exit Sub

    mSuppressEvents = True

    txt.ForeColor = UI_PLACEHOLDER_FORECOLOR
    txt.text = placeholderText
    txt.SelStart = 0

    mSuppressEvents = False

End Sub


Public Sub RestoreIfEmpty()

    If Trim$(txt.text) = "" Then

        RestorePlaceholder

    End If

End Sub


Private Sub RemovePlaceholder()

    If mSuppressEvents Then Exit Sub

    If Not IsPlaceholderShowing Then Exit Sub

    mSuppressEvents = True

    txt.text = ""
    txt.ForeColor = UI_NORMAL_FORECOLOR

    mSuppressEvents = False

End Sub


Private Sub txt_Enter()

    mParentForm.RestoreEmptyTextBoxPlaceholders
    mParentForm.RestoreEmptyComboBoxPlaceholders
    RemovePlaceholder

End Sub


Private Sub txt_GotFocus()

    mParentForm.RestoreEmptyTextBoxPlaceholders
    mParentForm.RestoreEmptyComboBoxPlaceholders
    RemovePlaceholder

End Sub


Private Sub txt_MouseDown( _
    ByVal Button As Integer, _
    ByVal Shift As Integer, _
    ByVal X As Single, _
    ByVal Y As Single)

    mParentForm.RestoreEmptyTextBoxPlaceholders
    mParentForm.RestoreEmptyComboBoxPlaceholders
    RemovePlaceholder

End Sub


Private Sub txt_KeyDown( _
    ByVal KeyCode As MSForms.ReturnInteger, _
    ByVal Shift As Integer)

    RemovePlaceholder

End Sub


Private Sub txt_Change()

    If mSuppressEvents Then Exit Sub

    If IsPlaceholderShowing Then

        txt.ForeColor = UI_NORMAL_FORECOLOR

    End If

End Sub


Private Sub txt_Exit( _
    ByVal Cancel As MSForms.ReturnBoolean)

    If Trim$(txt.text) = "" Then

        RestorePlaceholder

    Else

        txt.ForeColor = UI_NORMAL_FORECOLOR

    End If

    mParentForm.MarkDirty

End Sub

