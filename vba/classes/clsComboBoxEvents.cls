VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsComboBoxEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents cbo As MSForms.ComboBox
Attribute cbo.VB_VarHelpID = -1

Private mFieldName As String
Private mParentForm As frmItemEditor

Private mSuppressEvents As Boolean


Public Sub Initialize( _
    ByVal targetComboBox As MSForms.ComboBox, _
    ByVal fieldName As String, _
    ByVal parentForm As frmItemEditor)

    Set cbo = targetComboBox

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
        (cbo.ForeColor = UI_PLACEHOLDER_FORECOLOR)

End Function


Private Sub RestorePlaceholder()

    If mSuppressEvents Then Exit Sub

    If placeholderText = "" Then Exit Sub

    If Trim$(cbo.text) <> "" Then Exit Sub

    mSuppressEvents = True

    cbo.ForeColor = UI_PLACEHOLDER_FORECOLOR
    cbo.text = placeholderText

    mSuppressEvents = False

End Sub


Public Sub RestoreIfEmpty()

    If Trim$(cbo.text) = "" Then

        RestorePlaceholder

    End If

End Sub


Private Sub RemovePlaceholder()

    If mSuppressEvents Then Exit Sub

    If Not IsPlaceholderShowing Then Exit Sub

    mSuppressEvents = True

    cbo.text = ""
    cbo.ForeColor = UI_NORMAL_FORECOLOR

    mSuppressEvents = False

End Sub


Private Sub cbo_Enter()

    mParentForm.RestoreEmptyTextBoxPlaceholders
    mParentForm.RestoreEmptyComboBoxPlaceholders
    RemovePlaceholder

End Sub


Private Sub cbo_GotFocus()

    mParentForm.RestoreEmptyTextBoxPlaceholders
    mParentForm.RestoreEmptyComboBoxPlaceholders
    RemovePlaceholder

End Sub


Private Sub cbo_DropButtonClick()

    mParentForm.RestoreEmptyTextBoxPlaceholders
    mParentForm.RestoreEmptyComboBoxPlaceholders
    RemovePlaceholder

End Sub


Private Sub cbo_KeyDown( _
    ByVal KeyCode As MSForms.ReturnInteger, _
    ByVal Shift As Integer)

    RemovePlaceholder

End Sub


Private Sub cbo_Change()

    If mSuppressEvents Then Exit Sub

    If Trim$(cbo.text) = "" Then

        RestorePlaceholder
        Exit Sub

    End If

    If IsLookupAutoLearnEnabled(mFieldName) = False Then

        If IsRestrictedLookupPrefixValid(mFieldName, cbo.text) = False Then

            mSuppressEvents = True
            cbo.text = ""
            mSuppressEvents = False
            RestorePlaceholder
            Exit Sub

        End If

    End If

    cbo.ForeColor = UI_NORMAL_FORECOLOR

End Sub


Private Sub cbo_Exit( _
    ByVal Cancel As MSForms.ReturnBoolean)

    If Trim$(cbo.text) = "" Then

        RestorePlaceholder

    Else

        If IsRestrictedLookupValueValid(mFieldName, cbo.text) = False Then

            mSuppressEvents = True
            cbo.text = ""
            mSuppressEvents = False
            RestorePlaceholder
            mParentForm.MarkDirty
            Exit Sub

        End If

        PersistComboBoxLookupValue mFieldName, cbo.text
        cbo.ForeColor = UI_NORMAL_FORECOLOR

    End If

    mParentForm.MarkDirty

End Sub

