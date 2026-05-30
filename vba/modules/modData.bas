Attribute VB_Name = "modData"
Option Explicit

' =====================================================
' WORKSHEET REFERENCES
' =====================================================

Public Function wsData() As Worksheet

    Set wsData = ThisWorkbook.Worksheets("DATA")

End Function

Public Function wsInventory() As Worksheet

    Set wsInventory = ThisWorkbook.Worksheets("INVENTORY")

End Function


' =====================================================
' ENSURE FIELD DEFINITIONS LOADED
' =====================================================

Public Sub EnsureFieldDefinitionsLoaded()

    On Error Resume Next

    If UBound(FieldDefinitions) = 0 Then
        LoadFieldDefinitions
    End If

    If Err.Number <> 0 Then

        Err.Clear

        LoadFieldDefinitions

    End If

End Sub


' =====================================================
' GET CELL VALUE
' =====================================================

Public Function GetDataValue( _
    ByVal dataRow As Long, _
    ByVal dataColumn As Long _
) As String

    GetDataValue = Trim(wsData.Cells(dataRow, dataColumn).Value)

End Function


' =====================================================
' SAVE CELL VALUE
' =====================================================

Public Sub SaveDataValue( _
    ByVal dataRow As Long, _
    ByVal dataColumn As Long, _
    ByVal newValue As String _
)

    With wsData.Cells(dataRow, dataColumn)

        .NumberFormat = "@"
        .Value = CStr(Trim(newValue))

    End With

End Sub


' =====================================================
' GET FIELD VALUE
' =====================================================

Public Function GetFieldValue( _
    ByVal dataRow As Long, _
    ByVal fieldName As String _
) As String

    Dim fd As tFieldDefinition

    EnsureFieldDefinitionsLoaded

    fd = GetFieldDefinition(fieldName)

    If fd.dataColumn <= 0 Then Exit Function

    GetFieldValue = GetDataValue( _
        dataRow, _
        fd.dataColumn _
    )

End Function


' =====================================================
' SAVE FIELD VALUE
' =====================================================

Public Sub SaveFieldValue( _
    ByVal dataRow As Long, _
    ByVal fieldName As String, _
    ByVal fieldValue As String _
)

    Dim fd As tFieldDefinition

    EnsureFieldDefinitionsLoaded

    fd = GetFieldDefinition(fieldName)

    If fd.dataColumn <= 0 Then Exit Sub

    If fd.SaveToDataSheet = False Then Exit Sub

    SaveDataValue _
        dataRow, _
        fd.dataColumn, _
        fieldValue

End Sub


' =====================================================
' BOOLEAN SAVE
' =====================================================

Public Sub SaveBooleanField( _
    ByVal frm As frmItemEditor, _
    ByVal fieldName As String, _
    ByVal dataRow As Long, _
    Optional ByVal includeNA As Boolean = False)

    Dim valueToSave As String

    Dim notesValue As String

    Dim optYes As Object
    Dim optNo As Object
    Dim optNA As Object

    Dim txtNotes As Object

    On Error Resume Next

    Set optYes = frm.Controls("fra_" & fieldName).Controls("opt_" & fieldName & "_YES")
    Set optNo = frm.Controls("fra_" & fieldName).Controls("opt_" & fieldName & "_NO")

    If includeNA = True Then
        Set optNA = frm.Controls("fra_" & fieldName).Controls("opt_" & fieldName & "_NA")
    End If

    Set txtNotes = frm.Controls("txt_" & fieldName & "_NOTES")

    valueToSave = ""

    If optYes.Value = True Then valueToSave = "YES"
    If optNo.Value = True Then valueToSave = "NO"

    If includeNA = True Then
        If optNA.Value = True Then valueToSave = "N/A"
    End If

    notesValue = Trim(txtNotes.Value)

    If IsPlaceholderValue(fieldName & "_NOTES", notesValue) Then
        notesValue = ""
    End If

    If notesValue <> "" Then
        valueToSave = valueToSave & "|" & notesValue
    End If

    SaveFieldValue _
        dataRow, _
        fieldName, _
        valueToSave

End Sub


' =====================================================
' BOOLEAN LOAD
' =====================================================

Public Sub LoadBooleanField( _
    ByVal frm As frmItemEditor, _
    ByVal fieldName As String, _
    ByVal dataRow As Long, _
    Optional ByVal includeNA As Boolean = False)

    Dim savedValue As String

    Dim valueParts() As String

    Dim boolValue As String
    Dim notesValue As String

    Dim optYes As Object
    Dim optNo As Object
    Dim optNA As Object

    Dim txtNotes As Object

    Dim placeholderText As String

    On Error Resume Next

    savedValue = GetFieldValue(dataRow, fieldName)

    Set optYes = frm.Controls("fra_" & fieldName).Controls("opt_" & fieldName & "_YES")
    Set optNo = frm.Controls("fra_" & fieldName).Controls("opt_" & fieldName & "_NO")

    If includeNA = True Then
        Set optNA = frm.Controls("fra_" & fieldName).Controls("opt_" & fieldName & "_NA")
    End If

    Set txtNotes = frm.Controls("txt_" & fieldName & "_NOTES")

    placeholderText = GetPlaceholderText(fieldName & "_NOTES")

    If savedValue = "" Then

        txtNotes.Value = placeholderText
        txtNotes.ForeColor = UI_PLACEHOLDER_FORECOLOR

        Exit Sub

    End If

    valueParts = Split(savedValue, "|")

    boolValue = Trim(valueParts(0))

    If UBound(valueParts) >= 1 Then
        notesValue = Trim(valueParts(1))
    End If

    Select Case boolValue

        Case "YES"
            optYes.Value = True

        Case "NO"
            optNo.Value = True

        Case "N/A"

            If includeNA = True Then
                optNA.Value = True
            End If

    End Select

    If notesValue = "" Then

        txtNotes.Value = placeholderText
        txtNotes.ForeColor = UI_PLACEHOLDER_FORECOLOR

    Else

        txtNotes.Value = notesValue
        txtNotes.ForeColor = UI_NORMAL_FORECOLOR

    End If

End Sub


' =====================================================
' LOAD CONTROL VALUE
' =====================================================

Public Sub LoadControlValue( _
    ByVal fieldName As String, _
    ByVal dataRow As Long _
)

    Dim ctrl As Object

    Dim savedValue As String
    Dim placeholderText As String

    Set ctrl = GetDynamicControl(fieldName)

    If ctrl Is Nothing Then Exit Sub

    savedValue = GetFieldValue(dataRow, fieldName)

    placeholderText = GetPlaceholderText(fieldName)

    If savedValue = "" Then

        If placeholderText <> "" Then

            ctrl.Value = placeholderText
            ctrl.ForeColor = UI_PLACEHOLDER_FORECOLOR

        Else

            ctrl.Value = ""
            ctrl.ForeColor = UI_NORMAL_FORECOLOR

        End If

    Else

        ctrl.Value = savedValue
        ctrl.ForeColor = UI_NORMAL_FORECOLOR

    End If

End Sub


' =====================================================
' SAVE CONTROL VALUE
' =====================================================

Public Sub SaveControlValue( _
    ByVal fieldName As String, _
    ByVal dataRow As Long _
)

    Dim ctrl As Object

    Dim valueToSave As String

    Set ctrl = GetDynamicControl(fieldName)

    If ctrl Is Nothing Then Exit Sub

    valueToSave = Trim(ctrl.Value)

    If IsPlaceholderValue(fieldName, valueToSave) Then
        valueToSave = ""
    End If

    If TypeName(ctrl) = "ComboBox" Then
        If IsRestrictedLookupValueValid(fieldName, valueToSave) = False Then
            valueToSave = ""
            ctrl.Value = ""
            ctrl.ForeColor = UI_NORMAL_FORECOLOR
        End If
        PersistComboBoxLookupValue fieldName, valueToSave
    End If

    SaveFieldValue _
        dataRow, _
        fieldName, _
        valueToSave

    ' =============================================
    ' IMMEDIATELY RESTORE PLACEHOLDER AFTER SAVE
    ' =============================================

    If valueToSave = "" Then

        If GetPlaceholderText(fieldName) <> "" Then

            ctrl.Value = GetPlaceholderText(fieldName)
            ctrl.ForeColor = UI_PLACEHOLDER_FORECOLOR

        Else

            ctrl.Value = ""
            ctrl.ForeColor = UI_NORMAL_FORECOLOR

        End If

    End If

End Sub


' =====================================================
' LOAD ALL FIELD VALUES
' =====================================================

Public Sub LoadAllFieldValues(frm As frmItemEditor)

    Dim i As Long

    Dim fieldName As String

    EnsureFieldDefinitionsLoaded

    For i = LBound(FieldDefinitions) To UBound(FieldDefinitions)

        fieldName = FieldDefinitions(i).fieldName

        Select Case FieldDefinitions(i).ControlStyle

            Case STANDARD_TEXT
                LoadControlValue fieldName, frm.CurrentDataRow

            Case MULTILINE_TEXT
                LoadControlValue fieldName, frm.CurrentDataRow

            Case EDITABLE_DROPDOWN
                LoadControlValue fieldName, frm.CurrentDataRow

            Case BOOLEAN_WITH_NOTES

                LoadBooleanField _
                    frm, _
                    fieldName, _
                    frm.CurrentDataRow, _
                    False

            Case BOOLEAN_WITH_NOTES_NA

                LoadBooleanField _
                    frm, _
                    fieldName, _
                    frm.CurrentDataRow, _
                    True

        End Select

    Next i

End Sub


' =====================================================
' SAVE ALL FIELD VALUES
' =====================================================

Public Sub SaveAllFieldValues(frm As frmItemEditor)

    Dim i As Long

    Dim fieldName As String

    EnsureFieldDefinitionsLoaded

    For i = LBound(FieldDefinitions) To UBound(FieldDefinitions)

        fieldName = FieldDefinitions(i).fieldName

        Select Case FieldDefinitions(i).ControlStyle

            Case STANDARD_TEXT
                SaveControlValue fieldName, frm.CurrentDataRow

            Case MULTILINE_TEXT
                SaveControlValue fieldName, frm.CurrentDataRow

            Case EDITABLE_DROPDOWN
                SaveControlValue fieldName, frm.CurrentDataRow

            Case BOOLEAN_WITH_NOTES

                SaveBooleanField _
                    frm, _
                    fieldName, _
                    frm.CurrentDataRow, _
                    False

            Case BOOLEAN_WITH_NOTES_NA

                SaveBooleanField _
                    frm, _
                    fieldName, _
                    frm.CurrentDataRow, _
                    True

        End Select

    Next i

End Sub

