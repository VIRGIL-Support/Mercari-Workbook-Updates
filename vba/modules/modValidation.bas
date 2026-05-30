Attribute VB_Name = "modValidation"
Option Explicit

Private Function GetControlValueForValidation(ByVal fieldName As String) As String
    Dim ctrl As Object
    Set ctrl = GetDynamicControl(fieldName)
    If ctrl Is Nothing Then Exit Function
    GetControlValueForValidation = Trim$(CStr(ctrl.Value))
    If IsPlaceholderValue(fieldName, GetControlValueForValidation) Then
        GetControlValueForValidation = ""
    End If
End Function

Private Function IsWholeNumberText(ByVal valueText As String) As Boolean
    Dim i As Long
    valueText = Trim$(valueText)
    If valueText = "" Then Exit Function
    For i = 1 To Len(valueText)
        If Mid$(valueText, i, 1) < "0" Or Mid$(valueText, i, 1) > "9" Then Exit Function
    Next i
    IsWholeNumberText = True
End Function

Private Function IsPositiveNumericText(ByVal valueText As String) As Boolean
    If Trim$(valueText) = "" Then
        IsPositiveNumericText = True
        Exit Function
    End If
    If IsNumeric(valueText) = False Then Exit Function
    If CDbl(valueText) < 0 Then Exit Function
    IsPositiveNumericText = True
End Function

Public Function GetFieldValidationMessage(ByVal fieldName As String, ByVal valueText As String) As String
    valueText = Trim$(CStr(valueText))
    If IsPlaceholderValue(fieldName, valueText) Then valueText = ""
    Select Case UCase$(Trim$(fieldName))
        Case "UPC"
            If valueText <> "" Then
                If IsWholeNumberText(valueText) = False Then
                    GetFieldValidationMessage = "UPC must contain numbers only."
                ElseIf Len(valueText) <> 8 And Len(valueText) <> 12 And Len(valueText) <> 13 And Len(valueText) <> 14 Then
                    GetFieldValidationMessage = "UPC should be 8, 12, 13, or 14 digits."
                End If
            End If
        Case "ITEM_HEIGHT", "ITEM_WIDTH", "ITEM_LENGTH", "TOTAL_CONTENT_HEIGHT", "TOTAL_CONTENT_WIDTH", "TOTAL_CONTENT_LENGTH", "WEIGHT_POUNDS"
            If IsPositiveNumericText(valueText) = False Then
                GetFieldValidationMessage = GetValidationFieldLabel(fieldName) & " must be a positive number."
            End If
        Case "WEIGHT_OUNCES"
            If IsPositiveNumericText(valueText) = False Then
                GetFieldValidationMessage = "Weight ounces must be a positive number."
            ElseIf valueText <> "" Then
                If CDbl(valueText) >= 16 Then
                    GetFieldValidationMessage = "Weight ounces must be less than 16."
                End If
            End If
    End Select
End Function

Private Function GetValidationFieldLabel(ByVal fieldName As String) As String
    Select Case UCase$(Trim$(fieldName))
        Case "ITEM_HEIGHT"
            GetValidationFieldLabel = "Item width"
        Case "ITEM_WIDTH"
            GetValidationFieldLabel = "Item height"
        Case "ITEM_LENGTH"
            GetValidationFieldLabel = "Item length"
        Case "TOTAL_CONTENT_HEIGHT"
            GetValidationFieldLabel = "Total content width"
        Case "TOTAL_CONTENT_WIDTH"
            GetValidationFieldLabel = "Total content height"
        Case "TOTAL_CONTENT_LENGTH"
            GetValidationFieldLabel = "Total content length"
        Case "WEIGHT_POUNDS"
            GetValidationFieldLabel = "Weight pounds"
        Case Else
            GetValidationFieldLabel = fieldName
    End Select
End Function

Private Sub AddValidationMessage(ByRef validationMessages As Collection, ByVal messageText As String)
    validationMessages.Add messageText
End Sub

Private Function GetFieldDefinitionIndex(ByVal fieldName As String) As Long
    Dim i As Long
    EnsureFieldDefinitionsLoaded
    For i = LBound(FieldDefinitions) To UBound(FieldDefinitions)
        If UCase$(Trim$(FieldDefinitions(i).fieldName)) = UCase$(Trim$(fieldName)) Then
            GetFieldDefinitionIndex = i
            Exit Function
        End If
    Next i
End Function

Private Function GetFieldTabName(ByVal fieldName As String) As String
    Dim fieldIndex As Long
    fieldIndex = GetFieldDefinitionIndex(fieldName)
    If fieldIndex > 0 Then
        GetFieldTabName = FieldDefinitions(fieldIndex).tabName
    End If
End Function

Private Function GetFieldDisplayLabel(ByVal fieldName As String, Optional ByVal fallbackLabel As String = "") As String
    Dim fieldIndex As Long
    fieldIndex = GetFieldDefinitionIndex(fieldName)
    If fieldIndex > 0 Then
        GetFieldDisplayLabel = FieldDefinitions(fieldIndex).fieldLabel
    Else
        GetFieldDisplayLabel = fallbackLabel
    End If
End Function

Private Function BuildFieldValidationMessage(ByVal fieldName As String, ByVal fieldLabel As String, ByVal reasonText As String) As String
    Dim tabName As String
    tabName = GetFieldTabName(fieldName)
    If fieldLabel = "" Then fieldLabel = GetFieldDisplayLabel(fieldName, fieldName)
    If tabName <> "" Then
        BuildFieldValidationMessage = "Click the " & tabName & " tab and update the " & fieldLabel & " field to resolve this error: " & reasonText
    Else
        BuildFieldValidationMessage = "Update the " & fieldLabel & " field to resolve this error: " & reasonText
    End If
End Function

Private Sub ValidateRequiredField(ByRef validationMessages As Collection, ByVal fieldName As String, ByVal fieldLabel As String)
    If GetControlValueForValidation(fieldName) = "" Then
        AddValidationMessage validationMessages, BuildFieldValidationMessage(fieldName, fieldLabel, fieldLabel & " is required.")
    End If
End Sub

Private Sub ValidateUPC(ByRef validationMessages As Collection)
    Dim valueText As String
    valueText = GetControlValueForValidation("UPC")
    If valueText = "" Then Exit Sub
    If IsWholeNumberText(valueText) = False Then
        AddValidationMessage validationMessages, BuildFieldValidationMessage("UPC", "UPC", "UPC must contain numbers only.")
        Exit Sub
    End If
    If Len(valueText) <> 8 And Len(valueText) <> 12 And Len(valueText) <> 13 And Len(valueText) <> 14 Then
        AddValidationMessage validationMessages, BuildFieldValidationMessage("UPC", "UPC", "UPC should be 8, 12, 13, or 14 digits.")
    End If
End Sub

Private Sub ValidateNumericField(ByRef validationMessages As Collection, ByVal fieldName As String, ByVal fieldLabel As String)
    Dim valueText As String
    valueText = GetControlValueForValidation(fieldName)
    If IsPositiveNumericText(valueText) = False Then
        AddValidationMessage validationMessages, BuildFieldValidationMessage(fieldName, fieldLabel, fieldLabel & " must be a positive number.")
    End If
End Sub

Private Sub ValidateOunces(ByRef validationMessages As Collection)
    Dim valueText As String
    valueText = GetControlValueForValidation("WEIGHT_OUNCES")
    If valueText = "" Then Exit Sub
    If IsPositiveNumericText(valueText) = False Then
        AddValidationMessage validationMessages, BuildFieldValidationMessage("WEIGHT_OUNCES", "Weight ounces", "Weight ounces must be a positive number.")
        Exit Sub
    End If
    If CDbl(valueText) >= 16 Then
        AddValidationMessage validationMessages, BuildFieldValidationMessage("WEIGHT_OUNCES", "Weight ounces", "Weight ounces must be less than 16.")
    End If
End Sub

Private Function BuildValidationMessage(ByVal validationMessages As Collection) As String
    Dim i As Long
    Dim messageText As String
    messageText = "Oops! Please correct the following before saving:" & vbCrLf & vbCrLf
    For i = 1 To validationMessages.Count
        messageText = messageText & i & ". " & validationMessages(i) & vbCrLf
    Next i
    BuildValidationMessage = messageText
End Function

Public Function ValidateItemEditorForm(ByVal frm As frmItemEditor) As Boolean
    Dim validationMessages As New Collection
    EnsureFieldDefinitionsLoaded
    ValidateUPC validationMessages
    ValidateNumericField validationMessages, "ITEM_HEIGHT", "Item width"
    ValidateNumericField validationMessages, "ITEM_WIDTH", "Item height"
    ValidateNumericField validationMessages, "ITEM_LENGTH", "Item length"
    ValidateNumericField validationMessages, "TOTAL_CONTENT_HEIGHT", "Total content width"
    ValidateNumericField validationMessages, "TOTAL_CONTENT_WIDTH", "Total content height"
    ValidateNumericField validationMessages, "TOTAL_CONTENT_LENGTH", "Total content length"
    ValidateNumericField validationMessages, "WEIGHT_POUNDS", "Weight pounds"
    ValidateOunces validationMessages
    If validationMessages.Count > 0 Then
        MsgBox BuildValidationMessage(validationMessages), vbExclamation, "Validation"
        ValidateItemEditorForm = False
        Exit Function
    End If
    ValidateItemEditorForm = True
End Function
