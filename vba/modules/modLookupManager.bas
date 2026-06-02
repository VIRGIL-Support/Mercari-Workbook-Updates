Attribute VB_Name = "modLookupManager"
' =====================================================
' LOOKUP TABLE MANAGEMENT - SIMPLE DIALOG VERSION
' Uses InputBox and MsgBox only (no custom forms)
' =====================================================

' =====================================================
' MAIN ENTRY POINT
' =====================================================

Public Sub ManageLookupValues()
    Dim lookupType As String
    Dim lookupDisplay As String
    Dim action As String
    
    ' Step 1: Select lookup type
    lookupType = SelectLookupType()
    If lookupType = "" Then Exit Sub
    
    ' Get display name
    lookupDisplay = GetLookupDisplayName(lookupType)
    
    ' Step 2: Select action
    action = SelectAction(lookupType, lookupDisplay)
    If action = "" Then Exit Sub
    
    ' Step 3: Execute action
    Select Case action
        Case "VIEW"
            ViewLookupValues lookupType, lookupDisplay
        Case "EDIT"
            EditLookupValues lookupType, lookupDisplay
        Case "DELETE"
            DeleteLookupValues lookupType, lookupDisplay
    End Select
End Sub

' =====================================================
' STEP 1: SELECT LOOKUP TYPE
' =====================================================

Private Function SelectLookupType() As String
    Dim lookupTypes As Variant
    Dim lookupDisplays As Variant
    Dim typeList As String
    Dim resp As String
    Dim i As Long
    
    lookupTypes = Array("BRAND", "COLOR", "MATERIAL", "STYLE", "PATTERN", "FUNCTIONAL_STATUS", "ODORS", "SHIPPING_RESTRICTIONS", "STORAGE_ENVIRONMENTS")
    lookupDisplays = Array("Brand", "Color", "Material", "Style", "Pattern", "Functional Status", "Odor", "Shipping Restriction", "Storage Environment")
    
    ' Build numbered list
    typeList = "Select drop-down list to manage:" & vbCrLf & vbCrLf
    For i = LBound(lookupTypes) To UBound(lookupTypes)
        typeList = typeList & (i + 1) & ". " & lookupDisplays(i) & vbCrLf
    Next i
    typeList = typeList & vbCrLf & "Enter number 1-9:"
    
    Do
        resp = InputBox(typeList, "Select Drop-down List", "")
        
        ' Check cancel
        If StrPtr(resp) = 0 Then
            SelectLookupType = ""
            Exit Function
        End If
        
        If Trim$(resp) = "" Then
            SelectLookupType = ""
            Exit Function
        End If
        
        ' Validate
        If IsNumeric(resp) Then
            Dim idx As Long
            idx = CLng(resp) - 1
            If idx >= 0 And idx <= 8 Then
                SelectLookupType = lookupTypes(idx)
                Exit Function
            End If
        End If
        
        MsgBox "Please enter a number from 1 to 9.", vbExclamation, "Invalid Selection"
    Loop
End Function

' =====================================================
' STEP 2: SELECT ACTION
' =====================================================

Private Function SelectAction(ByVal lookupType As String, ByVal lookupDisplay As String) As String
    Dim msg As String
    Dim resp As String
    Dim article As String
    Dim plural As String
    
    ' Determine article (a/an) and plural form
    Select Case lookupType
        Case "ODORS"
            article = "an"
            plural = "Odors"
        Case "FUNCTIONAL_STATUS"
            article = "a"
            plural = "Functional Status'"
        Case Else
            article = "a"
            plural = lookupDisplay & "s"
    End Select
    
    msg = "What would you like to do with " & lookupDisplay & "s?" & vbCrLf & vbCrLf & _
          "1. EDIT " & article & " misspelled " & lookupDisplay & vbCrLf & _
          "2. DELETE " & article & " " & lookupDisplay & vbCrLf & _
          "3. VIEW current list of " & plural & vbCrLf & vbCrLf & _
          "Enter 1, 2, or 3:"
    
    Do
        resp = InputBox(msg, "Select Action - " & lookupDisplay & "s", "")
        
        If StrPtr(resp) = 0 Then
            SelectAction = ""
            Exit Function
        End If
        
        Select Case Trim$(resp)
            Case "1", "EDIT", "Edit", "edit"
                SelectAction = "EDIT"
                Exit Function
            Case "2", "DELETE", "Delete", "delete"
                SelectAction = "DELETE"
                Exit Function
            Case "3", "VIEW", "View", "view"
                SelectAction = "VIEW"
                Exit Function
            Case ""
                SelectAction = ""
                Exit Function
        End Select
        
        MsgBox "Please enter 1, 2, or 3.", vbExclamation, "Invalid Selection"
    Loop
End Function

' =====================================================
' STEP 3A: VIEW LOOKUP VALUES
' =====================================================

Private Sub ViewLookupValues(ByVal lookupType As String, ByVal lookupDisplay As String)
    Dim colValues As Collection
    Dim vItem As Variant
    Dim msg As String
    Dim i As Long
    
    Set colValues = GetLookupValues(lookupType)
    
    msg = UCase$(lookupDisplay) & "S - " & colValues.Count & " ITEMS" & vbCrLf
    msg = msg & String(50, "=") & vbCrLf & vbCrLf
    
    i = 1
    For Each vItem In colValues
        msg = msg & i & ". " & CStr(vItem) & vbCrLf
        i = i + 1
        ' Limit message box length
        If i > 50 Then
            msg = msg & "... (" & (colValues.Count - 50) & " more items)" & vbCrLf
            Exit For
        End If
    Next vItem
    
    MsgBox msg, vbInformation, "View " & lookupDisplay & "s"
End Sub

' =====================================================
' STEP 3B: EDIT LOOKUP VALUES
' =====================================================

Private Sub EditLookupValues(ByVal lookupType As String, ByVal lookupDisplay As String)
    Dim colValues As Collection
    Dim vItem As Variant
    Dim currentList As String
    Dim valueToEdit As String
    Dim newValue As String
    Dim i As Long
    Dim found As Boolean
    
    Set colValues = GetLookupValues(lookupType)
    
    If colValues.Count = 0 Then
        MsgBox "No " & lookupDisplay & "s to edit.", vbInformation
        Exit Sub
    End If
    
    ' Build numbered list
    currentList = "CURRENT " & UCase$(lookupDisplay) & "S:" & vbCrLf & vbCrLf
    i = 1
    For Each vItem In colValues
        currentList = currentList & i & ". " & CStr(vItem) & vbCrLf
        i = i + 1
        If i > 50 Then Exit For
    Next vItem
    
    ' Get value to edit
    valueToEdit = InputBox(currentList & vbCrLf & vbCrLf & "Type the EXACT " & lookupDisplay & " to edit:", "Edit " & lookupDisplay, "")
    
    If StrPtr(valueToEdit) = 0 Or Trim$(valueToEdit) = "" Then
        MsgBox "Edit cancelled.", vbInformation
        Exit Sub
    End If
    
    ' Verify it exists
    found = LookupValueExists(lookupType, valueToEdit)
    If Not found Then
        MsgBox "'" & valueToEdit & "' not found in " & lookupDisplay & " list." & vbCrLf & vbCrLf & "Please check spelling and try again.", vbExclamation
        Exit Sub
    End If
    
    ' Get new value
    newValue = InputBox("Edit " & lookupDisplay & ":" & vbCrLf & vbCrLf & _
                        "Current: " & valueToEdit & vbCrLf & vbCrLf & _
                        "Enter new spelling:", "Edit " & lookupDisplay, valueToEdit)
    
    If StrPtr(newValue) = 0 Or Trim$(newValue) = "" Then
        MsgBox "Edit cancelled.", vbInformation
        Exit Sub
    End If
    
    If newValue = valueToEdit Then
        MsgBox "No changes made.", vbInformation
        Exit Sub
    End If
    
    ' Confirm
    Dim resp As VbMsgBoxResult
    resp = MsgBox("Confirm edit:" & vbCrLf & vbCrLf & _
                  "FROM: " & valueToEdit & vbCrLf & _
                  "TO: " & newValue, vbYesNo + vbQuestion, "Confirm Edit")
    
    If resp <> vbYes Then
        MsgBox "Edit cancelled.", vbInformation
        Exit Sub
    End If
    
    ' Perform edit
    If EditLookupValueInternal(lookupType, valueToEdit, newValue) Then
        MsgBox "Successfully updated!" & vbCrLf & vbCrLf & _
               "Changed: " & valueToEdit & " → " & newValue, vbInformation, "Edit Complete"
    Else
        MsgBox "Could not update. Please try again.", vbExclamation, "Error"
    End If
End Sub

' =====================================================
' STEP 3C: DELETE LOOKUP VALUES
' =====================================================

Private Sub DeleteLookupValues(ByVal lookupType As String, ByVal lookupDisplay As String)
    Dim colValues As Collection
    Dim vItem As Variant
    Dim currentList As String
    Dim valuesToDelete As String
    Dim itemsToDelete As Collection
    Dim i As Long
    Dim resp As VbMsgBoxResult
    Dim deletedCount As Long
    Dim confirmMsg As String
    
    Set colValues = GetLookupValues(lookupType)
    
    If colValues.Count = 0 Then
        MsgBox "No " & lookupDisplay & "s to delete.", vbInformation
        Exit Sub
    End If
    
    ' Build numbered list
    currentList = "CURRENT " & UCase$(lookupDisplay) & "S:" & vbCrLf & vbCrLf
    i = 1
    For Each vItem In colValues
        currentList = currentList & i & ". " & CStr(vItem) & vbCrLf
        i = i + 1
        If i > 50 Then Exit For
    Next vItem
    
    ' Get numbers to delete
    valuesToDelete = InputBox(currentList & vbCrLf & vbCrLf & _
                               "Enter numbers to DELETE (comma-separated):" & vbCrLf & _
                               "Example: 1,3,5", "Delete " & lookupDisplay & "s", "")
    
    If StrPtr(valuesToDelete) = 0 Or Trim$(valuesToDelete) = "" Then
        MsgBox "Delete cancelled.", vbInformation
        Exit Sub
    End If
    
    ' Parse numbers
    Set itemsToDelete = New Collection
    Dim nums As Variant
    Dim num As Variant
    Dim idx As Long
    
    nums = Split(valuesToDelete, ",")
    For Each num In nums
        If IsNumeric(Trim$(num)) Then
            idx = CLng(Trim$(num))
            If idx >= 1 And idx <= colValues.Count Then
                itemsToDelete.Add colValues(idx)
            End If
        End If
    Next num
    
    If itemsToDelete.Count = 0 Then
        MsgBox "No valid items selected.", vbExclamation
        Exit Sub
    End If
    
    ' Build confirmation
    confirmMsg = "Delete these " & itemsToDelete.Count & " " & lookupDisplay & "s?" & vbCrLf & vbCrLf
    For i = 1 To itemsToDelete.Count
        confirmMsg = confirmMsg & "• " & itemsToDelete(i) & vbCrLf
        If i >= 15 And itemsToDelete.Count > 15 Then
            confirmMsg = confirmMsg & "... (" & (itemsToDelete.Count - 15) & " more)" & vbCrLf
            Exit For
        End If
    Next i
    
    resp = MsgBox(confirmMsg, vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete")
    
    If resp <> vbYes Then
        MsgBox "Delete cancelled.", vbInformation
        Exit Sub
    End If
    
    ' Perform deletions
    deletedCount = 0
    For i = 1 To itemsToDelete.Count
        If DeleteLookupValueInternal(lookupType, itemsToDelete(i)) Then
            deletedCount = deletedCount + 1
        End If
    Next i
    
    MsgBox deletedCount & " " & lookupDisplay & "(s) deleted from drop-down list.", vbInformation, "Delete Complete"
End Sub

' =====================================================
' HELPER FUNCTIONS
' =====================================================

Public Function GetLookupDisplayName(ByVal lookupType As String) As String
    Select Case UCase$(Trim$(lookupType))
        Case "BRAND", "BRANDS"
            GetLookupDisplayName = "Brand"
        Case "COLOR", "COLORS"
            GetLookupDisplayName = "Color"
        Case "MATERIAL", "MATERIALS"
            GetLookupDisplayName = "Material"
        Case "STYLE", "STYLES"
            GetLookupDisplayName = "Style"
        Case "PATTERN", "PATTERNS"
            GetLookupDisplayName = "Pattern"
        Case "FUNCTIONAL_STATUS"
            GetLookupDisplayName = "Functional Status"
        Case "ODOR", "ODORS"
            GetLookupDisplayName = "Odor"
        Case "SHIPPING_RESTRICTION", "SHIPPING_RESTRICTIONS"
            GetLookupDisplayName = "Shipping Restriction"
        Case "STORAGE_ENVIRONMENT", "STORAGE_ENVIRONMENTS"
            GetLookupDisplayName = "Storage Environment"
        Case Else
            GetLookupDisplayName = lookupType
    End Select
End Function

Private Function LookupValueExists(ByVal lookupType As String, ByVal lookupValue As String) As Boolean
    Dim colValues As Collection
    Dim vItem As Variant
    
    Set colValues = GetLookupValues(lookupType)
    For Each vItem In colValues
        If CStr(vItem) = lookupValue Then
            LookupValueExists = True
            Exit Function
        End If
    Next vItem
End Function

Private Function GetLookupTableByTypeInternal(ByVal lookupType As String) As ListObject
    Dim ws As Worksheet
    Dim tableName As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("LOOKUPS")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Function
    
    tableName = ResolveTableNameInternal(lookupType)
    If tableName = "" Then Exit Function
    
    On Error Resume Next
    Set GetLookupTableByTypeInternal = ws.ListObjects(tableName)
    On Error GoTo 0
End Function

Private Function ResolveTableNameInternal(ByVal lookupType As String) As String
    Select Case UCase$(Trim$(lookupType))
        Case "BRAND", "BRANDS"
            ResolveTableNameInternal = "tblBrands"
        Case "COLOR", "COLORS"
            ResolveTableNameInternal = "tblColors"
        Case "MATERIAL", "MATERIALS"
            ResolveTableNameInternal = "tblMaterials"
        Case "STYLE", "STYLES"
            ResolveTableNameInternal = "tblStyles"
        Case "PATTERN", "PATTERNS"
            ResolveTableNameInternal = "tblPatterns"
        Case "FUNCTIONAL_STATUS"
            ResolveTableNameInternal = "tblFunctionalStatus"
        Case "ODOR", "ODORS"
            ResolveTableNameInternal = "tblOdors"
        Case "SHIPPING_RESTRICTION", "SHIPPING_RESTRICTIONS"
            ResolveTableNameInternal = "tblShippingRestrictions"
        Case "STORAGE_ENVIRONMENT", "STORAGE_ENVIRONMENTS"
            ResolveTableNameInternal = "tblStorageEnvironments"
    End Select
End Function

Private Function DeleteLookupValueInternal(ByVal lookupType As String, ByVal valueToDelete As String) As Boolean
    Dim tbl As ListObject
    Dim row As ListRow
    
    Set tbl = GetLookupTableByTypeInternal(lookupType)
    If tbl Is Nothing Then Exit Function
    
    If Not tbl.DataBodyRange Is Nothing Then
        For Each row In tbl.ListRows
            If CStr(row.Range.Cells(1, 1).Value) = valueToDelete Then
                row.Delete
                DeleteLookupValueInternal = True
                Exit Function
            End If
        Next row
    End If
End Function

Private Function EditLookupValueInternal(ByVal lookupType As String, ByVal oldValue As String, ByVal newValue As String) As Boolean
    Dim tbl As ListObject
    Dim row As Range
    
    Set tbl = GetLookupTableByTypeInternal(lookupType)
    If tbl Is Nothing Then Exit Function
    
    If Not tbl.DataBodyRange Is Nothing Then
        For Each row In tbl.DataBodyRange.Rows
            If CStr(row.Cells(1, 1).Value) = oldValue Then
                row.Cells(1, 1).Value = newValue
                EditLookupValueInternal = True
                Exit Function
            End If
        Next row
    End If
End Function
