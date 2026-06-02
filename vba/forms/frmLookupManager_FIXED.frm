VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLookupManager 
   Caption         =   "Manage Drop-down Lists"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   6000
   OleObjectBlob   =   "frmLookupManager.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLookupManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' =====================================================
' LOOKUP MANAGER USERFORM
' Dynamic interface for VIEW, EDIT, DELETE operations
' =====================================================

Private mLookupType As String
Private mLookupTypeDisplay As String
Private mMode As String  'VIEW, EDIT, DELETE
Private mValues As Collection
Private mControls As Collection

' =====================================================
' INITIALIZE
' =====================================================

Private Sub UserForm_Initialize()
    Set mValues = New Collection
    Set mControls = New Collection
    Me.Caption = "Manage Drop-down Lists"
    Me.Width = 450
    Me.Height = 550
End Sub

' =====================================================
' PUBLIC METHODS - ENTRY POINTS
' =====================================================

' Show lookup type selection
Public Function ShowLookupTypeSelection() As String
    Dim msg As String
    Dim resp As Variant
    Dim typeList As String
    Dim i As Long
    Dim lookupTypes As Variant
    Dim lookupDisplays As Variant
    
    lookupTypes = Array("BRAND", "COLOR", "MATERIAL", "STYLE", "PATTERN", "FUNCTIONAL_STATUS", "ODORS", "SHIPPING_RESTRICTIONS", "STORAGE_ENVIRONMENTS")
    lookupDisplays = Array("Brand", "Color", "Material", "Style", "Pattern", "Functional Status", "Odor", "Shipping Restriction", "Storage Environment")
    
    ' Build list with proper formatting
    typeList = "Select drop-down list to manage:" & vbCrLf & vbCrLf
    For i = LBound(lookupTypes) To UBound(lookupTypes)
        typeList = typeList & (i + 1) & ". " & lookupDisplays(i) & vbCrLf
    Next i
    typeList = typeList & vbCrLf & "Enter number (1-9):"
    
    Do
        resp = InputBox(typeList, "Select Drop-down List", "")
        
        ' Check for cancel
        If StrPtr(resp) = 0 Then  ' User pressed Cancel
            ShowLookupTypeSelection = ""
            Exit Function
        End If
        
        If Trim$(resp) = "" Then
            ShowLookupTypeSelection = ""
            Exit Function
        End If
        
        ' Validate input
        If IsNumeric(resp) Then
            Dim idx As Long
            idx = CLng(resp) - 1
            If idx >= LBound(lookupTypes) And idx <= UBound(lookupTypes) Then
                mLookupType = lookupTypes(idx)
                mLookupTypeDisplay = lookupDisplays(idx)
                ShowLookupTypeSelection = mLookupType
                Exit Function
            End If
        End If
        
        ' Invalid selection - show error and loop back
        MsgBox "Invalid selection. Please enter a number from 1 to 9.", vbExclamation, "Invalid Selection"
        ' Loop continues, showing the input box again
    Loop
End Function

' Show action selection based on lookup type
Public Function ShowActionSelection(ByVal lookupType As String, ByVal lookupDisplay As String) As String
    Dim msg As String
    Dim resp As String
    
    mLookupType = lookupType
    mLookupTypeDisplay = lookupDisplay
    
    ' Build dynamic message based on lookup type
    Dim article As String
    Dim plural As String
    
    ' Determine article (a/an)
    If lookupType = "ODORS" Then
        article = "an"
        plural = "Odors"
    Else
        article = "a"
        plural = lookupDisplay & "s"
        ' Special case for plurals that don't just add 's'
        If lookupType = "FUNCTIONAL_STATUS" Then plural = "Functional Status'"
        If lookupType = "STORAGE_ENVIRONMENTS" Then plural = "Storage Environments"
        If lookupType = "SHIPPING_RESTRICTIONS" Then plural = "Shipping Restrictions"
    End If
    
    msg = "What would you like to do with " & lookupDisplay & "s?" & vbCrLf & vbCrLf & _
          "1. EDIT " & article & " misspelled " & lookupDisplay & vbCrLf & _
          "2. DELETE " & article & " " & lookupDisplay & vbCrLf & _
          "3. VIEW current list of " & plural & vbCrLf & vbCrLf & _
          "Enter 1, 2, or 3:"
    
    Do
        resp = InputBox(msg, "Select Action", "")
        
        ' Check for cancel
        If StrPtr(resp) = 0 Then  ' User pressed Cancel
            ShowActionSelection = ""
            Exit Function
        End If
        
        Select Case Trim$(resp)
            Case "1", "EDIT", "Edit", "edit"
                ShowActionSelection = "EDIT"
                Exit Function
            Case "2", "DELETE", "Delete", "delete"
                ShowActionSelection = "DELETE"
                Exit Function
            Case "3", "VIEW", "View", "view"
                ShowActionSelection = "VIEW"
                Exit Function
            Case ""
                ShowActionSelection = ""
                Exit Function
        End Select
        
        ' Invalid selection - show error and loop back
        MsgBox "Invalid selection. Please enter 1, 2, or 3.", vbExclamation, "Invalid Selection"
        ' Loop continues
    Loop
End Function

' Show the main interface for VIEW, EDIT, or DELETE
Public Sub ShowValuesManager(ByVal lookupType As String, ByVal lookupDisplay As String, ByVal mode As String)
    mLookupType = lookupType
    mLookupTypeDisplay = lookupDisplay
    mMode = mode
    
    ' Load values
    Set mValues = GetLookupValues(lookupType)
    
    ' Clear any existing controls
    ClearDynamicControls
    
    ' Setup form based on mode
    Select Case mode
        Case "VIEW"
            SetupViewMode
        Case "EDIT"
            SetupEditMode
        Case "DELETE"
            SetupDeleteMode
    End Select
    
    ' Form will be shown by caller
End Sub

' =====================================================
' SETUP MODES
' =====================================================

Private Sub SetupViewMode()
    Dim i As Long
    Dim lbl As MSForms.Label
    Dim vItem As Variant
    Dim topPos As Single
    
    Me.Caption = "View " & mLookupTypeDisplay & "s"
    topPos = 60
    
    ' Title label
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblTitle")
    With lbl
        .Caption = "Current " & mLookupTypeDisplay & "s (" & mValues.Count & " items)"
        .Left = 12
        .Top = 12
        .Width = 400
        .Height = 20
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    ' Add read-only labels for each value
    i = 0
    For Each vItem In mValues
        Set lbl = Me.Controls.Add("Forms.Label.1", "lblValue_" & i)
        With lbl
            .Caption = (i + 1) & ". " & CStr(vItem)
            .Left = 24
            .Top = topPos + (i * 22)
            .Width = 400
            .Height = 18
            .Font.Size = 9
        End With
        i = i + 1
        If i >= 25 Then Exit For  ' Limit display for smaller form
    Next vItem
    
    ' Add buttons at bottom
    AddButton "btnEdit", "Edit " & mLookupTypeDisplay & "s", 12, 480, 140, 25
    AddButton "btnDelete", "Delete " & mLookupTypeDisplay & "s", 164, 480, 140, 25
    AddButton "btnCancel", "Cancel", 316, 480, 80, 25
End Sub

Private Sub SetupEditMode()
    Dim i As Long
    Dim txt As MSForms.TextBox
    Dim lbl As MSForms.Label
    Dim vItem As Variant
    Dim topPos As Single
    
    Me.Caption = "Edit " & mLookupTypeDisplay & "s"
    topPos = 60
    
    ' Title label
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblTitle")
    With lbl
        .Caption = "Edit " & mLookupTypeDisplay & " names (" & mValues.Count & " items)"
        .Left = 12
        .Top = 12
        .Width = 400
        .Height = 20
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    ' Add instruction label
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblInstruction")
    With lbl
        .Caption = "Correct any misspellings below:"
        .Left = 24
        .Top = 36
        .Width = 400
        .Height = 16
        .Font.Size = 9
        .ForeColor = &H808080
    End With
    
    ' Add text boxes for each value
    i = 0
    For Each vItem In mValues
        ' Label with number
        Set lbl = Me.Controls.Add("Forms.Label.1", "lblNum_" & i)
        With lbl
            .Caption = (i + 1) & "."
            .Left = 12
            .Top = topPos + (i * 26)
            .Width = 24
            .Height = 18
        End With
        
        ' Text box with value
        Set txt = Me.Controls.Add("Forms.TextBox.1", "txtValue_" & i)
        With txt
            .Text = CStr(vItem)
            .Left = 40
            .Top = topPos + (i * 26)
            .Width = 350
            .Height = 20
            .Font.Size = 9
        End With
        
        mControls.Add txt
        i = i + 1
        If i >= 25 Then Exit For  ' Limit display
    Next vItem
    
    ' Add buttons at bottom
    AddButton "btnSubmitEdit", "Submit Edit", 12, 480, 100, 25
    AddButton "btnCancel", "Cancel", 124, 480, 80, 25
End Sub

Private Sub SetupDeleteMode()
    Dim i As Long
    Dim chk As MSForms.CheckBox
    Dim lbl As MSForms.Label
    Dim vItem As Variant
    Dim topPos As Single
    
    Me.Caption = "Delete " & mLookupTypeDisplay & "s"
    topPos = 60
    
    ' Title label
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblTitle")
    With lbl
        .Caption = "Select " & mLookupTypeDisplay & "s to delete (" & mValues.Count & " items)"
        .Left = 12
        .Top = 12
        .Width = 400
        .Height = 20
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    ' Add instruction label
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblInstruction")
    With lbl
        .Caption = "Check the items you want to remove:"
        .Left = 24
        .Top = 36
        .Width = 400
        .Height = 16
        .Font.Size = 9
        .ForeColor = &H808080
    End With
    
    ' Add checkboxes for each value
    i = 0
    For Each vItem In mValues
        Set chk = Me.Controls.Add("Forms.CheckBox.1", "chkDelete_" & i)
        With chk
            .Caption = " " & (i + 1) & ". " & CStr(vItem)
            .Left = 24
            .Top = topPos + (i * 24)
            .Width = 400
            .Height = 20
            .Font.Size = 9
        End With
        
        mControls.Add chk
        i = i + 1
        If i >= 25 Then Exit For  ' Limit display
    Next vItem
    
    ' Add buttons at bottom
    AddButton "btnDelete", "Delete Selected", 12, 480, 120, 25
    AddButton "btnCancel", "Cancel", 144, 480, 80, 25
End Sub

' =====================================================
' BUTTON CLICK HANDLERS
' =====================================================

Private Sub btnEdit_Click()
    ' Switch from VIEW to EDIT mode
    Me.Hide
    ClearDynamicControls
    SetupEditMode
    Me.Show
End Sub

Private Sub btnDelete_Click()
    If mMode = "VIEW" Then
        ' Switch from VIEW to DELETE mode
        Me.Hide
        ClearDynamicControls
        SetupDeleteMode
        Me.Show
    Else
        ' Actually delete checked items
        ProcessDeletions
    End If
End Sub

Private Sub btnSubmitEdit_Click()
    ProcessEdits
End Sub

Private Sub btnCancel_Click()
    mMode = "CANCEL"
    Me.Hide
End Sub

' =====================================================
' PROCESSING
' =====================================================

Private Sub ProcessDeletions()
    Dim i As Long
    Dim chk As MSForms.CheckBox
    Dim itemsToDelete As Collection
    Dim deletedCount As Long
    Dim resp As VbMsgBoxResult
    Dim confirmMsg As String
    
    Set itemsToDelete = New Collection
    
    ' Collect checked items
    For i = 0 To mControls.Count - 1
        If TypeName(mControls(i + 1)) = "CheckBox" Then
            Set chk = mControls(i + 1)
            If chk.Value = True Then
                ' Extract value from caption (remove " 1. " prefix)
                Dim val As String
                val = Mid(chk.Caption, InStr(chk.Caption, ".") + 2)
                itemsToDelete.Add val
            End If
        End If
    Next i
    
    If itemsToDelete.Count = 0 Then
        MsgBox "No items selected for deletion.", vbInformation, "Nothing Selected"
        Exit Sub
    End If
    
    ' Build confirmation message
    confirmMsg = "Are you sure you want to delete these " & itemsToDelete.Count & " " & mLookupTypeDisplay & "s?" & vbCrLf & vbCrLf
    For i = 1 To itemsToDelete.Count
        confirmMsg = confirmMsg & "• " & itemsToDelete(i) & vbCrLf
        If i >= 10 And itemsToDelete.Count > 10 Then
            confirmMsg = confirmMsg & "... and " & (itemsToDelete.Count - 10) & " more" & vbCrLf
            Exit For
        End If
    Next i
    
    resp = MsgBox(confirmMsg, vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Deletion")
    
    If resp <> vbYes Then Exit Sub
    
    ' Perform deletions
    deletedCount = 0
    For i = 1 To itemsToDelete.Count
        If DeleteLookupValueInternal(mLookupType, itemsToDelete(i)) Then
            deletedCount = deletedCount + 1
        End If
    Next i
    
    ' Show confirmation
    MsgBox deletedCount & " " & mLookupTypeDisplay & "(s) deleted from the drop-down list.", vbInformation, "Deletion Complete"
    
    mMode = "DONE"
    Me.Hide
End Sub

Private Sub ProcessEdits()
    Dim i As Long
    Dim txt As MSForms.TextBox
    Dim editsMade As Long
    Dim oldValue As String
    Dim newValue As String
    
    editsMade = 0
    
    ' Process each text box
    For i = 1 To mControls.Count
        If TypeName(mControls(i)) = "TextBox" Then
            Set txt = mControls(i)
            newValue = Trim$(txt.Text)
            
            ' Get original value from the collection
            If i <= mValues.Count Then
                oldValue = CStr(mValues(i))
                
                If newValue <> oldValue And newValue <> "" Then
                    ' Perform the edit
                    If EditLookupValueInternal(mLookupType, oldValue, newValue) Then
                        editsMade = editsMade + 1
                    End If
                End If
            End If
        End If
    Next i
    
    If editsMade > 0 Then
        MsgBox editsMade & " " & mLookupTypeDisplay & "(s) updated.", vbInformation, "Edit Complete"
    Else
        MsgBox "No changes were made.", vbInformation, "No Changes"
    End If
    
    mMode = "DONE"
    Me.Hide
End Sub

' =====================================================
' HELPER FUNCTIONS
' =====================================================

Private Sub AddButton(ByVal btnName As String, ByVal caption As String, ByVal leftPos As Single, ByVal topPos As Single, ByVal width As Single, ByVal height As Single)
    Dim btn As MSForms.CommandButton
    Set btn = Me.Controls.Add("Forms.CommandButton.1", btnName)
    With btn
        .caption = caption
        .Left = leftPos
        .Top = topPos
        .Width = width
        .Height = height
        .Font.Size = 9
    End With
End Sub

Private Sub ClearDynamicControls()
    Dim i As Long
    For i = Me.Controls.Count - 1 To 0 Step -1
        ' Remove dynamically created controls (keep the form)
        Dim ctrlName As String
        ctrlName = Me.Controls(i).Name
        If ctrlName <> "frmLookupManager" And Not (ctrlName Like "btn*") Then
            On Error Resume Next
            Me.Controls.Remove ctrlName
            On Error GoTo 0
        End If
    Next i
    Set mControls = New Collection
End Sub

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

Public Property Get Result() As String
    Result = mMode
End Property
