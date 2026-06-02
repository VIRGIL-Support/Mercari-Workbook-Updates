Attribute VB_Name = "modResetLookups"
' Reset lookup tables to starter defaults only

Public Sub ResetAllLookupsToDefaults()
    Dim ws As Worksheet
    Dim resp As VbMsgBoxResult
    
    resp = MsgBox("This will DELETE all custom lookup values and reset to defaults only." & vbCrLf & vbCrLf & _
                  "Are you sure?", vbYesNo + vbExclamation, "Reset Lookups")
    
    If resp <> vbYes Then Exit Sub
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("LOOKUPS")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "LOOKUPS sheet not found!", vbExclamation
        Exit Sub
    End If
    
    ' Reset each lookup table by re-initializing
    Application.ScreenUpdating = False
    InitializeLookupArchitecture
    Application.ScreenUpdating = True
    
    MsgBox "Lookup tables reset to default values.", vbInformation
End Sub

' Clear test items from a specific table (manual cleanup)
Public Sub ClearTestItemsFromTable(ByVal tableName As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    Dim resp As VbMsgBoxResult
    Dim testKeyword As String
    
    testKeyword = InputBox("Enter keyword to identify test items (e.g., 'TEST', 'XXX', or leave blank to delete selected rows):", _
                           "Clear Test Items", "TEST")
    
    If testKeyword = "" Then Exit Sub
    
    Set ws = ThisWorkbook.Worksheets("LOOKUPS")
    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    On Error GoTo 0
    
    If tbl Is Nothing Then
        MsgBox "Table '" & tableName & "' not found!", vbExclamation
        Exit Sub
    End If
    
    ' Delete rows matching test keyword
    If Not tbl.DataBodyRange Is Nothing Then
        For Each row In tbl.ListRows
            If InStr(1, CStr(row.Range.Cells(1, 1).Value), testKeyword, vbTextCompare) > 0 Then
                row.Delete
            End If
        Next row
    End If
    
    MsgBox "Test items cleared from " & tableName, vbInformation
End Sub
