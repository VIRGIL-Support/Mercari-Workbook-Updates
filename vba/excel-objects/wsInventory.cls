VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo ErrorHandler

    If Target.CountLarge > 1 Then Exit Sub

    If Target.row < 2 Then Exit Sub

    Application.EnableEvents = False

    ' ITEM NAME CHANGED
    If Target.Column = COL_ITEM_NAME Then

        If Trim(Target.Value) <> "" Then

            InitializeInventoryRow Target.row

        Else

            ClearInventoryRow Target.row

        End If

    End If

    ' DATE SOLD CHANGED
    If Target.Column = COL_DATE_SOLD Then

        If Trim(Target.Value) <> "" Then
            ' Date entered - gray the row (sold pending)
            Rows(Target.row).Interior.Color = RGB(220, 220, 220)
        Else
            ' Date cleared - restore row to active status
            ' Clear any SOLD status
            Cells(Target.row, COL_STATUS).ClearContents
            ' Restore alternating row colors
            RefreshRowColorsAfterChange
            ' Ensure row is visible
            Rows(Target.row).Hidden = False
        End If

    End If

CleanExit:

    Application.EnableEvents = True

    Exit Sub

ErrorHandler:

    Application.EnableEvents = True

    HandleError "InventoryWorksheetChange", Err.Number, Err.Description

End Sub
