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

    If Target.Row < 2 Then Exit Sub

    Application.EnableEvents = False

    ' ITEM NAME CHANGED
    If Target.Column = COL_ITEM_NAME Then

        If Trim(Target.Value) <> "" Then

            InitializeInventoryRow Target.Row

        Else

            ClearInventoryRow Target.Row

        End If

    End If

    ' DATE SOLD CHANGED
    If Target.Column = COL_DATE_SOLD Then

        If Trim(Target.Value) <> "" Then
            ' Date entered - gray the row and mark as sold
            Cells(Target.Row, COL_STATUS).Value = STATUS_SOLD
            Rows(Target.Row).Interior.Color = RGB(220, 220, 220)
        Else
            ' Date cleared - restore white background and reset status
            Cells(Target.Row, COL_STATUS).Value = STATUS_NEW
            Rows(Target.Row).Interior.Color = RGB(255, 255, 255)
        End If

    End If

CleanExit:

    Application.EnableEvents = True

    Exit Sub

ErrorHandler:

    Application.EnableEvents = True

    HandleError "InventoryWorksheetChange", Err.Number, Err.Description

End Sub
