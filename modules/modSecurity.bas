Attribute VB_Name = "modSecurity"
Public Sub ProtectSystemSheets()

    On Error GoTo ErrorHandler

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets

        Select Case ws.Name

            Case "SETTINGS", "TABLES", "DATA"

                ws.Protect _
                    Password:="", _
                    UserInterfaceOnly:=True

        End Select

    Next ws

    Exit Sub

ErrorHandler:

    Call HandleError("ProtectSystemSheets", Err.Number, Err.Description)

End Sub

