Attribute VB_Name = "modBackup"
Option Explicit

Private Function GetBackupFolderPath() As String
    GetBackupFolderPath = GetSettingValue("BACKUPS_FOLDER")
    If Trim$(GetBackupFolderPath) = "" Then GetBackupFolderPath = ThisWorkbook.Path & "\Backups"
    CreateFolderIfMissing GetBackupFolderPath
End Function

Public Sub CreateWorkbookBackup()

    On Error GoTo ErrorHandler

    Dim backupFolder As String
    Dim backupFileName As String
    Dim timestamp As String

    backupFolder = GetBackupFolderPath()

    timestamp = Format(Now, "yyyy-mm-dd_hh-nn-ss")

    backupFileName = backupFolder & "\" & _
                     "MercariSystem_Backup_" & timestamp & ".xlsm"

    ThisWorkbook.SaveCopyAs backupFileName

    Exit Sub

ErrorHandler:

    Call HandleError("CreateWorkbookBackup", Err.Number, Err.Description)

End Sub

Public Sub BackupWorkbookBeforeClose()

    On Error GoTo ErrorHandler

    Dim soldItemCount As Long

    soldItemCount = ProcessSoldItemsOnClose()
    ThisWorkbook.Save
    CreateWorkbookBackup

    If soldItemCount > 0 Then
        MsgBox _
            soldItemCount & " sold item(s) were removed from the Inventory worksheet and moved to the Sold Items worksheet for reference." & vbCrLf & vbCrLf & _
            "A backup copy of the workbook was also saved.", _
            vbInformation, _
            "Sold Items Processed"
    End If

    Exit Sub

ErrorHandler:

    Call HandleError("BackupWorkbookBeforeClose", Err.Number, Err.Description)

End Sub

