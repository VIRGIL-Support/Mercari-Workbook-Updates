VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    
    Call CheckAtkinsonHyperlegibleFont
    Call InitializeWorkbook
    Call CheckForUpdatesOnOpen

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

    ' Skip backup if an update transfer is in progress
    ' (prevents errors when the update script closes the old workbook)
    If Dir(Environ$("TEMP") & "\MercariUpdateSource.txt") <> "" Then Exit Sub
    
    Call BackupWorkbookBeforeClose

End Sub
