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
    
    ' Securely hide system worksheets (users cannot unhide via Format menu)
    On Error Resume Next
    ThisWorkbook.Worksheets("DATA").Visible = xlSheetVeryHidden
    ThisWorkbook.Worksheets("LOOKUPS").Visible = xlSheetVeryHidden
    ThisWorkbook.Worksheets("SETTINGS").Visible = xlSheetVeryHidden
    ThisWorkbook.Worksheets("TABLES").Visible = xlSheetVeryHidden
    ThisWorkbook.Worksheets("COPYRIGHT_INFO").Visible = xlSheetVeryHidden
    ThisWorkbook.Worksheets("SOLD_DATA").Visible = xlSheetVeryHidden
    On Error GoTo 0
    
    Call CheckAtkinsonHyperlegibleFont
    Call InitializeWorkbook
    Call CheckForUpdatesOnOpen
    
    ' Create developer test button on HELP sheet (A100)
    Call CreateUnhideButton
    
    ' Navigate to appropriate worksheet based on data
    NavigateToStartupWorksheet

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)

    ' Skip if an update transfer is in progress
    If Dir(Environ$("TEMP") & "\MercariUpdateSource.txt") <> "" Then Exit Sub
    
    ' Process any sold items (including CANCELLED) before closing
    Call ProcessSoldItemsOnClose
    
    Call BackupWorkbookBeforeClose

End Sub

