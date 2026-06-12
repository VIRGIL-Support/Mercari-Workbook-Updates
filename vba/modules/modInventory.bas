Attribute VB_Name = "modInventory"
Option Explicit

' =====================================================
' PROTECT INVENTORY WORKSHEET (LOCK COLUMNS BUT ALLOW BUTTONS)
' =====================================================

Public Sub ProtectInventoryWorksheet()
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_INVENTORY)
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    ' Unprotect first (in case already protected)
    ws.Unprotect Password:=""
    
    ' Protect with specific settings:
    ' - UserInterfaceOnly=True: macros can still modify
    ' - AllowFormattingColumns=False: prevent column resizing
    ' - AllowFiltering=True: allow filtering if needed
    ' - AllowSorting=True: allow sorting
    ' - AllowUsingPivotTables=True
    ws.Protect Password:="", _
                 UserInterfaceOnly:=True, _
                 AllowFormattingColumns:=False, _
                 AllowFormattingRows:=False, _
                 AllowInsertingColumns:=False, _
                 AllowInsertingRows:=False, _
                 AllowDeletingColumns:=False, _
                 AllowDeletingRows:=False, _
                 AllowSorting:=True, _
                 AllowFiltering:=True, _
                 AllowUsingPivotTables:=True
    
    ' Note: Shapes (buttons) remain clickable because they have OnAction macros
    ' and UserInterfaceOnly allows macro execution
    
End Sub

Public Sub UnprotectInventoryWorksheet()
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_INVENTORY)
    On Error GoTo 0
    
    If Not ws Is Nothing Then ws.Unprotect Password:=""
    
End Sub

' =====================================================
' FONT AVAILABILITY CHECK
' =====================================================

Public Sub CheckAtkinsonHyperlegibleFont()

    On Error GoTo ErrorHandler

    Dim fontInstalled As Boolean
    Dim response As VbMsgBoxResult
    Dim ws As Worksheet

    fontInstalled = IsFontInstalled("Atkinson Hyperlegible")

    If Not fontInstalled Then
        response = MsgBox( _
            "This workbook is best viewed with ATKINSON HYPERLEGIBLE font." & vbCrLf & vbCrLf & _
            "Atkinson Hyperlegible is the best font for making it easiest to differentiate between similar characters." & vbCrLf & vbCrLf & _
            "Click 'OK' to open the download page, or 'Cancel' to continue with a backup font." & vbCrLf & vbCrLf & _
            "Font Download: https://fonts.google.com/specimen/Atkinson+Hyperlegible", _
            vbOKCancel + vbInformation, _
            "Atkinson Hyperlegible Font Not Found")

        If response = vbOK Then
            ThisWorkbook.FollowHyperlink "https://fonts.google.com/specimen/Atkinson+Hyperlegible"
        Else
            SetBackupFont
        End If
    End If

    Exit Sub

ErrorHandler:
    HandleError "CheckAtkinsonHyperlegibleFont", Err.Number, Err.Description

End Sub

Private Function IsFontInstalled(ByVal fontName As String) As Boolean

    On Error Resume Next
    Dim testRange As Range
    Set testRange = ThisWorkbook.Sheets(1).Range("A1")
    testRange.Font.Name = fontName
    IsFontInstalled = (testRange.Font.Name = fontName)
    On Error GoTo 0

End Function

Private Sub SetBackupFont()

    On Error Resume Next
    ThisWorkbook.Sheets(1).Cells.Font.Name = "Calibri"
    On Error GoTo 0

End Sub

' =====================================================
' HIDDEN WATERMARK FOR COPYRIGHT TRACKING
' =====================================================

Public Sub SetHiddenWatermark()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim watermarkCell As Range
    Dim watermarkText As String
    Dim purchaserLastName As String
    Dim purchaseDate As String
    Dim orderNumber As String
    Dim watermarkData As String
    
    ' Check if watermark already set
    watermarkData = GetSettingValue("WATERMARK_DATA")
    
    If watermarkData = "" Then
        ' Prompt for 3 required fields (no defaults - user must enter info)
        Do
            purchaserLastName = Trim$(InputBox("Step 1 of 3: Enter the LAST NAME of the Purchaser (from your receipt):", "Workbook Activation - Step 1"))
            If purchaserLastName = "" Then
                MsgBox "Last name is required to activate this workbook.", vbExclamation, "Required Field"
            End If
        Loop While purchaserLastName = ""
        
        Do
            purchaseDate = Trim$(InputBox("Step 2 of 3: Enter the DATE of Purchase (MM/DD/YYYY):", "Workbook Activation - Step 2"))
            If purchaseDate = "" Then
                MsgBox "Date of purchase is required to activate this workbook.", vbExclamation, "Required Field"
            End If
        Loop While purchaseDate = ""
        
        Do
            orderNumber = Trim$(InputBox("Step 3 of 3: Enter the Order/Invoice Number from your receipt:" & vbCrLf & vbCrLf & _
                                   "Etsy: Order #" & vbCrLf & _
                                   "Gumroad: License Key" & vbCrLf & _
                                   "PayPal: Transaction ID" & vbCrLf & _
                                   "Other: Invoice Number", _
                                   "Workbook Activation - Step 3"))
            If orderNumber = "" Then
                MsgBox "Order/Invoice number is required to activate this workbook.", vbExclamation, "Required Field"
            End If
        Loop While orderNumber = ""
        
        ' Save combined watermark data
        watermarkData = purchaserLastName & "|" & purchaseDate & "|" & orderNumber
        UpdateSetting "WATERMARK_DATA", watermarkData
        UpdateSetting "PURCHASER_LASTNAME", purchaserLastName
        UpdateSetting "PURCHASE_DATE", purchaseDate
        UpdateSetting "ORDER_NUMBER", orderNumber
    Else
        ' Parse existing watermark data
        Dim parts() As String
        parts = Split(watermarkData, "|")
        If UBound(parts) >= 2 Then
            purchaserLastName = parts(0)
            purchaseDate = parts(1)
            orderNumber = parts(2)
        Else
            purchaserLastName = "UNKNOWN"
            purchaseDate = Format(Now, "MM/DD/YYYY")
            orderNumber = "NO-ORDER"
        End If
    End If
    
    Set ws = ThisWorkbook.Sheets(WS_INVENTORY)
    Set watermarkCell = ws.Range("Z1")
    
    ' Create watermark text with all 3 fields
    watermarkText = "AUTH-VIR-" & Format(Now, "yyyymmdd") & "-" & purchaserLastName & "-" & orderNumber
    
    ' Set the watermark
    watermarkCell.Value = watermarkText
    watermarkCell.Font.Color = RGB(255, 255, 255) ' White text
    watermarkCell.Interior.Color = RGB(255, 255, 255) ' White background
    watermarkCell.Font.Size = 8
    
    On Error GoTo 0
End Sub

' =====================================================
' INITIALIZE INVENTORY ROW
' =====================================================

Public Sub InitializeInventoryRow(ByVal rowNum As Long)

    On Error GoTo ErrorHandler

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets(WS_INVENTORY)

    If Trim(ws.Cells(rowNum, COL_ITEM_NUMBER).Value) = "" Then
        ws.Cells(rowNum, COL_ITEM_NUMBER).Value = GetNextItemNumber()
    End If

    CreateStartButton rowNum

    PrepareNextInventoryRow rowNum + 1

    Exit Sub

ErrorHandler:
    HandleError "InitializeInventoryRow", Err.Number, Err.Description

End Sub

' =====================================================
' PREP NEXT ROW
' =====================================================

Public Sub PrepareNextInventoryRow(ByVal rowNum As Long)

    On Error Resume Next

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets(WS_INVENTORY)

    If Trim(ws.Cells(rowNum, COL_ITEM_NUMBER).Value) = "" Then
        ws.Cells(rowNum, COL_ITEM_NUMBER).Value = GetNextItemNumber()
    End If
    
    ' Refresh alternating row colors after adding new row
    RefreshRowColorsAfterChange

End Sub

Private Sub EnsureOnePendingInventoryRow()

    Dim ws As Worksheet
    Dim rowNum As Long

    Set ws = ThisWorkbook.Sheets(WS_INVENTORY)

    For rowNum = 2 To ws.Cells(ws.Rows.Count, COL_ITEM_NUMBER).End(xlUp).row
        If Not IsRetiredInventoryRow(ws, rowNum) And _
           Trim$(CStr(ws.Cells(rowNum, COL_ITEM_NUMBER).Value)) <> "" And _
           Trim$(CStr(ws.Cells(rowNum, COL_ITEM_NAME).Value)) = "" And _
           Trim$(CStr(ws.Cells(rowNum, COL_ITEM_PRICE).Value)) = "" And _
           Trim$(CStr(ws.Cells(rowNum, COL_DATE_SOLD).Value)) = "" Then
            Exit Sub
        End If
    Next rowNum

    PrepareNextInventoryRow GetNextOpenInventoryRow()

End Sub

' =====================================================
' GET NEXT ITEM NUMBER
' =====================================================

Public Function GetNextItemNumber() As String

    Dim ws As Worksheet
    Dim wsSold As Worksheet
    Dim lastRow As Long
    Dim rowNum As Long
    Dim itemValue As String
    Dim prefixPart As String
    Dim numberPart As Long
    Dim maxNumberPart As Long

    Set ws = ThisWorkbook.Sheets(WS_INVENTORY)

    lastRow = ws.Cells(ws.Rows.Count, COL_ITEM_NUMBER).End(xlUp).row
    For rowNum = 2 To lastRow
        itemValue = Trim$(CStr(ws.Cells(rowNum, COL_ITEM_NUMBER).Value))
        If InStr(itemValue, "-") > 0 And _
           (Trim$(CStr(ws.Cells(rowNum, COL_ITEM_NAME).Value)) <> "" Or _
            Trim$(CStr(ws.Cells(rowNum, COL_ITEM_PRICE).Value)) <> "" Or _
            Trim$(CStr(ws.Cells(rowNum, COL_DATE_SOLD).Value)) <> "") Then
            numberPart = CLng(Split(itemValue, "-")(1))
            If numberPart > maxNumberPart Then maxNumberPart = numberPart
        End If
    Next rowNum

    On Error Resume Next
    Set wsSold = ThisWorkbook.Worksheets(WS_SOLD_ITEMS)
    On Error GoTo 0
    If Not wsSold Is Nothing Then
        lastRow = wsSold.Cells(wsSold.Rows.Count, 1).End(xlUp).row
        For rowNum = 2 To lastRow
            itemValue = Trim$(CStr(wsSold.Cells(rowNum, 1).Value))
            If InStr(itemValue, "-") > 0 Then
                numberPart = CLng(Split(itemValue, "-")(1))
                If numberPart > maxNumberPart Then maxNumberPart = numberPart
            End If
        Next rowNum
    End If

    prefixPart = "100"
    GetNextItemNumber = prefixPart & "-" & Format(maxNumberPart + 1, "000")

End Function

' =====================================================
' CREATE START BUTTON
' =====================================================

Public Sub CreateStartButton(ByVal rowNum As Long)

    On Error Resume Next

    Dim ws As Worksheet
    Dim shp As Shape

    Set ws = ThisWorkbook.Sheets(WS_INVENTORY)

    DeleteRowButtons rowNum

    Set shp = CreateButton(ws, rowNum, COL_START_EDIT, "BTN_START_" & rowNum, "START", "StartButtonClick")

End Sub

' =====================================================
' CREATE EDIT + AI BUTTONS
' =====================================================

Public Sub CreateEditAndAIButtons(ByVal rowNum As Long)

    On Error Resume Next

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets(WS_INVENTORY)

    DeleteRowButtons rowNum

    CreateButton ws, rowNum, COL_START_EDIT, "BTN_EDIT_" & rowNum, "EDIT", "EditButtonClick"

    CreateButton ws, rowNum, COL_COPY_FOR_AI, "BTN_AI_" & rowNum, "COPY FOR AI", "CopyForAIButtonClick"

End Sub

' =====================================================
' CREATE DETAILS BUTTON
' =====================================================

Public Sub CreateDetailsButton(ByVal rowNum As Long)

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets(WS_INVENTORY)

    CreateButton ws, rowNum, COL_DETAILS, "BTN_DETAILS_" & rowNum, "PASTE DETAILS", "PasteDetailsButtonClick"

End Sub

' =====================================================
' CREATE VIEW DETAILS BUTTON
' =====================================================

Public Sub CreateViewDetailsButton(ByVal rowNum As Long)

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets(WS_INVENTORY)

    ws.Cells(rowNum, COL_VIEW_DETAILS).ClearContents
    CreateButton ws, rowNum, COL_VIEW_DETAILS, "BTN_VIEW_DETAILS_" & rowNum, "VIEW ITEM DETAILS", "ViewDetailsButtonClick"

End Sub

' =====================================================
' CREATE DETAILS FOLDER BUTTON
' =====================================================

Public Sub CreateDetailsFolderButton(ByVal rowNum As Long)

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets(WS_INVENTORY)

    ws.Cells(rowNum, COL_DETAILS_FOLDER).ClearContents
    CreateButton ws, rowNum, COL_DETAILS_FOLDER, "BTN_FOLDER_" & rowNum, "VIEW ITEM FOLDER", "ViewFolderButtonClick"

End Sub

Public Sub CreateViewAllFoldersButton(ByVal rowNum As Long)

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets(WS_INVENTORY)

    ws.Cells(rowNum, COL_VIEW_ALL_FOLDERS).ClearContents
    CreateButton ws, rowNum, COL_VIEW_ALL_FOLDERS, "BTN_ALL_FOLDERS_" & rowNum, "VIEW ALL FOLDERS", "ViewAllFoldersButtonClick"

End Sub

Public Sub SetReadyToListStatus(ByVal rowNum As Long)

    Dim ws As Worksheet
    Dim itemFolderPath As String

    Set ws = ThisWorkbook.Sheets(WS_INVENTORY)
    itemFolderPath = GetItemPhotoFolderPath(rowNum)
    CreateFolderIfMissing itemFolderPath

    CreateButton ws, rowNum, COL_STATUS, "BTN_READY_" & rowNum, "READY TO LIST", "ReadyToListButtonClick"

End Sub

' =====================================================
' GENERIC BUTTON CREATOR
' =====================================================

Public Function CreateButton(ws As Worksheet, _
                             rowNum As Long, _
                             colNum As Long, _
                             buttonName As String, _
                             buttonText As String, _
                             macroName As String) As Shape

    Dim shp As Shape

    Dim btnLeft As Double
    Dim btnTop As Double
    Dim btnWidth As Double
    Dim btnHeight As Double
    Dim fillColor As Long
    Dim borderColor As Long

    On Error Resume Next
    ws.Shapes(buttonName).Delete
    On Error GoTo 0

    btnHeight = 20
    btnWidth = 100
    
    ' Center button within cell (row height = 25, cell width varies by column)
    ' Vertical center: (25 - 20) / 2 = 2.5
    ' Horizontal center: (cell width - button width) / 2
    btnLeft = ws.Cells(rowNum, colNum).Left + ((ws.Cells(rowNum, colNum).width - btnWidth) / 2)
    btnTop = ws.Cells(rowNum, colNum).Top + 2.5

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnWidth, btnHeight)

    With shp

        .Name = buttonName

        .TextFrame.Characters.text = buttonText

        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter

        .TextFrame.MarginLeft = 2
        .TextFrame.MarginRight = 2
        .TextFrame.MarginTop = 1
        .TextFrame.MarginBottom = 1

        .OnAction = macroName

        .Placement = xlMoveAndSize

        Select Case UCase$(buttonText)
            Case "START"
                ' Emerald #27AE60
                fillColor = RGB(39, 174, 96)
                borderColor = RGB(26, 130, 69)
            Case "EDIT"
                ' Orange #E8622A
                fillColor = RGB(232, 98, 42)
                borderColor = RGB(184, 72, 24)
            Case "COPY FOR AI"
                ' Emerald #27AE60 (initial state)
                fillColor = RGB(39, 174, 96)
                borderColor = RGB(26, 130, 69)
            Case "PASTE DETAILS"
                ' Emerald #27AE60 (initial state)
                fillColor = RGB(39, 174, 96)
                borderColor = RGB(26, 130, 69)
            Case "VIEW ITEM DETAILS", "VIEW ITEM FOLDER", "VIEW ALL FOLDERS"
                ' Amber #F5A623
                fillColor = RGB(245, 166, 35)
                borderColor = RGB(196, 123, 10)
            Case "READY TO LIST"
                ' Emerald #27AE60
                fillColor = RGB(39, 174, 96)
                borderColor = RGB(26, 130, 69)
            Case "LIST ANOTHER"
                ' Emerald #27AE60
                fillColor = RGB(39, 174, 96)
                borderColor = RGB(26, 130, 69)
            Case "VIEW ALL SOLD"
                fillColor = RGB(229, 234, 240)
                borderColor = RGB(196, 206, 219)
            Case Else
                fillColor = RGB(92, 127, 168)
                borderColor = RGB(61, 95, 130)
        End Select

        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = fillColor
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = borderColor
        .Line.Weight = 1.25
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .TextFrame.Characters.Font.Bold = True

    End With

    Set CreateButton = shp

End Function

' =====================================================
' CHANGE BUTTON COLOR
' =====================================================

Public Sub ChangeButtonColor(ByVal ws As Worksheet, ByVal buttonName As String, ByVal fillColor As Long, ByVal borderColor As Long)

    On Error Resume Next
    
    Dim shp As Shape
    
    Set shp = ws.Shapes(buttonName)
    If Not shp Is Nothing Then
        shp.Fill.ForeColor.RGB = fillColor
        shp.Line.ForeColor.RGB = borderColor
    End If
    
    On Error GoTo 0

End Sub

' =====================================================
' DELETE ROW BUTTONS
' =====================================================

Public Sub DeleteRowButtons(ByVal rowNum As Long)

    On Error Resume Next

    Dim ws As Worksheet
    Dim shp As Shape

    Set ws = ThisWorkbook.Sheets(WS_INVENTORY)

    For Each shp In ws.Shapes

        If InStr(shp.Name, "_" & rowNum) > 0 Then
            shp.Delete
        End If

    Next shp

End Sub

' =====================================================
' CLEAR INVENTORY ROW
' =====================================================

Public Sub ClearInventoryRow(ByVal rowNum As Long)

    On Error Resume Next

    Dim ws As Worksheet
    Dim wsData As Worksheet

    Set ws = ThisWorkbook.Sheets(WS_INVENTORY)
    Set wsData = ThisWorkbook.Sheets(WS_DATA)

    ws.Cells(rowNum, COL_ITEM_NAME).ClearContents
    ws.Cells(rowNum, COL_ITEM_PRICE).ClearContents
    ws.Cells(rowNum, COL_DATE_SOLD).ClearContents

    ws.Cells(rowNum, COL_STATUS).ClearContents
    wsData.Rows(rowNum).ClearContents

    DeleteRowButtons rowNum

End Sub

Private Sub RetireInventoryRow(ByVal rowNum As Long)

    Dim ws As Worksheet
    Dim colIdx As Long

    Set ws = ThisWorkbook.Sheets(WS_INVENTORY)
    ClearInventoryRow rowNum
    ws.Cells(rowNum, COL_ITEM_NUMBER).ClearContents
    ws.Cells(rowNum, COL_STATUS).Value = STATUS_SOLD
    For colIdx = COL_ITEM_NUMBER To COL_STATUS
        ws.Cells(rowNum, colIdx).Interior.Color = RGB(217, 217, 217)
    Next colIdx
    ws.Cells(rowNum, COL_STATUS).Font.Color = RGB(128, 128, 128)
    ws.Rows(rowNum).Hidden = True
    
    ' Refresh alternating row colors after hiding row
    RefreshRowColorsAfterChange

End Sub

Private Sub DeleteAllButtonsOnWorksheet(ByVal ws As Worksheet)

    On Error Resume Next
    
    Dim shp As Shape
    Dim shapeIndex As Long

    For shapeIndex = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(shapeIndex)
        If Left$(shp.Name, 4) = "BTN_" Then shp.Delete
    Next shapeIndex
    
    On Error GoTo 0

End Sub

Private Sub ResetInventoryWorksheet()

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets(WS_INVENTORY)
    
    ' Delete all buttons
    DeleteAllButtonsOnWorksheet ws
    
    ' Unhide ALL rows (including empty ones) - must do this BEFORE clearing
    ws.Rows("2:" & ws.Rows.Count).Hidden = False
    
    ' Reset row backgrounds to white
    ws.Rows("2:" & ws.Rows.Count).Interior.Color = RGB(255, 255, 255)
    ws.Rows("2:" & ws.Rows.Count).Interior.ColorIndex = xlNone
    
    ' Clear all data rows
    ws.Rows("2:" & ws.Rows.Count).Clear
    
    ' Set up first empty row with just the item number
    ' NO status text, NO button - button appears only when user types in Column B
    ws.Cells(2, COL_ITEM_NUMBER).Value = "100-001"
    
    ' Apply all formatting including validation, column widths, and protection
    ApplyColumnFormatting
    
    ' Apply new alternating row colors (Light/Pale tint scheme)
    RefreshRowColorsAfterChange

End Sub

Private Sub ResetSoldItemsWorksheet()

    Dim ws As Worksheet

    Set ws = GetSoldItemsWorksheet()
    
    ' Unprotect to allow clearing
    On Error Resume Next
    ws.Unprotect Password:=""
    On Error GoTo 0
    
    DeleteAllButtonsOnWorksheet ws
    ws.Rows("2:" & ws.Rows.Count).Clear
    ws.Columns("K:N").Hidden = True
    ' Note: Column widths are set in GetSoldItemsWorksheet, no AutoFit needed

End Sub

Private Sub ResetDataWorksheet()

    Dim wsData As Worksheet
    Set wsData = ThisWorkbook.Sheets(WS_DATA)
    wsData.Rows("2:" & wsData.Rows.Count).Clear

End Sub

Private Sub ResetSoldDataWorksheet()

    Dim ws As Worksheet

    Set ws = GetSoldDataWorksheet()
    ws.Rows.Clear
    ws.Visible = xlSheetVeryHidden

End Sub

Public Sub ResetWorkbookForFreshTesting()

    On Error GoTo ErrorHandler

    If MsgBox("This will clear INVENTORY rows, SOLD ITEMS rows, DATA rows, SOLD_DATA rows, and test buttons. It will not delete SETTINGS, LOOKUPS, AI_TEMPLATE, code, backups, logs, Description files, or 1 READY TO LIST folders." & vbCrLf & vbCrLf & "Continue?", vbYesNo + vbExclamation, "Reset Workbook Test Data") <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ResetInventoryWorksheet
    ResetSoldItemsWorksheet
    ResetDataWorksheet
    ResetSoldDataWorksheet

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Workbook test data has been reset. INVENTORY is back to Row 2 with item 100-001 ready to start.", vbInformation
    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    HandleError "ResetWorkbookForFreshTesting", Err.Number, Err.Description

End Sub

Public Sub ApplyColumnFormatting()
    Dim wsInv As Worksheet
    Dim wsSold As Worksheet
    Dim row As Long
    
    On Error Resume Next
    Set wsInv = ThisWorkbook.Worksheets(WS_INVENTORY)
    Set wsSold = ThisWorkbook.Worksheets(WS_SOLD_ITEMS)
    
    If Not wsInv Is Nothing Then
        ' === ROW HEIGHTS ===
        ' All rows height = 25
        wsInv.Rows.RowHeight = 25
        
        ' === COLUMN WIDTHS ===
        ' Column A = 15
        wsInv.Columns("A").ColumnWidth = 15
        ' Column B = 50
        wsInv.Columns("B").ColumnWidth = 50
        ' Columns C, D = 16 each
        wsInv.Columns("C").ColumnWidth = 16
        wsInv.Columns("D").ColumnWidth = 16
        ' Columns E, F, G, H, I, J, K = 20 each
        wsInv.Columns("E").ColumnWidth = 20
        wsInv.Columns("F").ColumnWidth = 20
        wsInv.Columns("G").ColumnWidth = 20
        wsInv.Columns("H").ColumnWidth = 20
        wsInv.Columns("I").ColumnWidth = 20
        wsInv.Columns("J").ColumnWidth = 20
        wsInv.Columns("K").ColumnWidth = 20
        
        ' === TEXT FORMATTING ===
        ' Column B: wrap text, left justified, vertical center
        wsInv.Columns("B").WrapText = True
        wsInv.Columns("B").HorizontalAlignment = xlLeft
        wsInv.Columns("B").VerticalAlignment = xlVAlignCenter
        
        ' Columns A, C, D: horizontal center, vertical center
        wsInv.Columns("A").HorizontalAlignment = xlCenter
        wsInv.Columns("A").VerticalAlignment = xlVAlignCenter
        wsInv.Columns("C").HorizontalAlignment = xlCenter
        wsInv.Columns("C").VerticalAlignment = xlVAlignCenter
        wsInv.Columns("C").NumberFormat = "$#,##0.00"
        wsInv.Columns("D").HorizontalAlignment = xlCenter
        wsInv.Columns("D").VerticalAlignment = xlVAlignCenter
        
        ' Row 1: centered text, Mercari brand color
        wsInv.Rows(1).HorizontalAlignment = xlCenter
        wsInv.Rows(1).VerticalAlignment = xlVAlignCenter
        wsInv.Rows(1).Interior.Color = RGB(94, 109, 242)
        wsInv.Rows(1).Font.Color = RGB(255, 255, 255)
        wsInv.Rows(1).Font.Bold = True
        
        ' === LOCKING ===
        ' Unprotect first to allow formatting changes
        wsInv.Unprotect Password:=""
        
        ' Apply date sold validation BEFORE locking (so it works properly)
        ApplyDateSoldValidationToRange wsInv
        
        ' Lock all cells by default
        wsInv.Cells.Locked = True
        
        ' Unlock columns A-D for data entry
        wsInv.Columns("A").Locked = False
        wsInv.Columns("B").Locked = False
        wsInv.Columns("C").Locked = False
        wsInv.Columns("D").Locked = False
        
        ' Columns E-K remain locked (cannot resize, buttons cannot move)
        ' Rows remain locked (cannot resize height)
        
        ' Protect with specific settings
        wsInv.Protect Password:="", _
                     UserInterfaceOnly:=True, _
                     AllowFormattingColumns:=False, _
                     AllowFormattingRows:=False, _
                     AllowInsertingColumns:=False, _
                     AllowInsertingRows:=False, _
                     AllowDeletingColumns:=False, _
                     AllowDeletingRows:=False, _
                     AllowSorting:=True, _
                     AllowFiltering:=True
    End If
    
    If Not wsSold Is Nothing Then
        wsSold.Rows.RowHeight = 25
        wsSold.Columns("A").HorizontalAlignment = xlCenter
        wsSold.Columns("C").HorizontalAlignment = xlCenter
        wsSold.Columns("C").NumberFormat = "$#,##0.00"
        wsSold.Columns("D").HorizontalAlignment = xlCenter
    End If
    On Error GoTo 0
End Sub

Private Function GetSoldItemsWorksheet() As Worksheet

    On Error Resume Next
    Set GetSoldItemsWorksheet = ThisWorkbook.Worksheets(WS_SOLD_ITEMS)
    On Error GoTo 0

    If GetSoldItemsWorksheet Is Nothing Then
        Set GetSoldItemsWorksheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetSoldItemsWorksheet.Name = WS_SOLD_ITEMS
    End If

    If Trim$(CStr(GetSoldItemsWorksheet.Cells(1, 1).Value)) = "" Then
        ' === HEADERS ===
        GetSoldItemsWorksheet.Cells(1, 1).Value = "ITEM NUMBER"
        GetSoldItemsWorksheet.Cells(1, 2).Value = "ITEM NAME"
        GetSoldItemsWorksheet.Cells(1, 3).Value = "ITEM PRICE"
        GetSoldItemsWorksheet.Cells(1, 4).Value = "DATE SOLD"
        GetSoldItemsWorksheet.Cells(1, 5).Value = "EDIT"
        GetSoldItemsWorksheet.Cells(1, 6).Value = "RELIST"
        GetSoldItemsWorksheet.Cells(1, 7).Value = "VIEW ITEM DETAILS"
        GetSoldItemsWorksheet.Cells(1, 8).Value = "VIEW ITEM FOLDER"
        GetSoldItemsWorksheet.Cells(1, 9).Value = "VIEW ALL SOLD"
        GetSoldItemsWorksheet.Cells(1, 10).Value = "STATUS"
        GetSoldItemsWorksheet.Cells(1, 11).Value = "SOURCE INVENTORY ROW"
        GetSoldItemsWorksheet.Cells(1, 12).Value = "DETAILS PATH"
        GetSoldItemsWorksheet.Cells(1, 13).Value = "ITEM FOLDER PATH"
        GetSoldItemsWorksheet.Cells(1, 14).Value = "ALL SOLD PATH"
        
        Dim wsSold As Worksheet
        Set wsSold = GetSoldItemsWorksheet
        
        ' === ROW HEIGHTS ===
        ' All rows height = 25
        wsSold.Rows.RowHeight = 25
        
        ' === COLUMN WIDTHS ===
        ' Column A = 15
        wsSold.Columns("A").ColumnWidth = 15
        ' Column B = 50
        wsSold.Columns("B").ColumnWidth = 50
        ' Columns C, D = 16 each
        wsSold.Columns("C").ColumnWidth = 16
        wsSold.Columns("D").ColumnWidth = 16
        ' Columns F, G, H, I = 20 each
        wsSold.Columns("F").ColumnWidth = 20
        wsSold.Columns("G").ColumnWidth = 20
        wsSold.Columns("H").ColumnWidth = 20
        wsSold.Columns("I").ColumnWidth = 20
        
        ' === TEXT FORMATTING ===
        ' Row 1: Centered horizontally and vertically, bold
        wsSold.Rows(1).HorizontalAlignment = xlCenter
        wsSold.Rows(1).VerticalAlignment = xlVAlignCenter
        wsSold.Rows(1).Font.Bold = True
        
        ' Columns A-D: Vertical center for all rows
        wsSold.Columns("A").VerticalAlignment = xlVAlignCenter
        wsSold.Columns("B").VerticalAlignment = xlVAlignCenter
        wsSold.Columns("C").VerticalAlignment = xlVAlignCenter
        wsSold.Columns("D").VerticalAlignment = xlVAlignCenter
        
        ' Column B: Left justified, wrap text
        wsSold.Columns("B").HorizontalAlignment = xlLeft
        wsSold.Columns("B").WrapText = True
        
        ' Columns A, C, D: Horizontally centered
        wsSold.Columns("A").HorizontalAlignment = xlCenter
        wsSold.Columns("C").HorizontalAlignment = xlCenter
        wsSold.Columns("D").HorizontalAlignment = xlCenter
        
        ' === LOCKING ===
        ' Unprotect first to allow formatting changes
        wsSold.Unprotect Password:=""
        
        ' Lock all cells by default
        wsSold.Cells.Locked = True
        
        ' Unlock columns A-E for data entry (A-D data, E is hidden but editable)
        wsSold.Columns("A").Locked = False
        wsSold.Columns("B").Locked = False
        wsSold.Columns("C").Locked = False
        wsSold.Columns("D").Locked = False
        wsSold.Columns("E").Locked = False
        
        ' Columns F-I remain locked (cannot resize, buttons cannot move)
        ' Rows remain locked (cannot resize height)
        
        ' Protect with specific settings
        wsSold.Protect Password:="", _
                     UserInterfaceOnly:=True, _
                     AllowFormattingColumns:=False, _
                     AllowFormattingRows:=False, _
                     AllowInsertingColumns:=False, _
                     AllowInsertingRows:=False, _
                     AllowDeletingColumns:=False, _
                     AllowDeletingRows:=False, _
                     AllowSorting:=True, _
                     AllowFiltering:=True
        
        ' Hide columns
        wsSold.Columns("E").Hidden = True
        wsSold.Columns("K:N").Hidden = True
    End If

End Function

Private Sub ApplySoldItemsRowFormatting(ByVal wsSold As Worksheet, ByVal soldRow As Long)
    ' Applies formatting to a specific row in SOLD ITEMS worksheet
    
    ' Row height = 25
    wsSold.Rows(soldRow).RowHeight = 25
    
    ' Columns A-D: Vertical center
    wsSold.Cells(soldRow, 1).VerticalAlignment = xlVAlignCenter
    wsSold.Cells(soldRow, 2).VerticalAlignment = xlVAlignCenter
    wsSold.Cells(soldRow, 3).VerticalAlignment = xlVAlignCenter
    wsSold.Cells(soldRow, 4).VerticalAlignment = xlVAlignCenter
    
    ' Column B: Left justified, wrap text
    wsSold.Cells(soldRow, 2).HorizontalAlignment = xlLeft
    wsSold.Cells(soldRow, 2).WrapText = True
    
    ' Columns A, C, D: Horizontally centered
    wsSold.Cells(soldRow, 1).HorizontalAlignment = xlCenter
    wsSold.Cells(soldRow, 3).HorizontalAlignment = xlCenter
    wsSold.Cells(soldRow, 4).HorizontalAlignment = xlCenter
    
End Sub

Private Function GetSoldDataWorksheet() As Worksheet

    On Error Resume Next
    Set GetSoldDataWorksheet = ThisWorkbook.Worksheets("SOLD_DATA")
    On Error GoTo 0

    If GetSoldDataWorksheet Is Nothing Then
        Set GetSoldDataWorksheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetSoldDataWorksheet.Name = "SOLD_DATA"
        GetSoldDataWorksheet.Visible = xlSheetVeryHidden
    End If

End Function

Private Sub ArchiveSoldDataRow(ByVal inventoryRow As Long, ByVal soldRow As Long)

    Dim wsSoldData As Worksheet
    Dim wsData As Worksheet

    If inventoryRow <= 0 Then Exit Sub
    If soldRow <= 0 Then Exit Sub

    Set wsSoldData = GetSoldDataWorksheet()
    Set wsData = ThisWorkbook.Sheets(WS_DATA)
    wsSoldData.Rows(soldRow).ClearContents
    wsData.Rows(inventoryRow).Copy Destination:=wsSoldData.Rows(soldRow)

End Sub

Private Function UniquePath(ByVal targetPath As String) As String

    Dim folderPath As String
    Dim fileName As String
    Dim baseName As String
    Dim extensionName As String

    If Dir(targetPath, vbDirectory) = "" And Dir(targetPath) = "" Then
        UniquePath = targetPath
        Exit Function
    End If

    If InStrRev(targetPath, "\") = 0 Then
        UniquePath = targetPath
        Exit Function
    End If

    folderPath = Left$(targetPath, InStrRev(targetPath, "\") - 1)
    fileName = Mid$(targetPath, InStrRev(targetPath, "\") + 1)

    If InStrRev(fileName, ".") > 0 Then
        baseName = Left$(fileName, InStrRev(fileName, ".") - 1)
        extensionName = Mid$(fileName, InStrRev(fileName, "."))
    Else
        baseName = fileName
        extensionName = ""
    End If

    UniquePath = folderPath & "\" & baseName & " - sold " & Format(Now, "yyyymmdd-hhnnss") & extensionName

End Function

Private Function MoveFileToFolder(ByVal sourcePath As String, ByVal targetFolder As String) As String

    Dim targetPath As String

    If Trim$(sourcePath) = "" Then Exit Function
    If Dir(sourcePath) = "" Then Exit Function

    CreateFolderIfMissing targetFolder
    targetPath = targetFolder & "\" & Mid$(sourcePath, InStrRev(sourcePath, "\") + 1)
    targetPath = UniquePath(targetPath)
    Name sourcePath As targetPath
    MoveFileToFolder = targetPath

End Function

Private Function MoveFolderToFolder(ByVal sourceFolder As String, ByVal targetParentFolder As String) As String

    Dim targetPath As String

    If Trim$(sourceFolder) = "" Then Exit Function
    If Dir(sourceFolder, vbDirectory) = "" Then Exit Function

    CreateFolderIfMissing targetParentFolder
    targetPath = targetParentFolder & "\" & Mid$(sourceFolder, InStrRev(sourceFolder, "\") + 1)
    targetPath = UniquePath(targetPath)
    Name sourceFolder As targetPath
    MoveFolderToFolder = targetPath

End Function

Private Function GetSoldDocxFolderPath() As String

    Dim soldRoot As String

    soldRoot = GetSettingValue("SOLD_FOLDER")
    If Trim$(soldRoot) = "" Then soldRoot = ThisWorkbook.Path & "\3 SOLD"
    GetSoldDocxFolderPath = soldRoot & "\Description Files"
    CreateFolderIfMissing GetSoldDocxFolderPath

End Function

Private Function GetSoldPhotosFolderPath() As String

    Dim soldRoot As String

    soldRoot = GetSettingValue("SOLD_FOLDER")
    If Trim$(soldRoot) = "" Then soldRoot = ThisWorkbook.Path & "\3 SOLD"
    GetSoldPhotosFolderPath = soldRoot & "\Item Folders"
    CreateFolderIfMissing GetSoldPhotosFolderPath

End Function

Private Function AddSoldButton(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal colNum As Long, ByVal buttonName As String, ByVal buttonText As String, ByVal macroName As String, Optional ByVal greenButton As Boolean = False) As Shape

    Dim shp As Shape
    Dim btnLeft As Double
    Dim btnTop As Double
    Dim btnWidth As Double
    Dim btnHeight As Double
    Dim fillColor As Long
    Dim borderColor As Long
    Dim fontColor As Long
    
    ' Match INVENTORY worksheet button sizing exactly
    btnHeight = 20
    btnLeft = ws.Cells(rowNum, colNum).Left + 5
    btnTop = ws.Cells(rowNum, colNum).Top + 3
    btnWidth = ws.Cells(rowNum, colNum).Width - 10
    If btnWidth < 20 Then btnWidth = 20
    
    ' Delete existing button
    On Error Resume Next
    ws.Shapes(buttonName).Delete
    On Error GoTo 0
    
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnWidth, btnHeight)
    
    ' Set colors based on button type
    Select Case UCase$(buttonText)
        Case "EDIT"
            fillColor = RGB(229, 234, 240)
            borderColor = RGB(196, 206, 219)
            fontColor = RGB(153, 153, 153)
        Case "RELIST"
            ' Emerald #27AE60 with border #1A8245
            fillColor = RGB(39, 174, 96)
            borderColor = RGB(26, 130, 69)
            fontColor = RGB(255, 255, 255)
        Case "VIEW ITEM DETAILS", "VIEW ITEM FOLDER", "VIEW ALL SOLD"
            ' Light gray #F4F4F4 with text/border #888888
            fillColor = RGB(244, 244, 244)
            borderColor = RGB(136, 136, 136)
            fontColor = RGB(136, 136, 136)
        Case Else
            fillColor = RGB(92, 127, 168)
            borderColor = RGB(61, 95, 130)
            fontColor = RGB(255, 255, 255)
    End Select
    
    With shp
        .Name = buttonName
        .TextFrame.Characters.text = buttonText
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .TextFrame.MarginLeft = 2
        .TextFrame.MarginRight = 2
        .TextFrame.MarginTop = 1
        .TextFrame.MarginBottom = 1
        .OnAction = macroName
        .Placement = xlMoveAndSize
        
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = fillColor
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = borderColor
        .Line.Weight = 1.25
        .TextFrame.Characters.Font.Color = fontColor
        .TextFrame.Characters.Font.Bold = True
    End With
    
    Set AddSoldButton = shp

End Function

Private Sub CreateSoldItemsButtons(ByVal soldRow As Long)

    Dim wsSold As Worksheet

    Set wsSold = ThisWorkbook.Worksheets(WS_SOLD_ITEMS)
    
    ' Column E (5) - Hidden column, no button
    ' Column F (6) - RELIST button (emerald green)
    AddSoldButton wsSold, soldRow, 6, "BTN_RELIST_" & soldRow, "RELIST", "SellAnotherButtonClick", True
    ' Column G (7) - VIEW ITEM DETAILS button (light gray)
    AddSoldButton wsSold, soldRow, 7, "BTN_SOLD_DETAILS_" & soldRow, "VIEW ITEM DETAILS", "SoldViewDetailsButtonClick"
    ' Column H (8) - VIEW ITEM FOLDER button (light gray)
    AddSoldButton wsSold, soldRow, 8, "BTN_SOLD_FOLDER_" & soldRow, "VIEW ITEM FOLDER", "SoldViewFolderButtonClick"
    ' Column I (9) - VIEW ALL SOLD button (light gray)
    AddSoldButton wsSold, soldRow, 9, "BTN_ALL_SOLD_" & soldRow, "VIEW ALL SOLD", "SoldViewAllButtonClick"

End Sub

Private Function NextSoldItemsRow(ByVal wsSold As Worksheet) As Long

    NextSoldItemsRow = wsSold.Cells(wsSold.Rows.Count, 1).End(xlUp).row + 1
    If NextSoldItemsRow < 2 Then NextSoldItemsRow = 2

End Function

Private Function IsValidSoldDateValue(ByVal soldValue As Variant) As Boolean

    If Trim$(CStr(soldValue)) = "" Then Exit Function
    If IsDate(soldValue) Then IsValidSoldDateValue = True

End Function

Private Sub MoveInventoryRowToSoldItems(ByVal inventoryRow As Long, ByVal wsSold As Worksheet)

    Dim wsInventory As Worksheet
    Dim soldRow As Long
    Dim detailsPath As String
    Dim itemFolderPath As String
    Dim soldDetailsPath As String
    Dim soldItemFolderPath As String
    Dim soldPhotosPath As String

    Set wsInventory = ThisWorkbook.Sheets(WS_INVENTORY)
    soldRow = NextSoldItemsRow(wsSold)
    detailsPath = GetDetailsDocxPathForInventoryRow(inventoryRow)
    itemFolderPath = GetItemPhotoFolderPath(inventoryRow)
    soldDetailsPath = MoveFileToFolder(detailsPath, GetSoldDocxFolderPath())
    soldPhotosPath = GetSoldPhotosFolderPath()
    soldItemFolderPath = MoveFolderToFolder(itemFolderPath, soldPhotosPath)

    wsSold.Cells(soldRow, 1).Value = wsInventory.Cells(inventoryRow, COL_ITEM_NUMBER).Value
    wsSold.Cells(soldRow, 2).Value = wsInventory.Cells(inventoryRow, COL_ITEM_NAME).Value
    wsSold.Cells(soldRow, 3).Value = wsInventory.Cells(inventoryRow, COL_ITEM_PRICE).Value
    wsSold.Cells(soldRow, 4).Value = CDate(wsInventory.Cells(inventoryRow, COL_DATE_SOLD).Value)
    wsSold.Cells(soldRow, 10).Value = STATUS_SOLD
    wsSold.Cells(soldRow, 11).Value = inventoryRow
    wsSold.Cells(soldRow, 12).Value = soldDetailsPath
    wsSold.Cells(soldRow, 13).Value = soldItemFolderPath
    wsSold.Cells(soldRow, 14).Value = soldPhotosPath
    ArchiveSoldDataRow inventoryRow, soldRow

    wsSold.Cells(soldRow, 4).NumberFormat = "m/d/yyyy"
    
    ' Apply alternating gray row colors
    ' Darker gray RGB(217, 217, 217) for odd rows, lighter gray RGB(238, 238, 238) for even rows
    Dim rowColor As Long
    If soldRow Mod 2 = 1 Then
        rowColor = RGB(217, 217, 217)  ' Darker gray
    Else
        rowColor = RGB(238, 238, 238)  ' Lighter gray
    End If
    
    Dim colIdx As Long
    For colIdx = 1 To 10
        wsSold.Cells(soldRow, colIdx).Interior.Color = rowColor
        wsSold.Cells(soldRow, colIdx).Font.Color = RGB(153, 153, 153)
    Next colIdx
    
    wsSold.Cells(soldRow, 10).Font.Bold = True
    wsSold.Cells(soldRow, 10).HorizontalAlignment = xlCenter
    wsSold.Cells(soldRow, 10).VerticalAlignment = xlCenter
    wsSold.Columns("E").Hidden = True
    CreateSoldItemsButtons soldRow
    
    ' Apply SOLD ITEMS specific formatting for this row
    ApplySoldItemsRowFormatting wsSold, soldRow

    RetireInventoryRow inventoryRow

End Sub

Public Function ProcessSoldItemsOnClose() As Long

    On Error GoTo ErrorHandler

    Dim wsInventory As Worksheet
    Dim wsSold As Worksheet
    Dim lastRow As Long
    Dim rowNum As Long
    Dim cellValue As Variant

    Set wsInventory = ThisWorkbook.Sheets(WS_INVENTORY)
    Set wsSold = GetSoldItemsWorksheet()
    
    ' Unprotect SOLD ITEMS to allow adding rows
    On Error Resume Next
    wsSold.Unprotect Password:=""
    On Error GoTo 0

    lastRow = wsInventory.Cells(wsInventory.Rows.Count, COL_ITEM_NUMBER).End(xlUp).row

    For rowNum = lastRow To 2 Step -1
        cellValue = wsInventory.Cells(rowNum, COL_DATE_SOLD).Value
        If IsValidSoldDateValue(cellValue) Then
            MoveInventoryRowToSoldItems rowNum, wsSold
            ProcessSoldItemsOnClose = ProcessSoldItemsOnClose + 1
        End If
    Next rowNum

    If ProcessSoldItemsOnClose > 0 Then
        ' Column widths: A=15, B=50, C=16, D=16, F=20, G=20, H=20, I=20
        wsSold.Columns("A").ColumnWidth = 15
        wsSold.Columns("B").ColumnWidth = 50
        wsSold.Columns("C").ColumnWidth = 16
        wsSold.Columns("D").ColumnWidth = 16
        wsSold.Columns("F").ColumnWidth = 20
        wsSold.Columns("G").ColumnWidth = 20
        wsSold.Columns("H").ColumnWidth = 20
        wsSold.Columns("I").ColumnWidth = 20
        wsSold.Columns("E").Hidden = True
        wsSold.Columns("K:N").Hidden = True
        
        ' All row heights = 25
        wsSold.Rows.RowHeight = 25
        
        ' Row 1 formatting: centered, bold
        wsSold.Rows(1).HorizontalAlignment = xlCenter
        wsSold.Rows(1).VerticalAlignment = xlVAlignCenter
        wsSold.Rows(1).Font.Bold = True
        
        ' Lock columns F-I and all rows
        wsSold.Unprotect Password:=""
        wsSold.Cells.Locked = True
        wsSold.Columns("A").Locked = False
        wsSold.Columns("B").Locked = False
        wsSold.Columns("C").Locked = False
        wsSold.Columns("D").Locked = False
        wsSold.Columns("E").Locked = False
        wsSold.Protect Password:="", _
                     UserInterfaceOnly:=True, _
                     AllowFormattingColumns:=False, _
                     AllowFormattingRows:=False, _
                     AllowInsertingColumns:=False, _
                     AllowInsertingRows:=False, _
                     AllowDeletingColumns:=False, _
                     AllowDeletingRows:=False, _
                     AllowSorting:=True, _
                     AllowFiltering:=True
        
        EnsureOnePendingInventoryRow
    End If
    
    ' Refresh alternating row colors after processing sold items
    RefreshRowColorsAfterChange
    
    ' Show popup if items were moved
    If ProcessSoldItemsOnClose > 0 Then
        MsgBox CStr(ProcessSoldItemsOnClose) & " item(s) have been moved to the SOLD ITEMS worksheet.", vbInformation, "Items Sold"
    End If

    Exit Function

ErrorHandler:
    HandleError "ProcessSoldItemsOnClose", Err.Number, Err.Description

End Function

Public Sub ApplyDateSoldValidation()

    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim validationRange As Range

    Set ws = ThisWorkbook.Sheets(WS_INVENTORY)
    Set validationRange = ws.Range(ws.Cells(2, COL_DATE_SOLD), ws.Cells(ws.Rows.Count, COL_DATE_SOLD))

    ' Clear any existing validation
    validationRange.Validation.Delete
    
    ' Add informational message only (no validation restriction)
    ' This allows any text input while still showing helpful popup
    validationRange.Validation.Add Type:=xlValidateInputOnly
    validationRange.Validation.InputTitle = "Date Sold"
    validationRange.Validation.InputMessage = "Enter the sold date in any normal Excel date format." & vbCrLf & vbCrLf & "Examples: 5/18/26, May 18, 18-May, or 18-May-2026." & vbCrLf & vbCrLf & "Excel will reformat it after entry. Leave blank if the item has not sold."

    MsgBox "Date Sold validation has been applied to the Inventory worksheet.", vbInformation
    Exit Sub

ErrorHandler:
    HandleError "ApplyDateSoldValidation", Err.Number, Err.Description

End Sub

Private Sub ApplyDateSoldValidationToRange(ByVal ws As Worksheet)
    ' Applies date sold validation to column D without showing MsgBox
    ' Called internally from ApplyColumnFormatting
    
    On Error GoTo ErrorHandler
    
    Dim validationRange As Range
    Set validationRange = ws.Range(ws.Cells(2, COL_DATE_SOLD), ws.Cells(ws.Rows.Count, COL_DATE_SOLD))
    
    ' Clear any existing validation
    validationRange.Validation.Delete
    
    ' Add informational message only (no validation restriction)
    ' This allows any text input while still showing helpful popup
    validationRange.Validation.Add Type:=xlValidateInputOnly
    validationRange.Validation.InputTitle = "Date Sold"
    validationRange.Validation.InputMessage = "Enter the sold date in any normal Excel date format." & vbCrLf & vbCrLf & "Examples: 5/18/26, May 18, 18-May, or 18-May-2026." & vbCrLf & vbCrLf & "Excel will reformat it after entry. Leave blank if the item has not sold."
    
    Exit Sub
    
ErrorHandler:
    ' Silent error handling for internal use
    Debug.Print "Error in ApplyDateSoldValidationToRange: " & Err.Description
    
End Sub

' =====================================================
' START BUTTON
' =====================================================

Public Sub StartButtonClick()

    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim shp As Shape
    Dim rowNum As Long

    Set ws = ThisWorkbook.Sheets(WS_INVENTORY)

    Set shp = ws.Shapes(Application.Caller)

    rowNum = shp.TopLeftCell.row
    OpenItemEditor rowNum

    CreateEditAndAIButtons rowNum

    Exit Sub

ErrorHandler:
    HandleError "StartButtonClick", Err.Number, Err.Description

End Sub

' =====================================================
' EDIT BUTTON
' =====================================================

Public Sub EditButtonClick()

    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim shp As Shape
    Dim rowNum As Long

    Set ws = ThisWorkbook.Sheets(WS_INVENTORY)
    Set shp = ws.Shapes(Application.Caller)
    rowNum = shp.TopLeftCell.row

    OpenItemEditor rowNum

    Exit Sub

ErrorHandler:
    HandleError "EditButtonClick", Err.Number, Err.Description

End Sub

' =====================================================
' COPY FOR AI BUTTON
' =====================================================

Public Sub CopyForAIButtonClick()

    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim shp As Shape
    Dim rowNum As Long

    Dim promptText As String

    Set ws = ThisWorkbook.Sheets(WS_INVENTORY)

    Set shp = ws.Shapes(Application.Caller)

    rowNum = CLng(Split(shp.Name, "_")(2))

    promptText = BuildMercariAIPrompt(rowNum)

    CopyTextToClipboard promptText

    CreateDetailsButton rowNum
    
    ' COPY FOR AI changes to Steel Blue #3D7BC5 (after click)
    ChangeButtonColor ws, "BTN_AI_" & rowNum, RGB(61, 123, 197), RGB(42, 90, 158)
    ' PASTE DETAILS changes to Emerald #27AE60 (active/waiting state)
    ChangeButtonColor ws, "BTN_DETAILS_" & rowNum, RGB(39, 174, 96), RGB(26, 130, 69)

    MsgBox _
        "The AI prompt for your item has been copied to your clipboard." & vbCrLf & vbCrLf & _
        "Here are your next steps:" & vbCrLf & vbCrLf & _
        "1 - Go to your preferred AI Chat, such as ChatGPT or Claude." & vbCrLf & _
        "2 - Click in the New Chat text box." & vbCrLf & _
        "3 - Right-click and choose Paste." & vbCrLf & _
        "4 - Press Enter." & vbCrLf & _
        "5 - Your AI Chat will create a table with the item information organized." & vbCrLf & _
        "6 - Click Copy, Copy Table, or a similar option." & vbCrLf & _
        "7 - Return to this worksheet and click PASTE DETAILS." & vbCrLf & vbCrLf & _
        "NOTE: It may take several seconds before you receive confirmation that the details were successfully saved. This is normal.", _
        vbInformation

    Exit Sub

ErrorHandler:
    HandleError "CopyForAIButtonClick", Err.Number, Err.Description

End Sub

' =====================================================
' PASTE DETAILS BUTTON
' =====================================================

Public Sub PasteDetailsButtonClick()

    On Error GoTo ErrorHandler

    MsgBox "Standby - this process takes a few moments to save details and format your Word documents.", vbInformation + vbOKOnly, "Process Initializing"

    Dim ws As Worksheet
    Dim shp As Shape
    Dim rowNum As Long
    Dim detailsPath As String
    Dim currentStep As String
    Dim existingDetailsPath As String

    currentStep = "Get inventory worksheet"
    Set ws = ThisWorkbook.Sheets(WS_INVENTORY)

    currentStep = "Get clicked button"
    Set shp = ws.Shapes(Application.Caller)

    currentStep = "Read row from button"
    rowNum = shp.TopLeftCell.row

    currentStep = "Check existing details DOCX"
    existingDetailsPath = GetDetailsDocxPathForInventoryRow(rowNum)
    If Dir(existingDetailsPath) <> "" Then
        If MsgBox( _
            "A Details file for this item has already been created." & vbCrLf & vbCrLf & _
            "Overwrite = move the existing DOCX files into ARCHIVED FILES folders, then create new details." & vbCrLf & _
            "Cancel = stop without changing the existing details files." & vbCrLf & vbCrLf & _
            "Click OK to Overwrite or Cancel to stop.", _
            vbQuestion + vbOKCancel, _
            "Replace Existing Details?") = vbCancel Then
            Exit Sub
        End If
        currentStep = "Archive existing details DOCX"
        ArchiveExistingDetailsDocx rowNum
    End If

    currentStep = "Save AI response DOCX"
    detailsPath = SaveAIResponseDetailsDocx(rowNum)
    currentStep = "Create View Details button"
    CreateViewDetailsButton rowNum

    currentStep = "Create View Folder button"
    CreateDetailsFolderButton rowNum

    currentStep = "Create View All Folders button"
    CreateViewAllFoldersButton rowNum

    currentStep = "Set Ready To List status"
    SetReadyToListStatus rowNum
    
    currentStep = "Change button colors"
    ' PASTE DETAILS changes to Rose #C94B8A (completed state)
    ChangeButtonColor ws, "BTN_DETAILS_" & rowNum, RGB(201, 75, 138), RGB(154, 48, 104)
    ' VIEW buttons change to Amber #F5A623 (active state)
    ChangeButtonColor ws, "BTN_VIEW_DETAILS_" & rowNum, RGB(245, 166, 35), RGB(196, 123, 10)
    ChangeButtonColor ws, "BTN_FOLDER_" & rowNum, RGB(245, 166, 35), RGB(196, 123, 10)
    ChangeButtonColor ws, "BTN_ALL_FOLDERS_" & rowNum, RGB(245, 166, 35), RGB(196, 123, 10)

    currentStep = "Show success message"
    MsgBox _
        "Details successfully saved to:" & vbCrLf & vbCrLf & _
        detailsPath, _
        vbInformation

    Exit Sub

ErrorHandler:
    HandleError "PasteDetailsButtonClick - " & currentStep & " -> " & Err.Source, Err.Number, Err.Description

End Sub

' =====================================================
' VIEW DETAILS BUTTON
' =====================================================

Public Sub ViewDetailsButtonClick()

    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim shp As Shape
    Dim rowNum As Long
    Dim detailsPath As String
    Dim itemFolderDocx As String

    Set ws = ThisWorkbook.Sheets(WS_INVENTORY)
    Set shp = ws.Shapes(Application.Caller)
    rowNum = shp.TopLeftCell.row
    
    itemFolderDocx = GetItemFolderDocxPathForInventoryRow(rowNum)
    If Dir(itemFolderDocx) <> "" Then
        ThisWorkbook.FollowHyperlink itemFolderDocx
    Else
        detailsPath = GetDetailsDocxPathForInventoryRow(rowNum)
        If Dir(detailsPath) = "" Then Err.Raise vbObjectError + 6401, "ViewDetailsButtonClick", "The details DOCX could not be found:" & vbCrLf & vbCrLf & detailsPath
        ThisWorkbook.FollowHyperlink detailsPath
    End If
    Exit Sub

ErrorHandler:
    HandleError "ViewDetailsButtonClick", Err.Number, Err.Description

End Sub

' =====================================================
' VIEW FOLDER BUTTON
' =====================================================

Public Sub ViewFolderButtonClick()

    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim shp As Shape
    Dim rowNum As Long

    Set ws = ThisWorkbook.Sheets(WS_INVENTORY)

    Set shp = ws.Shapes(Application.Caller)

    rowNum = CLng(Split(shp.Name, "_")(2))

    OpenItemPhotoFolder rowNum

    Exit Sub

ErrorHandler:
    HandleError "ViewFolderButtonClick", Err.Number, Err.Description

End Sub

' =====================================================
' VIEW ALL FOLDERS BUTTON
' =====================================================

Public Sub ViewAllFoldersButtonClick()

    On Error GoTo ErrorHandler

    OpenAllPhotoFolders

    Exit Sub

ErrorHandler:
    HandleError "ViewAllFoldersButtonClick", Err.Number, Err.Description

End Sub

' =====================================================
' READY TO LIST BUTTON
' =====================================================

Public Sub ReadyToListButtonClick()

    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim rowNum As Long
    Dim itemFolderPath As String
    Dim detailsPath As String
    Dim itemFolderDocx As String

    Set ws = ThisWorkbook.Sheets(WS_INVENTORY)
    rowNum = GetCallerRow()
    itemFolderPath = GetItemPhotoFolderPath(rowNum)
    detailsPath = GetDetailsDocxPathForInventoryRow(rowNum)
    itemFolderDocx = GetItemFolderDocxPathForInventoryRow(rowNum)

    If Dir(itemFolderPath, vbDirectory) <> "" Then
        ThisWorkbook.FollowHyperlink itemFolderPath
    End If

    If Dir(itemFolderDocx) <> "" Then
        ThisWorkbook.FollowHyperlink itemFolderDocx
    ElseIf Dir(detailsPath) <> "" Then
        ThisWorkbook.FollowHyperlink detailsPath
    End If

    Exit Sub

ErrorHandler:
    HandleError "ReadyToListButtonClick", Err.Number, Err.Description

End Sub

Private Function GetCallerRow() As Long

    Dim ws As Worksheet
    Dim shp As Shape

    Set ws = ActiveSheet
    Set shp = ws.Shapes(Application.Caller)
    GetCallerRow = shp.TopLeftCell.row

End Function

Private Function IsRetiredInventoryRow(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean

    IsRetiredInventoryRow = (UCase$(Trim$(CStr(ws.Cells(rowNum, COL_STATUS).Value))) = UCase$(STATUS_SOLD) Or _
                             ws.Cells(rowNum, COL_ITEM_NUMBER).Interior.Color = RGB(217, 217, 217) Or _
                             ws.Cells(rowNum, COL_STATUS).Interior.Color = RGB(217, 217, 217))

End Function

Private Function GetNextOpenInventoryRow() As Long

    Dim ws As Worksheet
    Dim rowNum As Long

    Set ws = ThisWorkbook.Sheets(WS_INVENTORY)
    rowNum = 2
    Do While IsRetiredInventoryRow(ws, rowNum) Or _
             Trim$(CStr(ws.Cells(rowNum, COL_ITEM_NAME).Value)) <> "" Or _
             Trim$(CStr(ws.Cells(rowNum, COL_ITEM_PRICE).Value)) <> "" Or _
             Trim$(CStr(ws.Cells(rowNum, COL_DATE_SOLD).Value)) <> ""
        rowNum = rowNum + 1
    Loop
    GetNextOpenInventoryRow = rowNum

End Function

Private Sub CopyRelistDataRow(ByVal sourceRow As Long, ByVal targetRow As Long)

    Dim wsSoldData As Worksheet
    Dim wsData As Worksheet

    If sourceRow <= 0 Then Exit Sub
    If targetRow <= 0 Then Exit Sub
    Set wsSoldData = GetSoldDataWorksheet()
    Set wsData = ThisWorkbook.Sheets(WS_DATA)
    wsData.Rows(targetRow).ClearContents
    wsSoldData.Rows(sourceRow).Copy Destination:=wsData.Rows(targetRow)

End Sub

Public Sub SoldViewDetailsButtonClick()

    On Error GoTo ErrorHandler

    Dim wsSold As Worksheet
    Dim soldRow As Long
    Dim detailsPath As String

    Set wsSold = ThisWorkbook.Sheets(WS_SOLD_ITEMS)
    soldRow = GetCallerRow()
    detailsPath = CStr(wsSold.Cells(soldRow, 12).Value)
    If Trim$(detailsPath) = "" Or Dir(detailsPath) = "" Then Err.Raise vbObjectError + 6501, "SoldViewDetailsButtonClick", "The sold item details DOCX could not be found:" & vbCrLf & vbCrLf & detailsPath
    ThisWorkbook.FollowHyperlink detailsPath
    Exit Sub

ErrorHandler:
    HandleError "SoldViewDetailsButtonClick", Err.Number, Err.Description

End Sub

Public Sub SoldViewFolderButtonClick()

    On Error GoTo ErrorHandler

    Dim wsSold As Worksheet
    Dim soldRow As Long
    Dim folderPath As String

    Set wsSold = ThisWorkbook.Sheets(WS_SOLD_ITEMS)
    soldRow = GetCallerRow()
    folderPath = CStr(wsSold.Cells(soldRow, 13).Value)
    If Trim$(folderPath) = "" Or Dir(folderPath, vbDirectory) = "" Then Err.Raise vbObjectError + 6502, "SoldViewFolderButtonClick", "The sold item folder could not be found:" & vbCrLf & vbCrLf & folderPath
    ThisWorkbook.FollowHyperlink folderPath
    Exit Sub

ErrorHandler:
    HandleError "SoldViewFolderButtonClick", Err.Number, Err.Description

End Sub

Public Sub SoldViewAllButtonClick()

    On Error GoTo ErrorHandler

    Dim wsSold As Worksheet
    Dim soldRow As Long
    Dim folderPath As String

    Set wsSold = ThisWorkbook.Sheets(WS_SOLD_ITEMS)
    soldRow = GetCallerRow()
    folderPath = CStr(wsSold.Cells(soldRow, 14).Value)
    If Trim$(folderPath) = "" Then folderPath = GetSoldPhotosFolderPath()
    CreateFolderIfMissing folderPath
    ThisWorkbook.FollowHyperlink folderPath
    Exit Sub

ErrorHandler:
    HandleError "SoldViewAllButtonClick", Err.Number, Err.Description

End Sub

Private Function CopyAndRenameItemFolder(ByVal sourceFolderPath As String, ByVal newItemNumber As String, ByVal itemName As String, ByVal soldDetailsPath As String, ByVal newDetailsPath As String) As String

    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim folderName As String
    Dim targetFolderPath As String
    Dim photosRoot As String
    Dim file As Object
    Dim newFileName As String
    Dim oldItemNumber As String
    Dim sourceFolder As Object
    Dim fileCount As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Verify source folder exists
    If Not fso.FolderExists(sourceFolderPath) Then
        MsgBox "Source folder not found: " & vbCrLf & sourceFolderPath, vbExclamation, "Copy Error"
        CopyAndRenameItemFolder = ""
        Exit Function
    End If
    
    ' Check if source folder has files
    Set sourceFolder = fso.GetFolder(sourceFolderPath)
    fileCount = sourceFolder.Files.Count
    If fileCount = 0 Then
        MsgBox "Source folder is empty: " & vbCrLf & sourceFolderPath, vbExclamation, "Copy Error"
    End If
    
    photosRoot = GetSettingValue("PHOTOS_FOLDER")
    If Trim$(photosRoot) = "" Then photosRoot = ThisWorkbook.Path & "\1 READY TO LIST"
    
    If Len(itemName) > 50 Then itemName = Left$(itemName, 50)
    targetFolderPath = photosRoot & "\" & newItemNumber & " - " & itemName
    If Len(targetFolderPath) > 180 Then targetFolderPath = photosRoot & "\" & newItemNumber
    targetFolderPath = UniquePath(targetFolderPath)
    
    ' Copy the entire folder with contents
    fso.CopyFolder sourceFolderPath, targetFolderPath, True
    
    ' Verify copy worked
    If Not fso.FolderExists(targetFolderPath) Then
        MsgBox "Failed to create target folder: " & vbCrLf & targetFolderPath, vbExclamation, "Copy Error"
        CopyAndRenameItemFolder = ""
        Exit Function
    End If
    
    ' Determine old item number from source folder name
    folderName = Mid$(sourceFolderPath, InStrRev(sourceFolderPath, "\") + 1)
    If InStr(folderName, " - ") > 0 Then
        oldItemNumber = Left$(folderName, InStr(folderName, " - ") - 1)
    Else
        oldItemNumber = folderName
    End If
    
    ' Rename files inside the copied folder
    Dim itemFolderDocxPath As String
    For Each file In fso.GetFolder(targetFolderPath).Files
        If InStr(file.Name, oldItemNumber) > 0 Then
            newFileName = Replace(file.Name, oldItemNumber, newItemNumber, 1, 1)
            file.Name = newFileName
            If LCase$(fso.GetExtensionName(newFileName)) = "docx" Then
                itemFolderDocxPath = targetFolderPath & "\" & newFileName
            End If
        End If
    Next file
    
    ' Copy the central .docx file if it exists
    If soldDetailsPath <> "" And fso.FileExists(soldDetailsPath) Then
        Dim destFolder As String
        destFolder = Left$(newDetailsPath, InStrRev(newDetailsPath, "\") - 1)
        CreateFolderIfMissing destFolder
        fso.CopyFile soldDetailsPath, newDetailsPath, True
    End If
    
    ' Update item numbers inside both docx copies
    If itemFolderDocxPath <> "" Then
        UpdateDocxItemNumber itemFolderDocxPath, oldItemNumber, newItemNumber
    End If
    If newDetailsPath <> "" And fso.FileExists(newDetailsPath) Then
        UpdateDocxItemNumber newDetailsPath, oldItemNumber, newItemNumber
    End If
    
    CopyAndRenameItemFolder = targetFolderPath
    Exit Function
    
ErrorHandler:
    MsgBox "Error in CopyAndRenameItemFolder: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           "Source: " & sourceFolderPath & vbCrLf & _
           "Target: " & targetFolderPath, vbCritical, "Copy Error"
    CopyAndRenameItemFolder = ""
    
End Function

Public Sub SellAnotherButtonClick()

    On Error GoTo ErrorHandler

    MsgBox "Standby - this process takes a few moments as we copy folders, rename photos, and update your Word document reference files.", vbInformation + vbOKOnly, "Process Initializing"

    Dim wsSold As Worksheet
    Dim wsInventory As Worksheet
    Dim soldRow As Long
    Dim inventoryRow As Long
    Dim sourceRow As Long
    Dim newItemNumber As String
    Dim soldItemFolderPath As String
    Dim newItemFolderPath As String
    Dim soldDetailsPath As String
    Dim newDetailsPath As String

    Set wsSold = ThisWorkbook.Sheets(WS_SOLD_ITEMS)
    Set wsInventory = ThisWorkbook.Sheets(WS_INVENTORY)
    soldRow = GetCallerRow()
    inventoryRow = GetNextOpenInventoryRow()
    sourceRow = soldRow ' The data on SOLD_DATA sheet is saved at soldRow, matching the sold items row

    If Trim$(CStr(wsInventory.Cells(inventoryRow, COL_ITEM_NUMBER).Value)) = "" Then
        wsInventory.Cells(inventoryRow, COL_ITEM_NUMBER).Value = GetNextItemNumber()
    End If
    newItemNumber = CStr(wsInventory.Cells(inventoryRow, COL_ITEM_NUMBER).Value)
    wsInventory.Cells(inventoryRow, COL_ITEM_NAME).Value = wsSold.Cells(soldRow, 2).Value
    wsInventory.Cells(inventoryRow, COL_ITEM_PRICE).Value = wsSold.Cells(soldRow, 3).Value
    CopyRelistDataRow sourceRow, inventoryRow
    
    soldItemFolderPath = CStr(wsSold.Cells(soldRow, 13).Value)
    soldDetailsPath = CStr(wsSold.Cells(soldRow, 12).Value)
    
    ' Verify the source folder exists (may be old path before folder restructuring)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(soldItemFolderPath) Then
        ' Try to find in new "3 SOLD" location using item number
        Dim soldItemNumber As String
        soldItemNumber = CStr(wsSold.Cells(soldRow, 1).Value)
        Dim alternativePath As String
        alternativePath = GetSoldPhotosFolderPath() & "\" & soldItemNumber
        ' Try with full name pattern
        If Not fso.FolderExists(alternativePath) Then
            Dim itemNameCheck As String
            itemNameCheck = CStr(wsSold.Cells(soldRow, 2).Value)
            If Len(itemNameCheck) > 50 Then itemNameCheck = Left$(itemNameCheck, 50)
            alternativePath = GetSoldPhotosFolderPath() & "\" & soldItemNumber & " - " & itemNameCheck
        End If
        If fso.FolderExists(alternativePath) Then
            soldItemFolderPath = alternativePath
            ' Update the stored path for future reference
            wsSold.Cells(soldRow, 13).Value = soldItemFolderPath
        Else
            MsgBox "Could not find the sold item folder." & vbCrLf & vbCrLf & _
                   "Tried: " & soldItemFolderPath & vbCrLf & vbCrLf & _
                   "And: " & alternativePath, vbExclamation, "Folder Not Found"
            Exit Sub
        End If
    End If
    
    newDetailsPath = GetDetailsDocxPathForInventoryRow(inventoryRow)
    newItemFolderPath = CopyAndRenameItemFolder(soldItemFolderPath, newItemNumber, wsInventory.Cells(inventoryRow, COL_ITEM_NAME).Value, soldDetailsPath, newDetailsPath)
    
    ' Clear columns E through K first to prevent overlay
    DeleteRowButtons inventoryRow
    
    ' 0. Display EDIT button in Column E (COL_START_EDIT)
    CreateButton wsInventory, inventoryRow, COL_START_EDIT, "BTN_EDIT_" & inventoryRow, "EDIT", "EditButtonClick"
    ChangeButtonColor wsInventory, "BTN_EDIT_" & inventoryRow, RGB(224, 123, 58), RGB(184, 88, 32)
    
    ' 1. Display COPY FOR AI button in Column F (COL_COPY_FOR_AI)
    CreateButton wsInventory, inventoryRow, COL_COPY_FOR_AI, "BTN_AI_" & inventoryRow, "COPY FOR AI", "CopyForAIButtonClick"
    ChangeButtonColor wsInventory, "BTN_AI_" & inventoryRow, RGB(0, 170, 204), RGB(0, 122, 153)
    
    ' 2. Display PASTE DETAILS button in Column G (COL_DETAILS)
    CreateButton wsInventory, inventoryRow, COL_DETAILS, "BTN_DETAILS_" & inventoryRow, "PASTE DETAILS", "PasteDetailsButtonClick"
    ChangeButtonColor wsInventory, "BTN_DETAILS_" & inventoryRow, RGB(92, 127, 168), RGB(61, 95, 130)
    
    ' 3. Display VIEW ITEM DETAILS button in Column H (COL_VIEW_DETAILS)
    CreateButton wsInventory, inventoryRow, COL_VIEW_DETAILS, "BTN_VIEW_DETAILS_" & inventoryRow, "VIEW ITEM DETAILS", "ViewDetailsButtonClick"
    ChangeButtonColor wsInventory, "BTN_VIEW_DETAILS_" & inventoryRow, RGB(255, 184, 0), RGB(204, 146, 0)
    
    ' 4. Display VIEW ITEM FOLDER button in Column I (COL_DETAILS_FOLDER)
    CreateButton wsInventory, inventoryRow, COL_DETAILS_FOLDER, "BTN_FOLDER_" & inventoryRow, "VIEW ITEM FOLDER", "ViewFolderButtonClick"
    ChangeButtonColor wsInventory, "BTN_FOLDER_" & inventoryRow, RGB(255, 184, 0), RGB(204, 146, 0)
    
    ' 5. Display VIEW ALL FOLDERS button in Column J (COL_VIEW_ALL_FOLDERS)
    CreateButton wsInventory, inventoryRow, COL_VIEW_ALL_FOLDERS, "BTN_ALL_FOLDERS_" & inventoryRow, "VIEW ALL FOLDERS", "ViewAllFoldersButtonClick"
    ChangeButtonColor wsInventory, "BTN_ALL_FOLDERS_" & inventoryRow, RGB(255, 184, 0), RGB(204, 146, 0)
    
    ' 6. Display READY TO LIST button in Column K (COL_STATUS)
    CreateButton wsInventory, inventoryRow, COL_STATUS, "BTN_READY_" & inventoryRow, "READY TO LIST", "ReadyToListButtonClick"
    ChangeButtonColor wsInventory, "BTN_READY_" & inventoryRow, RGB(61, 191, 140), RGB(42, 143, 101)

    ApplyColumnFormatting

    MsgBox "A new inventory row has been created from the sold item." & vbCrLf & vbCrLf & "New row: " & inventoryRow, vbInformation
    wsInventory.Activate
    wsInventory.Cells(inventoryRow, COL_ITEM_PRICE).Select
    Exit Sub

ErrorHandler:
    HandleError "SellAnotherButtonClick", Err.Number, Err.Description

End Sub

Public Sub SoldEditButtonClick()

    MsgBox "Sold item editing will be added in a later phase. For now, use VIEW DETAILS or VIEW ITEM FOLDER to review the sold item.", vbInformation

End Sub

' =====================================================
' OPEN ITEM EDITOR
' =====================================================

Public Sub OpenItemEditor(ByVal inventoryRow As Long)

    Dim wsInventory As Worksheet

    Dim itemNumber As String
    Dim itemName As String

    Set wsInventory = ThisWorkbook.Worksheets("INVENTORY")

    itemNumber = Trim(wsInventory.Cells(inventoryRow, 1).Value)
    itemName = Trim(wsInventory.Cells(inventoryRow, 2).Value)

    With frmItemEditor

        .CurrentDataRow = inventoryRow

        .itemNumber = itemNumber
        .itemName = itemName

        .Show

    End With

End Sub

' =====================================================
' ALTERNATING ROW COLORS FOR INVENTORY
' =====================================================

Public Sub ApplyAlternatingRowColors()
    ' Applies alternating row colors to visible rows in INVENTORY sheet
    ' Last row is always Light accent (#C5CBFA), alternates with Pale tint (#EEF0FD)
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim row As Long
    Dim visibleRowCount As Long
    Dim lastRowColor As Long
    Dim alternatingColor As Long
    
    ' Light accent #C5CBFA for last row (always)
    lastRowColor = RGB(197, 203, 250)
    ' Pale tint #EEF0FD for alternating rows
    alternatingColor = RGB(238, 240, 253)
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_INVENTORY)
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    ' Find last visible row with data in column A (ITEM #)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' If no data rows (only header), exit
    If lastRow <= 1 Then Exit Sub
    
    ' Clear existing colors from data rows (A to K)
    ws.Range("A2:K" & lastRow).Interior.ColorIndex = xlNone
    
    ' Count visible rows and apply pattern
    visibleRowCount = 0
    
    For row = lastRow To 2 Step -1
        ' Check if row is visible (not hidden)
        If ws.Rows(row).Hidden = False Then
            visibleRowCount = visibleRowCount + 1
            
            ' Odd rows from bottom (1, 3, 5...) get last row color (Light accent)
            ' Even rows from bottom (2, 4, 6...) get alternating color (Pale tint)
            If visibleRowCount Mod 2 = 1 Then
                ws.Range("A" & row & ":K" & row).Interior.Color = lastRowColor
            Else
                ws.Range("A" & row & ":K" & row).Interior.Color = alternatingColor
            End If
        End If
    Next row
    
End Sub

Public Sub RefreshRowColorsAfterChange()
    ' Call this after adding, deleting, or hiding rows
    Application.ScreenUpdating = False
    ApplyAlternatingRowColors
    Application.ScreenUpdating = True
End Sub

' =====================================================
' DEVELOPER: UNHIDE ALL WORKSHEETS
' =====================================================

Public Sub UnhideAllWorksheets()
    ' For development use only - unhides all system worksheets
    
    On Error Resume Next
    
    ThisWorkbook.Worksheets("DATA").Visible = xlSheetVisible
    ThisWorkbook.Worksheets("LOOKUPS").Visible = xlSheetVisible
    ThisWorkbook.Worksheets("SETTINGS").Visible = xlSheetVisible
    ThisWorkbook.Worksheets("TABLES").Visible = xlSheetVisible
    ThisWorkbook.Worksheets("COPYRIGHT_INFO").Visible = xlSheetVisible
    ThisWorkbook.Worksheets("SOLD_DATA").Visible = xlSheetVisible
    
    On Error GoTo 0
    
    MsgBox "All system worksheets are now visible.", vbInformation
    
End Sub

' =====================================================
' STARTUP NAVIGATION
' =====================================================

Public Sub NavigateToStartupWorksheet()
    ' Open to INVENTORY if data exists, otherwise WELCOME
    ' Scroll to top so Row 1 is always visible
    
    Dim wsInventory As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim hasData As Boolean
    Dim i As Long
    
    On Error Resume Next
    Set wsInventory = ThisWorkbook.Worksheets(WS_INVENTORY)
    On Error GoTo 0
    
    If wsInventory Is Nothing Then
        ' Fallback to WELCOME if INVENTORY not found
        Set wsTarget = ThisWorkbook.Worksheets(WS_WELCOME)
        wsTarget.Activate
        wsTarget.Range("A1").Select
        ActiveWindow.ScrollRow = 1
        Exit Sub
    End If
    
    ' Check if there's ANY data in INVENTORY (search all rows for item names in Column B)
    ' Column A always has item numbers, so we check Column B for actual item data
    lastRow = wsInventory.Cells(wsInventory.Rows.Count, COL_ITEM_NAME).End(xlUp).row
    
    ' Check rows 2 through lastRow for any item name
    If lastRow >= 2 Then
        For i = 2 To lastRow
            If Trim$(CStr(wsInventory.Cells(i, COL_ITEM_NAME).Value)) <> "" Then
                hasData = True
                Exit For
            End If
        Next i
    End If
    
    ' Navigate based on data presence
    If hasData Then
        Set wsTarget = wsInventory
        wsInventory.Activate
        ' Select first data row, column A and scroll to top
        wsInventory.Cells(2, COL_ITEM_NUMBER).Select
    Else
        Set wsTarget = ThisWorkbook.Worksheets(WS_WELCOME)
        wsTarget.Activate
        wsTarget.Range("A1").Select
    End If
    
    ' Always scroll to top so Row 1 is visible
    ActiveWindow.ScrollRow = 1
    
End Sub
