Attribute VB_Name = "modDOCX"
Option Explicit

Private Const DOCX_FOLDER_NAME As String = "02 DESCRIPTION FILES"
Private Const CF_UNICODETEXT As Long = 13

#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)
#Else
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
#End If

Private Function SafeFileName(ByVal fileName As String) As String
    Dim invalidChars As Variant
    Dim i As Long
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    SafeFileName = Trim$(CStr(fileName))
    For i = LBound(invalidChars) To UBound(invalidChars)
        SafeFileName = Replace(SafeFileName, CStr(invalidChars(i)), "-")
    Next i
    If SafeFileName = "" Then SafeFileName = "Mercari Details"
End Function

Private Function ShortSafeFileName(ByVal fileName As String, ByVal maxLength As Long) As String
    ShortSafeFileName = SafeFileName(fileName)
    If Len(ShortSafeFileName) > maxLength Then ShortSafeFileName = Left$(ShortSafeFileName, maxLength)
End Function

Private Function GetProjectRootPath() As String
    If Dir(ThisWorkbook.Path & "\" & DOCX_FOLDER_NAME, vbDirectory) <> "" Then
        GetProjectRootPath = ThisWorkbook.Path
    ElseIf Dir(ThisWorkbook.Path & "\Documentation", vbDirectory) <> "" Then
        GetProjectRootPath = ThisWorkbook.Path
    ElseIf InStrRev(ThisWorkbook.Path, "\") > 0 Then
        GetProjectRootPath = Left$(ThisWorkbook.Path, InStrRev(ThisWorkbook.Path, "\") - 1)
    Else
        GetProjectRootPath = ThisWorkbook.Path
    End If
End Function

Private Function GetDOCXFolderPath() As String
    GetDOCXFolderPath = GetProjectRootPath() & "\" & DOCX_FOLDER_NAME
    CreateFolderIfMissing GetDOCXFolderPath
End Function

Private Function GetClipboardText() As String
    Dim dataObject As Object
    On Error Resume Next
    GetClipboardText = GetClipboardUnicodeText()
    On Error GoTo ErrorHandler
    If Trim$(GetClipboardText) <> "" Then Exit Function
    Set dataObject = CreateObject("MSForms.DataObject")
    dataObject.GetFromClipboard
    GetClipboardText = dataObject.GetText
    Exit Function
ErrorHandler:
    Err.Raise vbObjectError + 6301, "GetClipboardText", "Unable to read text from the clipboard. Copy the full AI response table, then try again."
End Function

Private Function GetClipboardUnicodeText() As String
#If VBA7 Then
    Dim hData As LongPtr
    Dim lpData As LongPtr
    Dim textLength As Long
    Dim resultText As String
#Else
    Dim hData As Long
    Dim lpData As Long
    Dim textLength As Long
    Dim resultText As String
#End If
    If IsClipboardFormatAvailable(CF_UNICODETEXT) = 0 Then Exit Function
    If OpenClipboard(0) = 0 Then Exit Function
    hData = GetClipboardData(CF_UNICODETEXT)
    If hData <> 0 Then
        lpData = GlobalLock(hData)
        If lpData <> 0 Then
            textLength = lstrlenW(lpData)
            If textLength > 0 Then
                resultText = String$(textLength, vbNullChar)
#If VBA7 Then
                CopyMemory StrPtr(resultText), lpData, CLngPtr(textLength * 2)
#Else
                CopyMemory StrPtr(resultText), lpData, textLength * 2
#End If
                GetClipboardUnicodeText = resultText
            End If
            GlobalUnlock hData
        End If
    End If
    CloseClipboard
End Function

Private Function GetDetailsDocxPath(ByVal inventoryRow As Long) As String
    Dim ws As Worksheet
    Dim itemNumber As String
    Dim itemName As String
    Set ws = ThisWorkbook.Worksheets(WS_INVENTORY)
    itemNumber = ShortSafeFileName(ws.Cells(inventoryRow, COL_ITEM_NUMBER).Value, 30)
    itemName = ShortSafeFileName(ws.Cells(inventoryRow, COL_ITEM_NAME).Value, 60)
    GetDetailsDocxPath = GetDOCXFolderPath() & "\" & itemNumber & " - " & itemName & ".docx"
End Function

Public Function GetDetailsDocxPathForInventoryRow(ByVal inventoryRow As Long) As String
    GetDetailsDocxPathForInventoryRow = GetDetailsDocxPath(inventoryRow)
End Function

Private Function GetItemFolderDocxPath(ByVal inventoryRow As Long) As String
    Dim ws As Worksheet
    Dim itemFolderPath As String
    Dim itemNumber As String
    Dim itemName As String
    Set ws = ThisWorkbook.Worksheets(WS_INVENTORY)
    itemFolderPath = GetItemPhotoFolderPath(inventoryRow)
    CreateFolderIfMissing itemFolderPath
    itemNumber = ShortSafeFileName(ws.Cells(inventoryRow, COL_ITEM_NUMBER).Value, 30)
    itemName = ShortSafeFileName(ws.Cells(inventoryRow, COL_ITEM_NAME).Value, 60)
    GetItemFolderDocxPath = itemFolderPath & "\" & itemNumber & " - " & itemName & ".docx"
End Function

Private Sub ValidateDocxTargetPath(ByVal filePath As String)
    Dim folderPath As String
    folderPath = Left$(filePath, InStrRev(filePath, "\") - 1)
    If Len(filePath) > 218 Then Err.Raise vbObjectError + 6310, "ValidateDocxTargetPath", "The DOCX file path is too long:" & vbCrLf & vbCrLf & filePath
    If Dir(folderPath, vbDirectory) = "" Then Err.Raise vbObjectError + 6311, "ValidateDocxTargetPath", "The DOCX target folder does not exist:" & vbCrLf & vbCrLf & folderPath
End Sub

Private Function ArchivedDocxPath(ByVal filePath As String) As String
    Dim folderPath As String
    Dim archiveFolderPath As String
    Dim fileName As String
    Dim baseName As String
    Dim extensionName As String
    folderPath = Left$(filePath, InStrRev(filePath, "\") - 1)
    fileName = Mid$(filePath, InStrRev(filePath, "\") + 1)
    archiveFolderPath = folderPath & "\ARCHIVED FILES"
    CreateFolderIfMissing archiveFolderPath
    If InStrRev(fileName, ".") > 0 Then
        baseName = Left$(fileName, InStrRev(fileName, ".") - 1)
        extensionName = Mid$(fileName, InStrRev(fileName, "."))
    Else
        baseName = fileName
        extensionName = ".docx"
    End If
    ArchivedDocxPath = archiveFolderPath & "\" & baseName & " - archived " & Format(Now, "yyyymmdd-hhnnss") & extensionName
End Function

Private Sub ArchiveExistingDocx(ByVal filePath As String)
    Dim archivePath As String
    If Dir(filePath) = "" Then Exit Sub
    archivePath = ArchivedDocxPath(filePath)
    Name filePath As archivePath
End Sub

Public Function GetItemFolderDocxPathForInventoryRow(ByVal inventoryRow As Long) As String
    GetItemFolderDocxPathForInventoryRow = GetItemFolderDocxPath(inventoryRow)
End Function

Public Sub ArchiveExistingDetailsDocx(ByVal inventoryRow As Long)
    ArchiveExistingDocx GetDetailsDocxPath(inventoryRow)
    ArchiveExistingDocx GetItemFolderDocxPath(inventoryRow)
End Sub

Public Sub UpdateDocxItemNumber(ByVal filePath As String, ByVal oldItemNumber As String, ByVal newItemNumber As String)

    Dim wdApp As Object
    Dim wdDoc As Object
    Dim tbl As Object
    Dim rowIdx As Long
    
    On Error GoTo ErrorHandler
    
    If Dir(filePath) = "" Then Exit Sub
    
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    
    Set wdDoc = wdApp.Documents.Open(filePath)
    
    ' 1. Replace in the Subheader (Paragraph 2)
    If wdDoc.Paragraphs.Count >= 2 Then
        Dim subheaderRange As Object
        Set subheaderRange = wdDoc.Paragraphs(2).Range
        If InStr(subheaderRange.text, oldItemNumber) > 0 Then
            subheaderRange.text = Replace(subheaderRange.text, oldItemNumber, newItemNumber)
            ' Re-apply subheader formatting
            With subheaderRange
                .Font.Name = "Atkinson Hyperlegible"
                .Font.Size = 11
                .Font.Bold = False
                .Font.Color = RGB(68, 68, 68) ' #444444
            End With
        End If
    End If
    
    ' 2. Replace in Table 1, cell labeled "ITEM NUMBER"
    If wdDoc.Tables.Count > 0 Then
        Set tbl = wdDoc.Tables(1)
        For rowIdx = 1 To tbl.Rows.Count
            Dim labelText As String
            labelText = CellText(tbl, rowIdx, 2)
            If UCase$(labelText) = "ITEM NUMBER" Then
                SetCellText tbl, rowIdx, 3, newItemNumber
                Exit For
            End If
        Next rowIdx
    End If
    
    wdDoc.Save
    wdDoc.Close True
    wdApp.Quit
    
    Set wdDoc = Nothing
    Set wdApp = Nothing
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
    On Error GoTo 0

End Sub

Private Function CellText(ByVal tbl As Object, ByVal rowNumber As Long, ByVal columnNumber As Long) As String
    Dim valueText As String
    On Error GoTo ExitFunction
    valueText = tbl.cell(rowNumber, columnNumber).Range.text
    valueText = Replace(valueText, Chr$(13), "")
    valueText = Replace(valueText, Chr$(7), "")
    CellText = Trim$(valueText)
ExitFunction:
End Function

Private Sub SetCellText(ByVal tbl As Object, ByVal rowNumber As Long, ByVal columnNumber As Long, ByVal valueText As String)
    On Error GoTo ExitSub
    tbl.cell(rowNumber, columnNumber).Range.text = CStr(valueText)
ExitSub:
End Sub

Private Function TableCellExists(ByVal tbl As Object, ByVal rowNumber As Long, ByVal columnNumber As Long) As Boolean
    On Error GoTo ExitFunction
    Dim testText As String
    testText = tbl.cell(rowNumber, columnNumber).Range.text
    TableCellExists = True
ExitFunction:
End Function

Private Function NormalizeLabel(ByVal labelText As String) As String
    labelText = UCase$(Trim$(labelText))
    labelText = Replace(labelText, ChrW$(8211), "-")
    labelText = Replace(labelText, ChrW$(8212), "-")
    NormalizeLabel = labelText
End Function

Private Function GetTableRowByLabel(ByVal tbl As Object, ByVal labelText As String) As Long
    Dim i As Long
    For i = 1 To tbl.Rows.Count
        If TableCellExists(tbl, i, 2) And TableCellExists(tbl, i, 3) Then
            If NormalizeLabel(CellText(tbl, i, 2)) = NormalizeLabel(labelText) Then
                GetTableRowByLabel = i
                Exit Function
            End If
        End If
    Next i
End Function

Private Sub SetTableValueByLabel(ByVal tbl As Object, ByVal labelText As String, ByVal valueText As String)
    Dim rowNumber As Long
    rowNumber = GetTableRowByLabel(tbl, labelText)
    If rowNumber > 0 Then SetCellText tbl, rowNumber, 3, valueText
End Sub

Private Function ParseNumber(ByVal valueText As String) As Double
    valueText = Replace(CStr(valueText), """", "")
    valueText = Replace(valueText, "inches", "", , , vbTextCompare)
    valueText = Replace(valueText, "inch", "", , , vbTextCompare)
    valueText = Replace(valueText, "in", "", , , vbTextCompare)
    valueText = Trim$(valueText)
    If IsNumeric(valueText) Then ParseNumber = CDbl(valueText)
End Function

Private Function GetTableNumberByLabel(ByVal tbl As Object, ByVal labelText As String) As Double
    Dim rowNumber As Long
    rowNumber = GetTableRowByLabel(tbl, labelText)
    If rowNumber > 0 Then GetTableNumberByLabel = ParseNumber(CellText(tbl, rowNumber, 3))
End Function

Private Function SplitMarkdownRow(ByVal rowText As String) As Variant
    Dim parts As Variant
    rowText = Trim$(rowText)
    If Left$(rowText, 1) = "|" Then rowText = Mid$(rowText, 2)
    If Right$(rowText, 1) = "|" Then rowText = Left$(rowText, Len(rowText) - 1)
    parts = Split(rowText, "|")
    SplitMarkdownRow = parts
End Function

Private Function IsMarkdownSeparatorRow(ByVal rowText As String) As Boolean
    Dim strippedText As String
    strippedText = Replace(rowText, "|", "")
    strippedText = Replace(strippedText, "-", "")
    strippedText = Replace(strippedText, ":", "")
    strippedText = Trim$(strippedText)
    IsMarkdownSeparatorRow = (strippedText = "")
End Function

Private Function WordColorFromHex(ByVal hexColor As String) As Long
    Dim redValue As Long
    Dim greenValue As Long
    Dim blueValue As Long
    hexColor = Replace(hexColor, "#", "")
    redValue = CLng("&H" & Mid$(hexColor, 1, 2))
    greenValue = CLng("&H" & Mid$(hexColor, 3, 2))
    blueValue = CLng("&H" & Mid$(hexColor, 5, 2))
    WordColorFromHex = RGB(redValue, greenValue, blueValue)
End Function

Private Function PointsFromInches(ByVal inchesValue As Double) As Single
    PointsFromInches = CSng(inchesValue * 72)
End Function

Private Function IsUserEditableRow(ByVal rowNumber As Long) As Boolean
    IsUserEditableRow = (rowNumber = 12 Or rowNumber = 13 Or rowNumber = 15 Or rowNumber = 16 Or rowNumber = 17)
End Function

Private Function DetailsDisplayRowNumber(ByVal tbl As Object, ByVal rowNumber As Long) As Long
    Dim displayValue As String
    displayValue = Trim$(CellText(tbl, rowNumber, 1))
    If IsNumeric(displayValue) Then DetailsDisplayRowNumber = CLng(displayValue)
End Function

Private Function IsUserEditableDisplayRow(ByVal tbl As Object, ByVal rowNumber As Long) As Boolean
    IsUserEditableDisplayRow = IsUserEditableRow(DetailsDisplayRowNumber(tbl, rowNumber))
End Function

Private Function IsAmberDisplayRow(ByVal tbl As Object, ByVal rowNumber As Long) As Boolean
    Select Case DetailsDisplayRowNumber(tbl, rowNumber)
        Case 18 To 23, 26
            IsAmberDisplayRow = True
    End Select
End Function

Private Function ProperReferenceLabel(ByVal labelText As String) As String
    Dim words As Variant
    Dim i As Long
    Dim wordText As String
    labelText = LCase$(Trim$(labelText))
    If labelText = "" Then Exit Function
    words = Split(labelText, " ")
    For i = LBound(words) To UBound(words)
        wordText = CStr(words(i))
        Select Case wordText
            Case "usps"
                words(i) = "USPS"
            Case "in", "to", "of"
                words(i) = wordText
            Case Else
                If Len(wordText) > 0 Then words(i) = UCase$(Left$(wordText, 1)) & Mid$(wordText, 2)
        End Select
    Next i
    ProperReferenceLabel = Join(words, " ")
End Function

Private Sub NormalizeReferenceTableLabels(ByVal tbl As Object)
    Dim rowNumber As Long
    Dim labelText As String
    Dim dashText As String
    dashText = " " & ChrW$(8212) & " "
    For rowNumber = 1 To tbl.Rows.Count
        labelText = UCase$(CellText(tbl, rowNumber, 2))
        Select Case labelText
            Case "ITEM NAME"
                SetCellText tbl, rowNumber, 2, "Title"
            Case "SUB CATEGORY"
                SetCellText tbl, rowNumber, 2, "Sub Category"
            Case "SUB SUB CATEGORY"
                SetCellText tbl, rowNumber, 2, "Sub Sub Category"
            Case "SHIPPING - SHIPPING LABEL"
                SetCellText tbl, rowNumber, 2, "Shipping" & dashText & "Shipping Label"
            Case "SHIPPING - OFFER BUYERS FREE SHIPPING"
                SetCellText tbl, rowNumber, 2, "Shipping" & dashText & "Offer Buyers Free Shipping"
            Case "SHIPPING - PACKAGE WEIGHT - POUNDS"
                SetCellText tbl, rowNumber, 2, "Shipping" & dashText & "Package Weight (pounds)"
            Case "SHIPPING - PACKAGE WEIGHT - OUNCES"
                SetCellText tbl, rowNumber, 2, "Shipping" & dashText & "Package Weight (ounces)"
            Case "SHIPPING - WILL FIT IN SHOEBOX"
                SetCellText tbl, rowNumber, 2, "Shipping" & dashText & "Will Fit In Shoebox"
            Case "SHIPPING - PACKAGE - LENGTH"
                SetCellText tbl, rowNumber, 2, "Shipping" & dashText & "Package Length (inches)"
            Case "SHIPPING - PACKAGE - WIDTH"
                SetCellText tbl, rowNumber, 2, "Shipping" & dashText & "Package Width (inches)"
            Case "SHIPPING - PACKAGE - HEIGHT"
                SetCellText tbl, rowNumber, 2, "Shipping" & dashText & "Package Height (inches)"
            Case "USPS PRIORITY MAIL - BOX TYPE"
                SetCellText tbl, rowNumber, 2, "USPS Priority Mail" & dashText & "Box Type"
            Case "USPS PRIORITY MAIL PACKAGE - LENGTH"
                SetCellText tbl, rowNumber, 2, "USPS Priority Mail" & dashText & "Package Length (inches) *"
            Case "USPS PRIORITY MAIL PACKAGE - WIDTH"
                SetCellText tbl, rowNumber, 2, "USPS Priority Mail" & dashText & "Package Width (inches) *"
            Case "USPS PRIORITY MAIL PACKAGE - HEIGHT"
                SetCellText tbl, rowNumber, 2, "USPS Priority Mail" & dashText & "Package Height (inches) *"
            Case "USPS PRIORITY MAIL PACKAGE WEIGHT - POUNDS"
                SetCellText tbl, rowNumber, 2, "USPS Priority Mail" & dashText & "Package Weight (pounds) *"
            Case "USPS PRIORITY MAIL PACKAGE WEIGHT - OUNCES"
                SetCellText tbl, rowNumber, 2, "USPS Priority Mail" & dashText & "Package Weight (ounces) *"
            Case "SHIPPING - SELECT LABEL"
                SetCellText tbl, rowNumber, 2, "Shipping" & dashText & "Select Label"
            Case "PRICING"
                SetCellText tbl, rowNumber, 2, "Pricing"
            Case "PRICING RESEARCH"
                SetCellText tbl, rowNumber, 2, "Pricing Research *"
            Case "#", "FIELD"
                SetCellText tbl, rowNumber, 2, "Field"
            Case Else
                If labelText <> "" Then SetCellText tbl, rowNumber, 2, ProperReferenceLabel(CellText(tbl, rowNumber, 2))
        End Select
    Next rowNumber
End Sub

Private Function DetailsTableValue(ByVal tbl As Object, ByVal labelText As String) As String
    Dim rowNumber As Long
    rowNumber = GetTableRowByLabel(tbl, labelText)
    If rowNumber > 0 Then DetailsTableValue = CellText(tbl, rowNumber, 3)
End Function

Private Function DetailsReferenceTitle(ByVal tbl As Object) As String
    DetailsReferenceTitle = "Mercari Seller — Item Reference Sheet"
End Function

Private Function DetailsReferenceSubtitle(ByVal tbl As Object) As String
    Dim itemNumber As String
    Dim itemTitle As String
    itemNumber = CellText(tbl, 2, 3)
    itemTitle = CellText(tbl, 4, 3)
    If itemNumber = "" Then itemNumber = DetailsTableValue(tbl, "ITEM NUMBER")
    If itemTitle = "" Then itemTitle = DetailsTableValue(tbl, "TITLE")
    If itemNumber <> "" Or itemTitle <> "" Then
        DetailsReferenceSubtitle = "Item " & itemNumber
        If itemTitle <> "" Then DetailsReferenceSubtitle = DetailsReferenceSubtitle & " " & ChrW$(183) & " " & itemTitle
    Else
        DetailsReferenceSubtitle = "Item reference details"
    End If
End Function

Private Function DetailsReferenceSubtitleFromTable(ByVal tbl As Object) As String
    Dim itemNumber As String
    Dim itemTitle As String
    itemNumber = DetailsTableValue(tbl, "ITEM NUMBER")
    itemTitle = DetailsTableValue(tbl, "TITLE")
    If itemTitle = "" Then itemTitle = DetailsTableValue(tbl, "ITEM NAME")
    If itemNumber <> "" Or itemTitle <> "" Then
        DetailsReferenceSubtitleFromTable = "Item " & itemNumber
        If itemTitle <> "" Then DetailsReferenceSubtitleFromTable = DetailsReferenceSubtitleFromTable & " " & ChrW$(183) & " " & itemTitle
    Else
        DetailsReferenceSubtitleFromTable = "Item reference details"
    End If
End Function

Private Function IsLegacyUserEditableRow(ByVal rowNumber As Long) As Boolean
    Select Case rowNumber
        Case 12, 13, 15, 16, 17
            IsLegacyUserEditableRow = True
    End Select
End Function

Private Function IsAmberRow(ByVal tbl As Object, ByVal rowNumber As Long) As Boolean
    Dim labelText As String
    labelText = UCase$(CellText(tbl, rowNumber, 2))
    IsAmberRow = (IsAmberDisplayRow(tbl, rowNumber) Or _
                  InStr(labelText, "SUPPLEMENT") > 0 Or _
                  InStr(labelText, "NOTE") > 0 Or _
                  InStr(labelText, "REVIEW") > 0)
End Function

Private Sub ApplyCellMargins(ByVal tbl As Object)
    On Error Resume Next
    tbl.TopPadding = PointsFromInches(0.08)
    tbl.BottomPadding = PointsFromInches(0.08)
    tbl.LeftPadding = PointsFromInches(0.08)
    tbl.RightPadding = PointsFromInches(0.08)
    On Error GoTo 0
End Sub

Private Sub InsertDetailsDocumentHeading(ByVal wdDoc As Object, ByVal subtitleText As String)
    Dim titleText As String
    Dim ruleRange As Object

    titleText = "Mercari Seller " & ChrW$(8212) & " Item Reference Sheet"
    If Trim$(subtitleText) = "" Then subtitleText = "Item reference details"

    wdDoc.Range(0, 0).InsertBefore titleText & vbCrLf & subtitleText & vbCrLf & vbCrLf & vbCrLf

    With wdDoc.Paragraphs(1).Range
        .Font.Name = "Atkinson Hyperlegible"
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = WordColorFromHex("#5E6DF2")
        .ParagraphFormat.Alignment = 0
    End With

    With wdDoc.Paragraphs(2).Range
        .Font.Name = "Atkinson Hyperlegible"
        .Font.Size = 11
        .Font.Bold = False
        .Font.Color = WordColorFromHex("#444444")
        .ParagraphFormat.Alignment = 0
    End With

    Set ruleRange = wdDoc.Paragraphs(4).Range
    With ruleRange
        .Font.Name = "Atkinson Hyperlegible"
        .Font.Size = 1
        .Font.Color = WordColorFromHex("#5E6DF2")
        .ParagraphFormat.Alignment = 0
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.SpaceAfter = 8
        .Borders(-1).LineStyle = 1
        .Borders(-1).LineWidth = 4
        .Borders(-1).Color = WordColorFromHex("#5E6DF2")
    End With
End Sub

Private Sub RefreshDetailsDocumentHeading(ByVal wdDoc As Object, ByVal tbl As Object)
    On Error Resume Next
    wdDoc.Paragraphs(2).Range.text = DetailsReferenceSubtitleFromTable(tbl) & vbCr
    With wdDoc.Paragraphs(2).Range
        .Font.Name = "Atkinson Hyperlegible"
        .Font.Size = 11
        .Font.Bold = False
        .Font.Color = WordColorFromHex("#444444")
        .ParagraphFormat.Alignment = 0
    End With
    On Error GoTo 0
End Sub

Private Sub ApplyDetailsTableBorders(ByVal tbl As Object)
    Dim borderIndex As Variant
    Dim borderIndexes As Variant
    On Error Resume Next
    tbl.Borders.Enable = True
    borderIndexes = Array(-1, -2, -3, -4, -5, -6)
    For Each borderIndex In borderIndexes
        tbl.Borders(borderIndex).LineStyle = 1
        tbl.Borders(borderIndex).Color = WordColorFromHex("#CCCCCC")
        tbl.Borders(borderIndex).LineWidth = 2
    Next borderIndex
    On Error GoTo 0
End Sub

Private Sub ApplyDetailsRowFormatting(ByVal tbl As Object, ByVal rowNumber As Long)
    Dim backgroundColor As Long
    Dim leftStripeColor As Long
    Dim isAmber As Boolean

    isAmber = IsAmberRow(tbl, rowNumber)
    If isAmber Then
        If rowNumber Mod 2 = 0 Then
            backgroundColor = WordColorFromHex("#FFF8E0")
        Else
            backgroundColor = WordColorFromHex("#FFF3CD")
        End If
        leftStripeColor = WordColorFromHex("#FFB800")
    Else
        If rowNumber Mod 2 = 0 Then
            backgroundColor = WordColorFromHex("#ECEEFF")
        Else
            backgroundColor = WordColorFromHex("#FFFFFF")
        End If
        If IsUserEditableDisplayRow(tbl, rowNumber) Then leftStripeColor = WordColorFromHex("#3D5FC4")
    End If

    tbl.Rows(rowNumber).Shading.BackgroundPatternColor = backgroundColor
    tbl.Rows(rowNumber).Range.Font.Color = WordColorFromHex("#000000")

    If leftStripeColor <> 0 Then
        On Error Resume Next
        tbl.cell(rowNumber, 1).Borders(-2).LineStyle = 1
        tbl.cell(rowNumber, 1).Borders(-2).LineWidth = 12
        tbl.cell(rowNumber, 1).Borders(-2).Color = leftStripeColor
        On Error GoTo 0
    End If
End Sub

Private Sub AddDetailsLegend(ByVal wdDoc As Object)
    Dim legendRange As Object
    Dim wordRange As Object
    Dim legendText As String
    Dim legendStart As Long
    Dim blueStart As Long
    Dim amberStart As Long
    Set legendRange = wdDoc.content
    legendRange.Collapse 0
    legendText = "Border and Asterisk guide: Blue = Field value the user may wish to use alternate data from the Amber fields. Rows with an asterisk are supplemental / not part of the standard workflow. If shipping via USPS Priority Mail, use rows 18-23 instead of 12-17. To assist you in determining your final listing price, Row 26 may include pricing research details performed by AI."
    legendStart = legendRange.Start + 1
    legendRange.InsertAfter vbCrLf & legendText
    Set legendRange = wdDoc.Paragraphs(wdDoc.Paragraphs.Count).Range
    legendRange.Font.Name = "Atkinson Hyperlegible"
    legendRange.Font.Size = 10
    legendRange.Font.Color = WordColorFromHex("#666666")
    On Error Resume Next
    blueStart = InStr(1, legendText, "Blue", vbTextCompare)
    If blueStart > 0 Then
        Set wordRange = wdDoc.Range(legendStart + blueStart - 1, legendStart + blueStart + 3)
        wordRange.Font.Color = WordColorFromHex("#3D5FC4")
        wordRange.Font.Bold = True
    End If
    amberStart = InStr(1, legendText, "Amber", vbTextCompare)
    If amberStart > 0 Then
        Set wordRange = wdDoc.Range(legendStart + amberStart - 1, legendStart + amberStart + 4)
        wordRange.Font.Color = WordColorFromHex("#996600")
        wordRange.Font.Bold = True
    End If
    Set legendRange = wdDoc.Paragraphs(wdDoc.Paragraphs.Count).Range
    legendRange.ParagraphFormat.SpaceBefore = 8
    legendRange.ParagraphFormat.SpaceAfter = 0
    legendRange.Borders(-2).LineStyle = 1
    legendRange.Borders(-2).LineWidth = 4
    legendRange.Borders(-2).Color = WordColorFromHex("#3D5FC4")
    On Error GoTo 0
End Sub

Private Function IsUsableMarkdownTableRow(ByVal rowText As String) As Boolean
    Dim parts As Variant
    If InStr(rowText, "|") = 0 Then Exit Function
    If IsMarkdownSeparatorRow(rowText) Then Exit Function
    parts = SplitMarkdownRow(rowText)
    If UBound(parts) < 2 Then Exit Function
    If Trim$(CStr(parts(0))) = "" And Trim$(CStr(parts(1))) = "" And Trim$(CStr(parts(2))) = "" Then Exit Function
    IsUsableMarkdownTableRow = True
End Function

Private Function DetailsReferenceSubtitleFromResponse(ByVal responseText As String) As String
    Dim lines As Variant
    Dim i As Long
    Dim parts As Variant
    Dim labelText As String
    Dim itemNumber As String
    Dim itemTitle As String

    lines = Split(Replace(responseText, vbCrLf, vbLf), vbLf)
    For i = LBound(lines) To UBound(lines)
        If IsUsableMarkdownTableRow(CStr(lines(i))) Then
            parts = SplitMarkdownRow(CStr(lines(i)))
            labelText = NormalizeLabel(CStr(parts(1)))
            Select Case labelText
                Case "ITEM NUMBER"
                    itemNumber = Trim$(CStr(parts(2)))
                Case "ITEM NAME", "TITLE"
                    itemTitle = Trim$(CStr(parts(2)))
            End Select
        End If
    Next i

    If itemNumber <> "" Or itemTitle <> "" Then
        DetailsReferenceSubtitleFromResponse = "Item " & itemNumber
        If itemTitle <> "" Then DetailsReferenceSubtitleFromResponse = DetailsReferenceSubtitleFromResponse & " · " & itemTitle
    Else
        DetailsReferenceSubtitleFromResponse = "Item reference details"
    End If
End Function

Private Sub AddMarkdownTableToDocument(ByVal wdDoc As Object, ByVal responseText As String)
    Dim lines As Variant
    Dim i As Long
    Dim parts As Variant
    Dim rowCount As Long
    Dim currentRow As Long
    Dim tbl As Object
    Dim insertRange As Object
    lines = Split(Replace(responseText, vbCrLf, vbLf), vbLf)
    For i = LBound(lines) To UBound(lines)
        If IsUsableMarkdownTableRow(CStr(lines(i))) Then
            parts = SplitMarkdownRow(CStr(lines(i)))
            rowCount = rowCount + 1
        End If
    Next i
    If rowCount = 0 Then Err.Raise vbObjectError + 6303, "AddMarkdownTableToDocument", "No 3-column Markdown table was found on the clipboard. Copy the full AI response table, then try again."
    InsertDetailsDocumentHeading wdDoc, DetailsReferenceSubtitleFromResponse(responseText)
    Set insertRange = wdDoc.content
    insertRange.Collapse 0
    Set tbl = wdDoc.Tables.Add(insertRange, rowCount, 3)
    currentRow = 1
    For i = LBound(lines) To UBound(lines)
        If IsUsableMarkdownTableRow(CStr(lines(i))) Then
            parts = SplitMarkdownRow(CStr(lines(i)))
            SetCellText tbl, currentRow, 1, Trim$(CStr(parts(0)))
            SetCellText tbl, currentRow, 2, Trim$(CStr(parts(1)))
            SetCellText tbl, currentRow, 3, Trim$(CStr(parts(2)))
            currentRow = currentRow + 1
        End If
    Next i
End Sub

Private Sub FormatDetailsDocument(ByVal wdDoc As Object, ByVal tbl As Object)
    Dim i As Long
    On Error Resume Next
    With wdDoc.PageSetup
        .Orientation = 1
        .TopMargin = PointsFromInches(0.75)
        .BottomMargin = PointsFromInches(0.75)
        .LeftMargin = PointsFromInches(0.75)
        .RightMargin = PointsFromInches(0.75)
    End With
    NormalizeReferenceTableLabels tbl
    RefreshDetailsDocumentHeading wdDoc, tbl
    tbl.AllowAutoFit = False
    tbl.AutoFitBehavior 0
    tbl.Rows.LeftIndent = 0
    tbl.PreferredWidthType = 3
    tbl.PreferredWidth = PointsFromInches(9.5)
    tbl.Columns(1).Width = PointsFromInches(0.5)
    tbl.Columns(2).Width = PointsFromInches(3.75)
    tbl.Columns(3).Width = PointsFromInches(5.25)
    tbl.Rows.AllowBreakAcrossPages = False
    tbl.Rows(1).HeadingFormat = True
    ApplyCellMargins tbl
    ApplyDetailsTableBorders tbl
    tbl.Range.Font.Name = "Atkinson Hyperlegible"
    tbl.Range.Font.Size = 12
    tbl.Range.Font.Color = WordColorFromHex("#000000")
    tbl.Range.Font.Underline = False
    tbl.Range.ParagraphFormat.SpaceBefore = 0
    tbl.Range.ParagraphFormat.SpaceAfter = 0
    tbl.Range.ParagraphFormat.LineSpacingRule = 0
    SetCellText tbl, 1, 1, "#"
    SetCellText tbl, 1, 2, "Field"
    SetCellText tbl, 1, 3, "Value   (copy to paste)"
    tbl.Rows(1).Shading.BackgroundPatternColor = WordColorFromHex("#5E6DF2")
    tbl.Rows(1).Range.Font.Color = WordColorFromHex("#FFFFFF")
    tbl.Rows(1).Range.Font.Bold = True
    tbl.Rows(1).Range.Font.Size = 14
    tbl.Columns(1).Cells.VerticalAlignment = 1
    tbl.Columns(2).Cells.VerticalAlignment = 1
    tbl.Columns(3).Cells.VerticalAlignment = 1
    tbl.Columns(1).Cells.Range.ParagraphFormat.Alignment = 1
    tbl.Columns(2).Cells.Range.ParagraphFormat.Alignment = 0
    tbl.Columns(3).Cells.Range.ParagraphFormat.Alignment = 0
    tbl.Rows(1).Range.ParagraphFormat.Alignment = 1
    For i = 2 To tbl.Rows.Count
        ApplyDetailsRowFormatting tbl, i
        tbl.cell(i, 1).Range.ParagraphFormat.Alignment = 1
        tbl.cell(i, 2).Range.Font.Underline = False
        tbl.cell(i, 3).Range.Font.Underline = False
    Next i
    AddDetailsLegend wdDoc
    On Error GoTo 0
End Sub

Private Sub EnrichDetailsTable(ByVal tbl As Object, ByVal inventoryRow As Long)
    Dim ws As Worksheet
    Dim photoFolderPath As String
    Dim itemLength As Double
    Dim itemWidth As Double
    Dim itemHeight As Double
    Dim boxMatch As tUSPSBoxMatch
    Dim contentPounds As Double
    Dim contentOunces As Double
    Dim shippingPounds As Long
    Dim shippingOunces As Long
    Dim totalShippingOunces As Double

    Set ws = ThisWorkbook.Worksheets(WS_INVENTORY)
    photoFolderPath = CStr(ws.Cells(inventoryRow, COL_DETAILS_FOLDER).Value)
    If Trim$(photoFolderPath) = "" Then photoFolderPath = GetItemPhotoFolderPath(inventoryRow)

    SetTableValueByLabel tbl, "ITEM NUMBER", CStr(ws.Cells(inventoryRow, COL_ITEM_NUMBER).Value)
    SetTableValueByLabel tbl, "PHOTOS", photoFolderPath
    SetTableValueByLabel tbl, "PRICING", CStr(ws.Cells(inventoryRow, COL_ITEM_PRICE).Value)

    itemLength = GetTableNumberByLabel(tbl, "SHIPPING - PACKAGE - LENGTH")
    itemWidth = GetTableNumberByLabel(tbl, "SHIPPING - PACKAGE - WIDTH")
    itemHeight = GetTableNumberByLabel(tbl, "SHIPPING - PACKAGE - HEIGHT")

    If itemLength > 0 And itemWidth > 0 And itemHeight > 0 Then
        On Error GoTo ShippingLookupError
        boxMatch = FindSmallestUSPSPriorityMailBox(itemLength, itemWidth, itemHeight)
        On Error GoTo 0
        If boxMatch.Found Then
            SetTableValueByLabel tbl, "USPS PRIORITY MAIL - BOX TYPE", boxMatch.BoxName
            SetTableValueByLabel tbl, "USPS PRIORITY MAIL PACKAGE - LENGTH", CStr(boxMatch.OuterLength)
            SetTableValueByLabel tbl, "USPS PRIORITY MAIL PACKAGE - WIDTH", CStr(boxMatch.OuterWidth)
            SetTableValueByLabel tbl, "USPS PRIORITY MAIL PACKAGE - HEIGHT", CStr(boxMatch.OuterHeight)
            contentPounds = CDbl(Val(GetFieldValue(inventoryRow, "WEIGHT_POUNDS")))
            contentOunces = CDbl(Val(GetFieldValue(inventoryRow, "WEIGHT_OUNCES")))
            totalShippingOunces = (contentPounds * 16) + contentOunces + boxMatch.BoxWeightOunces + boxMatch.PackingWeightOunces
            ConvertTotalOuncesToPoundsOunces totalShippingOunces, shippingPounds, shippingOunces
            SetTableValueByLabel tbl, "USPS PRIORITY MAIL PACKAGE WEIGHT - POUNDS", CStr(shippingPounds)
            SetTableValueByLabel tbl, "USPS PRIORITY MAIL PACKAGE WEIGHT - OUNCES", CStr(shippingOunces)
        End If
    End If
    Exit Sub

ShippingLookupError:
    SetTableValueByLabel tbl, "USPS PRIORITY MAIL - BOX TYPE", "USPS box lookup needs review: " & Err.Description
End Sub

Public Function SaveAIResponseDetailsDocx(ByVal inventoryRow As Long) As String
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim responseText As String
    Dim filePath As String
    Dim itemFolderFilePath As String
    Dim currentStep As String
    On Error GoTo ErrorHandler

    currentStep = "Read clipboard"
    responseText = GetClipboardText()
    If Trim$(responseText) = "" Then Err.Raise vbObjectError + 6302, "SaveAIResponseDetailsDocx", "The clipboard is empty. Copy the AI response table, then try again."
    currentStep = "Build DOCX path"
    filePath = GetDetailsDocxPath(inventoryRow)
    itemFolderFilePath = GetItemFolderDocxPath(inventoryRow)
    ValidateDocxTargetPath filePath
    ValidateDocxTargetPath itemFolderFilePath
    currentStep = "Start Word"
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    currentStep = "Create Word document"
    Set wdDoc = wdApp.Documents.Add
    currentStep = "Create table from clipboard"
    AddMarkdownTableToDocument wdDoc, responseText
    currentStep = "Enrich details table"
    If wdDoc.Tables.Count > 0 Then EnrichDetailsTable wdDoc.Tables(1), inventoryRow
    currentStep = "Format details document"
    If wdDoc.Tables.Count > 0 Then FormatDetailsDocument wdDoc, wdDoc.Tables(1)
    currentStep = "Save DOCX"
    wdDoc.SaveAs2 filePath, 16
    If Dir(filePath) = "" Then Err.Raise vbObjectError + 6312, "SaveAIResponseDetailsDocx", "The DOCX file was not saved to the DOCX folder:" & vbCrLf & vbCrLf & filePath
    currentStep = "Save DOCX to item folder"
    wdDoc.SaveAs2 itemFolderFilePath, 16
    If Dir(itemFolderFilePath) = "" Then Err.Raise vbObjectError + 6313, "SaveAIResponseDetailsDocx", "The DOCX file was not saved to the item folder:" & vbCrLf & vbCrLf & itemFolderFilePath
    wdDoc.Saved = True
    wdDoc.Close SaveChanges:=False
    Set wdDoc = Nothing
    wdApp.Quit
    Set wdApp = Nothing
    DoEvents
    SaveAIResponseDetailsDocx = filePath
    Exit Function

ErrorHandler:
    Dim originalNumber As Long
    Dim originalDescription As String
    originalNumber = Err.Number
    originalDescription = Err.Description
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
    Err.Raise originalNumber, "SaveAIResponseDetailsDocx - " & currentStep, originalDescription
End Function
