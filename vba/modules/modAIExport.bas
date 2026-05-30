Attribute VB_Name = "modAIExport"
Option Explicit

Private Const AI_TEMPLATE_RELATIVE_PATH As String = "Documentation\MERCARI AI PROMPT - FINAL.docx"
Private Const AI_TEMPLATE_TEXT_RELATIVE_PATH As String = "Documentation\MERCARI AI PROMPT - FINAL.txt"
Private Const AI_TEMPLATE_WORKSHEET_NAME As String = "AI_TEMPLATE"

Private Function CleanPromptValue(ByVal valueText As String) As String
    CleanPromptValue = Trim$(CStr(valueText))
End Function

Private Function CleanUTF8ForClipboard(ByVal text As String) As String
    ' Remove UTF-8 BOM
    text = Replace(text, "ï»¿", "")
    ' Replace garbled UTF-8 sequences with correct characters
    text = Replace(text, "â€”", "—")
    text = Replace(text, "â€“", "–")
    text = Replace(text, "â†’", "→")
    text = Replace(text, "âœ“", "✓")
    text = Replace(text, "â‰¤", "≤")
    text = Replace(text, "âœ—", "✗")
    text = Replace(text, "â‰¥", "≥")
    CleanUTF8ForClipboard = text
End Function

Private Function ShouldIncludePromptValue(ByVal valueText As String) As Boolean
    valueText = UCase$(CleanPromptValue(valueText))
    If valueText = "" Then Exit Function
    If valueText = "N/A" Then Exit Function
    ShouldIncludePromptValue = True
End Function

Private Function FormatBooleanPromptValue(ByVal valueText As String) As String
    Dim parts() As String
    valueText = CleanPromptValue(valueText)
    If InStr(valueText, "|") = 0 Then
        FormatBooleanPromptValue = valueText
    Else
        parts = Split(valueText, "|")
        FormatBooleanPromptValue = Trim$(parts(0))
        If UBound(parts) >= 1 Then
            If Trim$(parts(1)) <> "" Then
                FormatBooleanPromptValue = FormatBooleanPromptValue & " - " & Trim$(parts(1))
            End If
        End If
    End If
End Function

Private Function AppendPromptLine(ByVal promptText As String, ByVal labelText As String, ByVal valueText As String) As String
    valueText = FormatBooleanPromptValue(valueText)
    If ShouldIncludePromptValue(valueText) Then
        AppendPromptLine = promptText & "- " & labelText & ": " & valueText & vbCrLf
    Else
        AppendPromptLine = promptText
    End If
End Function

Private Function GetTemplatePath() As String
    Dim workbookFolderPath As String
    Dim projectSourceFolderPath As String
    Dim parentFolderPath As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    workbookFolderPath = fso.GetParentFolderName(ThisWorkbook.FullName)
    If InStr(1, workbookFolderPath, "\PROJECT - Mercari Workbook", vbTextCompare) > 0 Then
        GetTemplatePath = fso.BuildPath(workbookFolderPath, AI_TEMPLATE_RELATIVE_PATH)
        Exit Function
    End If
    parentFolderPath = fso.GetParentFolderName(workbookFolderPath)
    If InStr(1, parentFolderPath, "\PROJECT - Mercari Workbook", vbTextCompare) > 0 Then
        projectSourceFolderPath = parentFolderPath
    Else
        Dim grandparentFolderPath As String
        grandparentFolderPath = fso.GetParentFolderName(parentFolderPath)
        If InStr(1, grandparentFolderPath, "\PROJECT - Mercari Workbook", vbTextCompare) > 0 Then
            projectSourceFolderPath = grandparentFolderPath
        Else
            projectSourceFolderPath = workbookFolderPath
        End If
    End If
    GetTemplatePath = fso.BuildPath(projectSourceFolderPath, AI_TEMPLATE_RELATIVE_PATH)
End Function

Private Function ReadDocxText(ByVal filePath As String) As String
    Dim wdApp As Object
    Dim wdDoc As Object
    On Error GoTo ErrorHandler
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set wdDoc = wdApp.Documents.Open(filePath, False, True)
    ReadDocxText = wdDoc.content.text
    wdDoc.Close SaveChanges:=False
    wdApp.Quit
    Set wdDoc = Nothing
    Set wdApp = Nothing
    Exit Function
ErrorHandler:
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=False
    If Not wdApp Is Nothing Then wdApp.Quit
    Err.Raise vbObjectError + 6201, "ReadDocxText", "Unable to read AI prompt template: " & filePath
End Function

Private Function GetTextTemplatePath() As String
    Dim workbookFolderPath As String
    Dim projectSourceFolderPath As String
    Dim parentFolderPath As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    workbookFolderPath = fso.GetParentFolderName(ThisWorkbook.FullName)
    If InStr(1, workbookFolderPath, "\PROJECT - Mercari Workbook", vbTextCompare) > 0 Then
        GetTextTemplatePath = fso.BuildPath(workbookFolderPath, AI_TEMPLATE_TEXT_RELATIVE_PATH)
        Exit Function
    End If
    parentFolderPath = fso.GetParentFolderName(workbookFolderPath)
    If InStr(1, parentFolderPath, "\PROJECT - Mercari Workbook", vbTextCompare) > 0 Then
        projectSourceFolderPath = parentFolderPath
    Else
        Dim grandparentFolderPath As String
        grandparentFolderPath = fso.GetParentFolderName(parentFolderPath)
        If InStr(1, grandparentFolderPath, "\PROJECT - Mercari Workbook", vbTextCompare) > 0 Then
            projectSourceFolderPath = grandparentFolderPath
        Else
            projectSourceFolderPath = workbookFolderPath
        End If
    End If
    GetTextTemplatePath = fso.BuildPath(projectSourceFolderPath, AI_TEMPLATE_TEXT_RELATIVE_PATH)
End Function

Private Function ReadTextFile(ByVal filePath As String) As String
    Dim stream As Object
    On Error GoTo ErrorHandler
    Set stream = CreateObject("ADODB.Stream")
    stream.Charset = "UTF-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadTextFile = stream.ReadText
    stream.Close
    Exit Function
ErrorHandler:
    ReadTextFile = ""
End Function

Private Function ReadWorksheetTemplate() As String
    Dim ws As Worksheet
    Dim usedRange As Range
    Dim rowRange As Range
    Dim cellValue As String
    Dim firstSheet As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(AI_TEMPLATE_WORKSHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        ReadWorksheetTemplate = ""
        Exit Function
    End If
    If ws.Cells(1, 1).Value = "" Then
        ReadWorksheetTemplate = ""
        Exit Function
    End If
    Set usedRange = ws.usedRange
    For Each rowRange In usedRange.Rows
        cellValue = CStr(ws.Cells(rowRange.Row, 1).Value)
        If Trim$(cellValue) <> "" Then
            ReadWorksheetTemplate = ReadWorksheetTemplate & cellValue & vbCrLf
        End If
    Next rowRange
End Function

Private Function GetTemplateText() As String
    Dim templateText As String
    templateText = ReadWorksheetTemplate()
    If Trim$(templateText) <> "" Then
        GetTemplateText = templateText
        Exit Function
    End If
    templateText = ReadTextFile(GetTextTemplatePath())
    If Trim$(templateText) <> "" Then
        GetTemplateText = templateText
        Exit Function
    End If
    GetTemplateText = ReadDocxText(GetTemplatePath())
End Function

Public Sub ImportMercariAITemplateToWorksheet()
    Dim ws As Worksheet
    Dim templateText As String
    Dim lines As Variant
    Dim i As Long
    templateText = GetTemplateText()
    If Trim$(templateText) = "" Then
        MsgBox "AI template is empty or could not be loaded.", vbExclamation
        Exit Sub
    End If
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(AI_TEMPLATE_WORKSHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = AI_TEMPLATE_WORKSHEET_NAME
    End If
    templateText = Replace(templateText, vbCrLf, vbLf)
    templateText = Replace(templateText, vbCr, vbLf)
    lines = Split(templateText, vbLf)
    ws.Cells.ClearContents
    For i = LBound(lines) To UBound(lines)
        ws.Cells(i + 1, 1).Value = lines(i)
    Next i
    MsgBox "AI template imported into the workbook worksheet: " & AI_TEMPLATE_WORKSHEET_NAME, vbInformation
End Sub

Private Function BuildFieldSection(ByVal dataRow As Long, ByVal tabName As String) As String
    Dim i As Long
    Dim fieldValue As String
    EnsureFieldDefinitionsLoaded
    For i = LBound(FieldDefinitions) To UBound(FieldDefinitions)
        If UCase$(FieldDefinitions(i).tabName) = UCase$(tabName) Then
            fieldValue = CStr(Cells(dataRow, FieldDefinitions(i).dataColumn).Value)
            BuildFieldSection = AppendPromptLine(BuildFieldSection, FieldDefinitions(i).fieldLabel, fieldValue)
        End If
    Next i
End Function

Private Function BuildCombinedSpecsSection(ByVal dataRow As Long) As String
    BuildCombinedSpecsSection = AppendPromptLine(BuildCombinedSpecsSection, "Item Dimensions (L x H x W in inches)", BuildDimensionValue(dataRow, "ITEM"))
    BuildCombinedSpecsSection = AppendPromptLine(BuildCombinedSpecsSection, "Total Content Dimensions (L x H x W in inches)", BuildDimensionValue(dataRow, "TOTAL_CONTENT"))
    BuildCombinedSpecsSection = AppendPromptLine(BuildCombinedSpecsSection, "Total Content Weight (Pounds and Ounces)", BuildWeightValue(dataRow))
    BuildCombinedSpecsSection = AppendPromptLine(BuildCombinedSpecsSection, "Size", GetFieldValue(dataRow, "SIZE"))
    BuildCombinedSpecsSection = AppendPromptLine(BuildCombinedSpecsSection, "Size (Additional Info)", GetFieldValue(dataRow, "SIZE_ADDITIONAL"))
    BuildCombinedSpecsSection = AppendPromptLine(BuildCombinedSpecsSection, "Shoe Size", GetFieldValue(dataRow, "SHOE_SIZE"))
    BuildCombinedSpecsSection = AppendPromptLine(BuildCombinedSpecsSection, "Shoe Width", GetFieldValue(dataRow, "SHOE_WIDTH"))
End Function

Private Function BuildDimensionValue(ByVal dataRow As Long, ByVal dimensionPrefix As String) As String
    Dim lengthValue As String
    Dim heightValue As String
    Dim widthValue As String
    lengthValue = GetFieldValue(dataRow, dimensionPrefix & "_LENGTH")
    heightValue = GetFieldValue(dataRow, dimensionPrefix & "_HEIGHT")
    widthValue = GetFieldValue(dataRow, dimensionPrefix & "_WIDTH")
    If lengthValue <> "" Or heightValue <> "" Or widthValue <> "" Then
        BuildDimensionValue = lengthValue & " x " & heightValue & " x " & widthValue
    End If
End Function

Private Function BuildWeightValue(ByVal dataRow As Long) As String
    Dim poundsValue As String
    Dim ouncesValue As String
    poundsValue = GetFieldValue(dataRow, "WEIGHT_POUNDS")
    ouncesValue = GetFieldValue(dataRow, "WEIGHT_OUNCES")
    If poundsValue <> "" Or ouncesValue <> "" Then
        BuildWeightValue = poundsValue & " lbs " & ouncesValue & " oz"
    End If
End Function

Private Function BuildPhotosSection(ByVal dataRow As Long) As String
    Dim i As Long
    Dim photoPath As String
    For i = 0 To DEFAULT_MAX_PHOTOS - 1
        photoPath = GetFieldValue(dataRow, "PHOTO_" & (i + 1))
        If photoPath <> "" Then
            BuildPhotosSection = BuildPhotosSection & "Photo " & (i + 1) & ": " & photoPath & vbCrLf
        End If
    Next i
End Function

Private Function BuildItemDetailsSection(ByVal inventoryRow As Long) As String
    Dim dataRow As Long
    dataRow = inventoryRow
    BuildItemDetailsSection = "## ITEM DETAILS" & vbCrLf
    BuildItemDetailsSection = BuildItemDetailsSection & "### BASIC INFORMATION" & vbCrLf & BuildFieldSection(dataRow, "BASIC INFORMATION")
    BuildItemDetailsSection = BuildItemDetailsSection & "### PRICING AND SHIPPING" & vbCrLf & BuildFieldSection(dataRow, "PRICING AND SHIPPING")
    BuildItemDetailsSection = BuildItemDetailsSection & "### PHYSICAL SPECIFICATIONS" & vbCrLf & BuildCombinedSpecsSection(dataRow)
    BuildItemDetailsSection = BuildItemDetailsSection & "### CONDITION" & vbCrLf & BuildFieldSection(dataRow, "CONDITION")
    BuildItemDetailsSection = BuildItemDetailsSection & "### PHOTOS" & vbCrLf & BuildPhotosSection(dataRow)
    BuildItemDetailsSection = BuildItemDetailsSection & "### HISTORY" & vbCrLf & BuildFieldSection(dataRow, "HISTORY")
End Function

Private Function ReplaceItemDetailsSection(ByVal templateText As String, ByVal itemDetailsText As String) As String
    Dim startPos As Long
    Dim endPos As Long
    startPos = InStr(1, templateText, "## ITEM DETAILS", vbTextCompare)
    If startPos = 0 Then
        ReplaceItemDetailsSection = templateText & vbCrLf & vbCrLf & itemDetailsText
    Else
        ' Search for next ## AFTER the ## ITEM DETAILS (add 1 to skip past the current match)
        endPos = InStr(startPos + 1, templateText, "##", vbTextCompare)
        If endPos = 0 Then
            ReplaceItemDetailsSection = Left$(templateText, startPos - 1) & itemDetailsText
        Else
            ReplaceItemDetailsSection = Left$(templateText, startPos - 1) & itemDetailsText & vbCrLf & Mid$(templateText, endPos)
        End If
    End If
End Function

Public Function BuildMercariAIPrompt(ByVal inventoryRow As Long) As String
    Dim templateText As String
    Dim itemDetailsText As String
    templateText = GetTemplateText()
    itemDetailsText = BuildItemDetailsSection(inventoryRow)
    BuildMercariAIPrompt = CleanUTF8ForClipboard(ReplaceItemDetailsSection(templateText, itemDetailsText))
End Function
