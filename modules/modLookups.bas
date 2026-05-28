Attribute VB_Name = "modLookups"
Option Explicit

Private Const LOOKUP_VALUE_COLUMN As String = "Value"
Private mLookupInitializing As Boolean

Private Function GetLookupWorksheet() As Worksheet
    On Error Resume Next
    Set GetLookupWorksheet = ThisWorkbook.Worksheets(WS_LOOKUPS)
    On Error GoTo 0
    If GetLookupWorksheet Is Nothing Then
        Set GetLookupWorksheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetLookupWorksheet.Name = WS_LOOKUPS
    End If
End Function

Private Function ResolveLookupTableName(ByVal lookupKey As String) As String
    Dim key As String
    key = UCase$(Trim$(lookupKey))
    key = Replace(key, "TABLE_", "")
    Select Case key
        Case "BRAND", "BRANDS"
            ResolveLookupTableName = "tblBrands"
        Case "COLOR", "COLORS"
            ResolveLookupTableName = "tblColors"
        Case "MATERIAL", "MATERIALS"
            ResolveLookupTableName = "tblMaterials"
        Case "STYLE", "STYLES"
            ResolveLookupTableName = "tblStyles"
        Case "PATTERN", "PATTERNS"
            ResolveLookupTableName = "tblPatterns"
        Case "CONDITION", "CONDITIONS", "OVERALL_CONDITION"
            ResolveLookupTableName = "tblConditions"
        Case "FUNCTIONAL_STATUS"
            ResolveLookupTableName = "tblFunctionalStatus"
        Case "ODOR", "ODORS"
            ResolveLookupTableName = "tblOdors"
        Case "SHIPPING_RESTRICTION", "SHIPPING_RESTRICTIONS"
            ResolveLookupTableName = "tblShippingRestrictions"
        Case "STORAGE_ENVIRONMENT", "STORAGE_ENVIRONMENTS"
            ResolveLookupTableName = "tblStorageEnvironments"
    End Select
End Function

Private Function GetLookupTable(ByVal lookupKey As String) As ListObject
    Dim ws As Worksheet
    Dim tableName As String
    Set ws = GetLookupWorksheet()
    tableName = ResolveLookupTableName(lookupKey)
    If tableName = "" Then Exit Function
    On Error Resume Next
    Set GetLookupTable = ws.ListObjects(tableName)
    On Error GoTo 0
End Function

Private Function StarterValues(ByVal lookupKey As String) As Variant
    Select Case ResolveLookupTableName(lookupKey)
        Case "tblBrands"
            StarterValues = Array("Apple", "Samsung", "Sony", "Nintendo", "Microsoft", "Canon", "Nikon", "HP", "Dell", "Lenovo", "LG", "Panasonic", "Bose", "JBL", "Dyson", "KitchenAid", "Black & Decker", "Craftsman", "Makita", "DeWalt", "Nike", "Adidas", "Levi's", "Coach", "Michael Kors", "Disney", "Funko", "Mattel", "Hasbro", "Fisher-Price")
        Case "tblColors"
            StarterValues = Array("Black", "White", "Gray", "Silver", "Gold", "Beige", "Brown", "Tan", "Cream", "Red", "Maroon", "Pink", "Orange", "Yellow", "Green", "Olive", "Mint", "Teal", "Blue", "Navy", "Purple", "Lavender", "Clear", "Multicolor", "Rainbow", "Transparent", "Wood Tone", "Bronze", "Copper", "Chrome")
        Case "tblMaterials"
            StarterValues = Array("Plastic", "Metal", "Wood", "Glass", "Ceramic", "Porcelain", "Leather", "Faux Leather", "Cotton", "Polyester", "Wool", "Silk", "Linen", "Denim", "Canvas", "Rubber", "Foam", "Paper", "Cardboard", "Aluminum", "Steel", "Stainless Steel", "Cast Iron", "Brass", "Copper", "Bamboo", "Stone", "Marble", "Resin", "Vinyl", "Acrylic", "Carbon Fiber", "Composite", "Mixed Materials", "Unknown")
        Case "tblStyles"
            StarterValues = Array("Modern", "Contemporary", "Traditional", "Vintage", "Retro", "Mid-Century Modern", "Industrial", "Farmhouse", "Rustic", "Minimalist", "Bohemian", "Art Deco", "Victorian", "Classic", "Casual", "Formal", "Sport", "Athletic", "Streetwear", "Designer", "Luxury", "Gothic", "Steampunk", "Military", "Nautical", "Western", "Cottagecore", "Y2K", "Kawaii", "Preppy", "Eclectic", "Novelty", "Collectible")
        Case "tblPatterns"
            StarterValues = Array("Solid", "Striped", "Plaid", "Checkered", "Polka Dot", "Floral", "Paisley", "Camouflage", "Animal Print", "Leopard Print", "Zebra Print", "Snake Print", "Geometric", "Abstract", "Tie-Dye", "Graphic", "Logo", "Textured", "Herringbone", "Chevron", "Argyle", "Patchwork", "Toile", "Damask", "Fair Isle", "Ombre", "Galaxy", "Character Print", "Holiday", "Seasonal", "Novelty", "Multicolor")
        Case "tblConditions"
            StarterValues = Array("NEW - Brand new, never used", "LIKE NEW - Minimal to no signs of use", "GOOD - Light wear from normal use", "FAIR - Visible wear but functional", "POOR - Heavy wear or damage")
        Case "tblFunctionalStatus"
            StarterValues = Array("Fully Functional", "Partially Functional", "Powers On Only", "For Parts", "Non-Functional", "Restored", "Modified", "Incomplete", "Unknown")
        Case "tblOdors"
            StarterValues = Array("No Noticeable Odor", "Smoke Odor", "Pet Odor", "Musty Odor", "Mildew Odor", "Perfume Odor", "Chemical Odor", "Storage Odor", "Basement Odor", "Garage Odor", "Strong Odor Present", "Unknown")
        Case "tblShippingRestrictions"
            StarterValues = Array("None", "Ground Shipping Only", "Contains Lithium Battery", "Oversized", "Heavy Item", "Hazardous Material", "Local Pickup Recommended", "Adult Signature Recommended", "Cannot Ship Internationally", "Temperature Sensitive", "Glass Item", "Sharp Components", "Multiple Packages Required")
        Case "tblStorageEnvironments"
            StarterValues = Array("Indoor Storage", "Garage Storage", "Basement Storage", "Attic Storage", "Storage Unit", "Warehouse Storage", "Retail Shelf Storage", "Display Case", "Humid Environment", "Unknown Storage Conditions")
        Case Else
            StarterValues = Array()
    End Select
End Function

Private Sub EnsureLookupTable(ByVal lookupKey As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim tableName As String
    Dim nextColumn As Long
    Dim headerCell As Range
    Dim values As Variant
    Dim i As Long
    Set ws = GetLookupWorksheet()
    tableName = ResolveLookupTableName(lookupKey)
    If tableName = "" Then Exit Sub
    Set tbl = GetLookupTable(lookupKey)
    If tbl Is Nothing Then
        nextColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If nextColumn = 1 And Trim$(CStr(ws.Cells(1, 1).Value)) = "" Then
            nextColumn = 1
        Else
            nextColumn = nextColumn + 2
        End If
        Set headerCell = ws.Cells(1, nextColumn)
        headerCell.Value = LOOKUP_VALUE_COLUMN
        Set tbl = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=headerCell.Resize(2, 1), XlListObjectHasHeaders:=xlYes)
        tbl.Name = tableName
        If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.ClearContents
    End If
    values = StarterValues(lookupKey)
    If tableName = "tblConditions" Or tableName = "tblFunctionalStatus" Then
        ResetLookupTable tbl, values
        Exit Sub
    End If
    mLookupInitializing = True
    For i = LBound(values) To UBound(values)
        AddLookupValue lookupKey, CStr(values(i))
    Next i
    mLookupInitializing = False
End Sub

Private Sub ResetLookupTable(ByVal tbl As ListObject, ByVal values As Variant)
    Dim i As Long
    Dim newRow As ListRow
    If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Delete
    For i = LBound(values) To UBound(values)
        Set newRow = tbl.ListRows.Add
        newRow.Range.Cells(1, tbl.ListColumns(LOOKUP_VALUE_COLUMN).Index).NumberFormat = "@"
        newRow.Range.Cells(1, tbl.ListColumns(LOOKUP_VALUE_COLUMN).Index).Value = CStr(values(i))
    Next i
End Sub

Public Sub InitializeLookupArchitecture()
    EnsureLookupTable "BRAND"
    EnsureLookupTable "COLOR"
    EnsureLookupTable "MATERIAL"
    EnsureLookupTable "STYLE"
    EnsureLookupTable "PATTERN"
    EnsureLookupTable "CONDITION"
    EnsureLookupTable "FUNCTIONAL_STATUS"
    EnsureLookupTable "ODORS"
    EnsureLookupTable "SHIPPING_RESTRICTIONS"
    EnsureLookupTable "STORAGE_ENVIRONMENTS"
    GetLookupWorksheet.Columns.AutoFit
End Sub

Private Function NormalizeLookupValue(ByVal lookupValue As String) As String
    NormalizeLookupValue = UCase$(Application.WorksheetFunction.Trim(CStr(lookupValue)))
End Function

Public Function LookupValueExists(ByVal lookupKey As String, ByVal lookupValue As String) As Boolean
    Dim tbl As ListObject
    Dim cell As Range
    Dim normalizedValue As String
    Set tbl = GetLookupTable(lookupKey)
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    normalizedValue = NormalizeLookupValue(lookupValue)
    For Each cell In tbl.ListColumns(LOOKUP_VALUE_COLUMN).DataBodyRange.Cells
        If NormalizeLookupValue(cell.Value) = normalizedValue Then
            LookupValueExists = True
            Exit Function
        End If
    Next cell
End Function

Public Function IsLookupAutoLearnEnabled(ByVal fieldName As String) As Boolean
    Select Case UCase$(Trim$(fieldName))
        Case "BRAND", "COLOR", "MATERIAL", "STYLE", "PATTERN", "FUNCTIONAL_STATUS", "ODORS", "SHIPPING_RESTRICTIONS", "STORAGE_ENVIRONMENT"
            IsLookupAutoLearnEnabled = True
    End Select
End Function

Public Function IsRestrictedLookupValueValid(ByVal fieldName As String, ByVal fieldValue As String) As Boolean
    fieldValue = Application.WorksheetFunction.Trim(CStr(fieldValue))
    If fieldValue = "" Then
        IsRestrictedLookupValueValid = True
        Exit Function
    End If
    If IsPlaceholderValue(fieldName, fieldValue) Then
        IsRestrictedLookupValueValid = True
        Exit Function
    End If
    If IsLookupAutoLearnEnabled(fieldName) Then
        IsRestrictedLookupValueValid = True
        Exit Function
    End If
    IsRestrictedLookupValueValid = LookupValueExists(fieldName, fieldValue)
End Function

Public Function IsRestrictedLookupPrefixValid(ByVal fieldName As String, ByVal fieldValue As String) As Boolean
    Dim tbl As ListObject
    Dim cell As Range
    Dim normalizedValue As String
    fieldValue = Application.WorksheetFunction.Trim(CStr(fieldValue))
    If fieldValue = "" Then
        IsRestrictedLookupPrefixValid = True
        Exit Function
    End If
    If IsPlaceholderValue(fieldName, fieldValue) Then
        IsRestrictedLookupPrefixValid = True
        Exit Function
    End If
    If IsLookupAutoLearnEnabled(fieldName) Then
        IsRestrictedLookupPrefixValid = True
        Exit Function
    End If
    Set tbl = GetLookupTable(fieldName)
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    normalizedValue = NormalizeLookupValue(fieldValue)
    For Each cell In tbl.ListColumns(LOOKUP_VALUE_COLUMN).DataBodyRange.Cells
        If Left$(NormalizeLookupValue(cell.Value), Len(normalizedValue)) = normalizedValue Then
            IsRestrictedLookupPrefixValid = True
            Exit Function
        End If
    Next cell
End Function

Public Function AddLookupValue(ByVal lookupKey As String, ByVal lookupValue As String) As Boolean
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim cleanValue As String
    cleanValue = Application.WorksheetFunction.Trim(CStr(lookupValue))
    If cleanValue = "" Then Exit Function
    If mLookupInitializing = False Then EnsureLookupTable lookupKey
    Set tbl = GetLookupTable(lookupKey)
    If tbl Is Nothing Then Exit Function
    If LookupValueExists(lookupKey, cleanValue) Then Exit Function
    Set newRow = tbl.ListRows.Add
    newRow.Range.Cells(1, tbl.ListColumns(LOOKUP_VALUE_COLUMN).Index).NumberFormat = "@"
    newRow.Range.Cells(1, tbl.ListColumns(LOOKUP_VALUE_COLUMN).Index).Value = cleanValue
    If ResolveLookupTableName(lookupKey) <> "tblConditions" Then SortLookupTable lookupKey
    AddLookupValue = True
End Function

Public Sub SortLookupTable(ByVal lookupKey As String)
    Dim tbl As ListObject
    If ResolveLookupTableName(lookupKey) = "tblConditions" Then Exit Sub
    Set tbl = GetLookupTable(lookupKey)
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    With tbl.Sort
        .SortFields.Clear
        .SortFields.Add key:=tbl.ListColumns(LOOKUP_VALUE_COLUMN).DataBodyRange, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .Apply
    End With
End Sub

Public Sub PersistComboBoxLookupValue(ByVal fieldName As String, ByVal fieldValue As String)
    fieldValue = Application.WorksheetFunction.Trim(CStr(fieldValue))
    If fieldValue = "" Then Exit Sub
    If IsPlaceholderValue(fieldName, fieldValue) Then Exit Sub
    If IsLookupAutoLearnEnabled(fieldName) = False Then
        If LookupValueExists(fieldName, fieldValue) = False Then Exit Sub
    End If
    AddLookupValue fieldName, fieldValue
End Sub

Public Function GetLookupValues(ByVal lookupType As String) As Collection
    Dim tbl As ListObject
    Dim colValues As New Collection
    Dim cell As Range
    EnsureLookupTable lookupType
    SortLookupTable lookupType
    Set tbl = GetLookupTable(lookupType)
    If tbl Is Nothing Then
        Set GetLookupValues = colValues
        Exit Function
    End If
    If tbl.DataBodyRange Is Nothing Then
        Set GetLookupValues = colValues
        Exit Function
    End If
    For Each cell In tbl.ListColumns(LOOKUP_VALUE_COLUMN).DataBodyRange.Cells
        If Len(Trim$(CStr(cell.Value))) > 0 Then colValues.Add CStr(cell.Value)
    Next cell
    Set GetLookupValues = colValues
End Function

Public Function GetLookupTypeFromSource(ByVal DropdownSource As String) As String
    GetLookupTypeFromSource = Trim$(DropdownSource)
End Function

Public Sub ConfigureLookupComboBox(ByRef cbo As MSForms.ComboBox)
    With cbo
        .Style = fmStyleDropDownCombo
        .MatchEntry = fmMatchEntryComplete
        .MatchRequired = False
        .ListRows = 12
    End With
End Sub

Public Sub PopulateComboBox(ByRef cbo As MSForms.ComboBox, ByVal lookupType As String)
    Dim colValues As Collection
    Dim vItem As Variant
    ConfigureLookupComboBox cbo
    cbo.Clear
    cbo.AddItem ""
    Set colValues = GetLookupValues(lookupType)
    For Each vItem In colValues
        cbo.AddItem CStr(vItem)
    Next vItem
End Sub

Public Sub TestLookupRetrieval()
    Dim colBrands As Collection
    Dim vItem As Variant
    Set colBrands = GetLookupValues("BRAND")
    Debug.Print "===== BRAND LOOKUPS ====="
    For Each vItem In colBrands
        Debug.Print vItem
    Next vItem
    Debug.Print "========================="
End Sub
