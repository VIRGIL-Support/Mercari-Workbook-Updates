Attribute VB_Name = "modFieldDefinitions"
Option Explicit

Public Enum eControlStyle

    STANDARD_TEXT = 1
    MULTILINE_TEXT = 2
    EDITABLE_DROPDOWN = 3
    BOOLEAN_WITH_NOTES = 4
    BOOLEAN_WITH_NOTES_NA = 5

End Enum

Public Type tFieldDefinition

    fieldName As String
    fieldLabel As String

    tabName As String

    ControlStyle As eControlStyle

    Required As Boolean

    DropdownSource As String

    dataColumn As Long

    SaveToDataSheet As Boolean

End Type

Public FieldDefinitions() As tFieldDefinition


Public Sub LoadFieldDefinitions()

    Dim i As Long

    ReDim FieldDefinitions(1 To 999)

    i = 1

    ' =====================================================
    ' IDENTITY
    ' =====================================================

    AddField i, "BRAND", "Brand", "IDENTITY", EDITABLE_DROPDOWN, False, "TABLE_BRANDS", 1
    AddField i, "MODEL_NAME", "Model Name", "IDENTITY", STANDARD_TEXT, False, "", 2
    AddField i, "MODEL_NUMBER", "Model Number", "IDENTITY", STANDARD_TEXT, False, "", 3
    AddField i, "UPC", "UPC", "IDENTITY", STANDARD_TEXT, False, "", 4

    ' =====================================================
    ' FEATURES
    ' =====================================================

    AddField i, "COLOR", "Color", "FEATURES", EDITABLE_DROPDOWN, False, "TABLE_COLORS", 5
    AddField i, "MATERIAL", "Material", "FEATURES", EDITABLE_DROPDOWN, False, "TABLE_MATERIALS", 6
    AddField i, "STYLE", "Style", "FEATURES", EDITABLE_DROPDOWN, False, "TABLE_STYLES", 7
    AddField i, "PATTERN", "Pattern", "FEATURES", EDITABLE_DROPDOWN, False, "TABLE_PATTERNS", 8
    AddField i, "CHARACTER_OR_FRANCHISE", "Character or Franchise", "FEATURES", STANDARD_TEXT, False, "", 9
    AddField i, "ADDITIONAL_FEATURES", "Additional Features", "FEATURES", MULTILINE_TEXT, False, "", 10

    ' =====================================================
    ' SPECS
    ' =====================================================

    ' ITEM DIMENSIONS
    AddField i, "ITEM_HEIGHT", "Item Dimensions", "SPECS", STANDARD_TEXT, False, "", 11
    AddField i, "ITEM_WIDTH", "Item Width", "SPECS", STANDARD_TEXT, False, "", 12
    AddField i, "ITEM_LENGTH", "Item Length", "SPECS", STANDARD_TEXT, False, "", 13

    ' TOTAL CONTENT DIMENSIONS
    AddField i, "TOTAL_CONTENT_HEIGHT", "Total Content Dimensions", "SPECS", STANDARD_TEXT, False, "", 50
    AddField i, "TOTAL_CONTENT_WIDTH", "Total Content Width", "SPECS", STANDARD_TEXT, False, "", 51
    AddField i, "TOTAL_CONTENT_LENGTH", "Total Content Length", "SPECS", STANDARD_TEXT, False, "", 52

    ' TOTAL CONTENT WEIGHT
    AddField i, "WEIGHT_POUNDS", "Total Content Weight", "SPECS", STANDARD_TEXT, False, "", 14
    AddField i, "WEIGHT_OUNCES", "Weight Ounces", "SPECS", STANDARD_TEXT, False, "", 15

    ' REMAINING SPECS
    AddField i, "CAPACITY", "Capacity", "SPECS", STANDARD_TEXT, False, "", 16
    AddField i, "SCREEN_SIZE", "Screen Size", "SPECS", STANDARD_TEXT, False, "", 17
    AddField i, "CLOTHING_MEASUREMENTS", "Clothing Measurements", "SPECS", STANDARD_TEXT, False, "", 18
    AddField i, "SHOE_SIZE", "Shoe Size", "SPECS", STANDARD_TEXT, False, "", 19
    AddField i, "SHOE_WIDTH", "Shoe Width", "SPECS", STANDARD_TEXT, False, "", 20

    ' =====================================================
    ' CONDITION
    ' =====================================================

    AddField i, "OVERALL_CONDITION", "Overall Condition", "CONDITION", EDITABLE_DROPDOWN, True, "TABLE_CONDITIONS", 21
    AddField i, "DEFECTS", "Defects", "CONDITION", STANDARD_TEXT, False, "", 22
    AddField i, "WEAR", "Wear", "CONDITION", STANDARD_TEXT, False, "", 23
    AddField i, "DAMAGE", "Damage", "CONDITION", STANDARD_TEXT, False, "", 24
    AddField i, "STAINS", "Stains", "CONDITION", STANDARD_TEXT, False, "", 25
    AddField i, "MISSING_PARTS", "Missing Parts", "CONDITION", STANDARD_TEXT, False, "", 26
    AddField i, "FUNCTIONAL_STATUS", "Functional Status", "CONDITION", EDITABLE_DROPDOWN, False, "TABLE_FUNCTIONAL_STATUS", 27
    AddField i, "ITEM_TESTED", "Item Tested", "CONDITION", BOOLEAN_WITH_NOTES_NA, False, "", 28
    AddField i, "ODORS", "Odors", "CONDITION", EDITABLE_DROPDOWN, False, "TABLE_ODORS", 29
    AddField i, "RESTORATION_NOTES", "Restoration Notes", "CONDITION", STANDARD_TEXT, False, "", 30

    ' =====================================================
    ' ACCESSORIES
    ' =====================================================

    AddField i, "ORIGINAL_PACKAGING", "Original Packaging", "ACCESSORIES", BOOLEAN_WITH_NOTES_NA, False, "", 31
    AddField i, "MANUALS_INCLUDED", "Manuals Included", "ACCESSORIES", BOOLEAN_WITH_NOTES_NA, False, "", 32
    AddField i, "CHARGERS_INCLUDED", "Chargers Included", "ACCESSORIES", BOOLEAN_WITH_NOTES_NA, False, "", 33
    AddField i, "CABLES_INCLUDED", "Cables Included", "ACCESSORIES", BOOLEAN_WITH_NOTES_NA, False, "", 34
    AddField i, "REMOTE_INCLUDED", "Remote Included", "ACCESSORIES", BOOLEAN_WITH_NOTES_NA, False, "", 35
    AddField i, "OTHER_INCLUDED_ITEMS", "Other Included Items", "ACCESSORIES", STANDARD_TEXT, False, "", 36
    AddField i, "MISSING_ITEMS", "Missing Items", "ACCESSORIES", STANDARD_TEXT, False, "", 37

    ' =====================================================
    ' SHIPPING
    ' =====================================================

    AddField i, "FRAGILE_ITEM", "Fragile Item", "SHIPPING", BOOLEAN_WITH_NOTES, False, "", 38
    AddField i, "HAZARDOUS_MATERIAL", "Hazardous Material", "SHIPPING", BOOLEAN_WITH_NOTES, False, "", 39
    AddField i, "SHIPPING_RESTRICTIONS", "Shipping Restrictions", "SHIPPING", EDITABLE_DROPDOWN, False, "TABLE_SHIPPING_RESTRICTIONS", 40
    AddField i, "LOCAL_PICKUP_ALLOWED", "Local Pickup Allowed", "SHIPPING", BOOLEAN_WITH_NOTES, False, "", 41

    ' =====================================================
    ' HISTORY
    ' =====================================================

    AddField i, "OWNERSHIP_HISTORY", "Ownership History", "HISTORY", STANDARD_TEXT, False, "", 42
    AddField i, "PURCHASE_SOURCE", "Purchase Source", "HISTORY", STANDARD_TEXT, False, "", 43
    AddField i, "USAGE_HISTORY", "Usage History", "HISTORY", STANDARD_TEXT, False, "", 44
    AddField i, "STORAGE_ENVIRONMENT", "Storage Environment", "HISTORY", EDITABLE_DROPDOWN, False, "TABLE_STORAGE_ENVIRONMENTS", 45
    AddField i, "SMOKE_FREE_HOME", "Smoke-Free Home", "HISTORY", BOOLEAN_WITH_NOTES, False, "", 46
    AddField i, "PET_EXPOSURE", "Pet Exposure", "HISTORY", BOOLEAN_WITH_NOTES, False, "", 47
    AddField i, "VINTAGE_NOTES", "Vintage Notes", "HISTORY", STANDARD_TEXT, False, "", 48
    AddField i, "AUTHENTICITY_NOTES", "Authenticity Notes", "HISTORY", STANDARD_TEXT, False, "", 49
    AddField i, "OTHER_NOTES", "Other Notes", "HISTORY", STANDARD_TEXT, False, "", 53

    ReDim Preserve FieldDefinitions(1 To i - 1)

End Sub


Private Sub AddField( _
    ByRef i As Long, _
    ByVal pFieldName As String, _
    ByVal pFieldLabel As String, _
    ByVal pTabName As String, _
    ByVal pControlStyle As eControlStyle, _
    Optional ByVal pRequired As Boolean = False, _
    Optional ByVal pDropdownSource As String = "", _
    Optional ByVal pDataColumn As Long = 0)

    FieldDefinitions(i).fieldName = pFieldName
    FieldDefinitions(i).fieldLabel = pFieldLabel

    FieldDefinitions(i).tabName = pTabName

    FieldDefinitions(i).ControlStyle = pControlStyle

    FieldDefinitions(i).Required = pRequired

    FieldDefinitions(i).DropdownSource = pDropdownSource

    FieldDefinitions(i).SaveToDataSheet = True

    FieldDefinitions(i).dataColumn = pDataColumn

    i = i + 1

End Sub


Public Function GetFieldDefinition( _
    ByVal fieldName As String _
) As tFieldDefinition

    Dim i As Long

    For i = LBound(FieldDefinitions) To UBound(FieldDefinitions)

        If FieldDefinitions(i).fieldName = fieldName Then

            GetFieldDefinition = FieldDefinitions(i)

            Exit Function

        End If

    Next i

End Function

