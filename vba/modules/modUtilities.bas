Attribute VB_Name = "modUtilities"
Option Explicit

' =====================================================
' GET WORKSHEET
' =====================================================

Public Function GetWorksheet(sheetName As String) As Worksheet

    Set GetWorksheet = ThisWorkbook.Worksheets(sheetName)

End Function

' =====================================================
' GET LAST ROW
' =====================================================

Public Function GetLastRow(ws As Worksheet, columnNumber As Long) As Long

    GetLastRow = ws.Cells(ws.Rows.Count, columnNumber).End(xlUp).Row

End Function

' =====================================================
' GET SETTING VALUE
' =====================================================

Public Function GetSettingValue(settingName As String) As String

    Dim wsSettings As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set wsSettings = ThisWorkbook.Sheets(WS_SETTINGS)

    lastRow = wsSettings.Cells(wsSettings.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow

        If wsSettings.Cells(i, 1).Value = settingName Then

            GetSettingValue = wsSettings.Cells(i, 2).Value
            Exit Function

        End If

    Next i

    GetSettingValue = ""

End Function

' =====================================================
' SET SETTING VALUE
' =====================================================

Public Sub SetSettingValue(settingName As String, settingValue As String)

    Dim wsSettings As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set wsSettings = ThisWorkbook.Sheets(WS_SETTINGS)

    lastRow = wsSettings.Cells(wsSettings.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow

        If wsSettings.Cells(i, 1).Value = settingName Then

            wsSettings.Cells(i, 2).Value = settingValue
            Exit Sub

        End If

    Next i

End Sub

' =====================================================
' CREATE FOLDER IF MISSING
' =====================================================

Public Sub CreateFolderIfMissing(folderPath As String)

    If Dir(folderPath, vbDirectory) = "" Then

        MkDir folderPath

    End If

End Sub

' =====================================================
' COPY TEXT TO CLIPBOARD
' =====================================================

Public Sub CopyTextToClipboard(textToCopy As String)

    Dim objHTML As Object

    Set objHTML = CreateObject("htmlfile")

    objHTML.ParentWindow.ClipboardData.SetData "text", textToCopy

End Sub

' =====================================================
' GET PLACEHOLDER TEXT
' =====================================================

Public Function GetPlaceholderText( _
    ByVal fieldName As String _
) As String

    Select Case fieldName

        ' =================================================
        ' IDENTITY
        ' =================================================

        Case "BRAND"
            GetPlaceholderText = "Example: Sony, Nintendo, Apple"

        Case "MODEL_NAME"
            GetPlaceholderText = _
                "Example: Walkman Sports Cassette Player"

        Case "MODEL_NUMBER"
            GetPlaceholderText = "Example: WM-FS111"

        Case "UPC"
            GetPlaceholderText = _
                "12-digit UPC if available"

        ' =================================================
        ' FEATURES
        ' =================================================

        Case "COLOR"
            GetPlaceholderText = _
                "Example: Black, Red, Blue"

        Case "MATERIAL"
            GetPlaceholderText = _
                "Example: Plastic, Cotton, Aluminum"

        Case "STYLE"
            GetPlaceholderText = _
                "Example: Retro, Modern, Casual"

        Case "PATTERN"
            GetPlaceholderText = _
                "Example: Solid, Floral, Striped"

        Case "CHARACTER_OR_FRANCHISE"
            GetPlaceholderText = _
                "Example: Disney, Pokémon, Star Wars"

        Case "ADDITIONAL_FEATURES"
            GetPlaceholderText = _
                "List notable features, special functions, bundled capabilities, limited editions, etc."

        ' =================================================
        ' SPECS
        ' =================================================

        Case "ITEM_HEIGHT"
            GetPlaceholderText = "W"

        Case "ITEM_WIDTH"
            GetPlaceholderText = "H"

        Case "ITEM_LENGTH"
            GetPlaceholderText = "L"

        Case "TOTAL_CONTENT_HEIGHT"
            GetPlaceholderText = "W"

        Case "TOTAL_CONTENT_WIDTH"
            GetPlaceholderText = "H"

        Case "TOTAL_CONTENT_LENGTH"
            GetPlaceholderText = "L"

        Case "WEIGHT_POUNDS"
            GetPlaceholderText = "lbs"

        Case "WEIGHT_OUNCES"
            GetPlaceholderText = "oz"

        Case "CAPACITY"
            GetPlaceholderText = _
                "Example: 64GB, 2TB, 32oz"

        Case "SCREEN_SIZE"
            GetPlaceholderText = _
                "Example: 15.6 inches"

        Case "CLOTHING_MEASUREMENTS"
            GetPlaceholderText = _
                "Example: Pit to pit 22"", inseam 30"", waist 34"""

        Case "SHOE_SIZE"
            GetPlaceholderText = _
                "Example: Men’s 10.5"

        Case "SHOE_WIDTH"
            GetPlaceholderText = _
                "Example: Regular, Wide, Narrow"

        ' =================================================
        ' CONDITION
        ' =================================================

        Case "OVERALL_CONDITION"
            GetPlaceholderText = "Select the item's overall condition"

        Case "DEFECTS"
            GetPlaceholderText = _
                "Describe defects, flaws, cracks, dents, chips, missing pieces, etc."

        Case "WEAR"
            GetPlaceholderText = _
                "Describe visible wear from use or age"

        Case "DAMAGE"
            GetPlaceholderText = _
                "Describe any damage or repairs"

        Case "STAINS"
            GetPlaceholderText = _
                "Describe stains, discoloration, fading, yellowing, etc."

        Case "MISSING_PARTS"
            GetPlaceholderText = _
                "List missing components, accessories, or pieces"

        Case "FUNCTIONAL_STATUS"
            GetPlaceholderText = "Select or type functional status"

        Case "ODORS"
            GetPlaceholderText = "Select or type odor status"

        Case "RESTORATION_NOTES"
            GetPlaceholderText = _
                "Describe restoration, repairs, cleaning, or refurbishment performed"

        Case "ITEM_TESTED_NOTES"
            GetPlaceholderText = _
                "Add testing details, methods, limitations, or results"

        ' =================================================
        ' ACCESSORIES
        ' =================================================

        Case "ORIGINAL_PACKAGING_NOTES"
            GetPlaceholderText = "Packaging notes"

        Case "MANUALS_INCLUDED_NOTES"
            GetPlaceholderText = "Manuals notes"

        Case "CHARGERS_INCLUDED_NOTES"
            GetPlaceholderText = "Chargers notes"

        Case "CABLES_INCLUDED_NOTES"
            GetPlaceholderText = "Cables notes"

        Case "REMOTE_INCLUDED_NOTES"
            GetPlaceholderText = "Remote notes"

        Case "OTHER_INCLUDED_ITEMS"
            GetPlaceholderText = _
                "List additional included items not already covered"

        Case "MISSING_ITEMS"
            GetPlaceholderText = _
                "List known missing accessories or components"

        ' =================================================
        ' SHIPPING
        ' =================================================

        Case "FRAGILE_ITEM_NOTES"
            GetPlaceholderText = _
                "Describe fragile materials or special handling concerns"

        Case "HAZARDOUS_MATERIAL_NOTES"
            GetPlaceholderText = _
                "Describe batteries, chemicals, flammables, or shipping concerns"

        Case "LOCAL_PICKUP_ALLOWED_NOTES"
            GetPlaceholderText = _
                "Add pickup details or restrictions if applicable"

        Case "SHIPPING_RESTRICTIONS"
            GetPlaceholderText = "Select or type shipping restrictions"

        Case "STORAGE_ENVIRONMENT"
            GetPlaceholderText = "Select or type storage environment"

        ' =================================================
        ' HISTORY
        ' =================================================

        Case "OWNERSHIP_HISTORY"
            GetPlaceholderText = _
                "Describe previous ownership history if known"

        Case "PURCHASE_SOURCE"
            GetPlaceholderText = _
                "Example: Estate sale, retail store, thrift store, original owner"

        Case "USAGE_HISTORY"
            GetPlaceholderText = _
                "Describe how the item was used or stored"

        Case "SMOKE_FREE_HOME_NOTES"
            GetPlaceholderText = _
                "Add details about smoke exposure if applicable"

        Case "PET_EXPOSURE_NOTES"
            GetPlaceholderText = _
                "Describe pet exposure if applicable"

        Case "VINTAGE_NOTES"
            GetPlaceholderText = _
                "Include production era, collectible significance, rarity, etc."

        Case "AUTHENTICITY_NOTES"
            GetPlaceholderText = _
                "Describe authentication, serial verification, provenance, etc."

        Case "OTHER_NOTES"
            GetPlaceholderText = _
                "Add any additional details useful to buyers"

    End Select

End Function

' =====================================================
' IS PLACEHOLDER VALUE
' =====================================================

Public Function IsPlaceholderValue( _
    ByVal fieldName As String, _
    ByVal fieldValue As String _
) As Boolean

    If Trim(fieldValue) = "" Then Exit Function

    If Trim(fieldValue) = _
        Trim(GetPlaceholderText(fieldName)) Then

        IsPlaceholderValue = True

    End If

End Function

