Attribute VB_Name = "modShipping"
Option Explicit

Public Type tUSPSBoxMatch
    found As Boolean
    BoxName As String
    OuterLength As Double
    OuterWidth As Double
    OuterHeight As Double
    BoxWeightOunces As Double
    PackingWeightOunces As Double
End Type

Private Const USPS_BOX_TABLE_NAME As String = "tblUSPSPriorityMailBoxes"

Private Function NormalizeTableHeader(ByVal headerText As String) As String
    headerText = UCase$(Trim$(CStr(headerText)))
    headerText = Replace(headerText, " ", "")
    headerText = Replace(headerText, "_", "")
    headerText = Replace(headerText, "-", "")
    headerText = Replace(headerText, "(", "")
    headerText = Replace(headerText, ")", "")
    headerText = Replace(headerText, ".", "")
    NormalizeTableHeader = headerText
End Function

Private Function GetTableColumnIndex(ByVal tableObject As ListObject, ByVal columnName As String) As Long
    Dim col As ListColumn
    Dim requestedName As String
    Dim actualName As String
    requestedName = NormalizeTableHeader(columnName)
    For Each col In tableObject.ListColumns
        actualName = NormalizeTableHeader(col.Name)
        If actualName = requestedName Then
            GetTableColumnIndex = col.Index
            Exit Function
        End If
        Select Case requestedName
            Case "BOXNAME"
                If actualName Like "*BOX*NAME*" Or actualName Like "*BOX*TYPE*" Then GetTableColumnIndex = col.Index: Exit Function
            Case "INNERLENGTH"
                If actualName Like "*INNER*LENGTH*" Or actualName Like "*INSIDE*LENGTH*" Then GetTableColumnIndex = col.Index: Exit Function
            Case "INNERWIDTH"
                If actualName Like "*INNER*WIDTH*" Or actualName Like "*INSIDE*WIDTH*" Then GetTableColumnIndex = col.Index: Exit Function
            Case "INNERHEIGHT"
                If actualName Like "*INNER*HEIGHT*" Or actualName Like "*INSIDE*HEIGHT*" Then GetTableColumnIndex = col.Index: Exit Function
            Case "OUTERLENGTH"
                If actualName Like "*OUTER*LENGTH*" Or actualName Like "*OUTSIDE*LENGTH*" Then GetTableColumnIndex = col.Index: Exit Function
            Case "OUTERWIDTH"
                If actualName Like "*OUTER*WIDTH*" Or actualName Like "*OUTSIDE*WIDTH*" Then GetTableColumnIndex = col.Index: Exit Function
            Case "OUTERHEIGHT"
                If actualName Like "*OUTER*HEIGHT*" Or actualName Like "*OUTSIDE*HEIGHT*" Then GetTableColumnIndex = col.Index: Exit Function
            Case "BOXWEIGHTOUNCES"
                If actualName Like "*BOX*WEIGHT*OUNCE*" Or actualName Like "*BOX*WEIGHT*OZ*" Then GetTableColumnIndex = col.Index: Exit Function
            Case "PACKINGWEIGHTOUNCES"
                If actualName Like "*PACK*WEIGHT*OUNCE*" Or actualName Like "*PACK*WEIGHT*OZ*" Then GetTableColumnIndex = col.Index: Exit Function
        End Select
    Next col
    Err.Raise vbObjectError + 6401, "GetTableColumnIndex", "Column not found in " & USPS_BOX_TABLE_NAME & ": " & columnName
End Function

Private Function GetTableColumnValue(ByVal rowRange As Range, ByVal tableObject As ListObject, ByVal columnName As String) As Variant
    GetTableColumnValue = rowRange.Cells(1, GetTableColumnIndex(tableObject, columnName)).Value
End Function

Private Sub SortThreeDescending(ByRef a As Double, ByRef b As Double, ByRef c As Double)
    Dim tempValue As Double
    If a < b Then
        tempValue = a
        a = b
        b = tempValue
    End If
    If b < c Then
        tempValue = b
        b = c
        c = tempValue
    End If
    If a < b Then
        tempValue = a
        a = b
        b = tempValue
    End If
End Sub

Private Function DimensionsFitAnyOrientation( _
    ByVal itemLength As Double, _
    ByVal itemWidth As Double, _
    ByVal itemHeight As Double, _
    ByVal boxLength As Double, _
    ByVal boxWidth As Double, _
    ByVal boxHeight As Double) As Boolean

    SortThreeDescending itemLength, itemWidth, itemHeight
    SortThreeDescending boxLength, boxWidth, boxHeight
    DimensionsFitAnyOrientation = (itemLength <= boxLength And itemWidth <= boxWidth And itemHeight <= boxHeight)
End Function

Private Function BoxVolume(ByVal boxLength As Double, ByVal boxWidth As Double, ByVal boxHeight As Double) As Double
    BoxVolume = boxLength * boxWidth * boxHeight
End Function

Public Function FindSmallestUSPSPriorityMailBox( _
    ByVal itemLength As Double, _
    ByVal itemWidth As Double, _
    ByVal itemHeight As Double) As tUSPSBoxMatch

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dataRow As Range
    Dim innerLength As Double
    Dim innerWidth As Double
    Dim innerHeight As Double
    Dim currentVolume As Double
    Dim bestVolume As Double
    Dim matchResult As tUSPSBoxMatch

    Set ws = ThisWorkbook.Worksheets(WS_TABLES)
    Set tbl = ws.ListObjects(USPS_BOX_TABLE_NAME)

    For Each dataRow In tbl.DataBodyRange.Rows
        innerLength = CDbl(GetTableColumnValue(dataRow, tbl, "INNER LENGTH"))
        innerWidth = CDbl(GetTableColumnValue(dataRow, tbl, "INNER WIDTH"))
        innerHeight = CDbl(GetTableColumnValue(dataRow, tbl, "INNER HEIGHT"))
        If DimensionsFitAnyOrientation(itemLength, itemWidth, itemHeight, innerLength, innerWidth, innerHeight) Then
            currentVolume = BoxVolume(innerLength, innerWidth, innerHeight)
            If matchResult.found = False Or currentVolume < bestVolume Then
                bestVolume = currentVolume
                matchResult.found = True
                matchResult.BoxName = CStr(GetTableColumnValue(dataRow, tbl, "BOX NAME"))
                matchResult.OuterLength = CDbl(GetTableColumnValue(dataRow, tbl, "OUTER LENGTH"))
                matchResult.OuterWidth = CDbl(GetTableColumnValue(dataRow, tbl, "OUTER WIDTH"))
                matchResult.OuterHeight = CDbl(GetTableColumnValue(dataRow, tbl, "OUTER HEIGHT"))
                matchResult.BoxWeightOunces = CDbl(GetTableColumnValue(dataRow, tbl, "BOX WEIGHT OUNCES"))
                matchResult.PackingWeightOunces = CDbl(GetTableColumnValue(dataRow, tbl, "PACKING WEIGHT OUNCES"))
            End If
        End If
    Next dataRow

    FindSmallestUSPSPriorityMailBox = matchResult
End Function

Public Sub ConvertTotalOuncesToPoundsOunces(ByVal totalOunces As Double, ByRef poundsValue As Long, ByRef ouncesValue As Long)
    Dim roundedOunces As Long
    roundedOunces = CLng(totalOunces + 0.999999)
    poundsValue = roundedOunces \ 16
    ouncesValue = roundedOunces Mod 16
End Sub
