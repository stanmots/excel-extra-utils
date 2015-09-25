Attribute VB_Name = "CommonHelpers"
Option Explicit
Option Private Module

Public Function GetRangeStringFromOffsets(ByVal BaseCellRange As Range, ByVal TLCColumnOffset As Long, ByVal TLCRowOffset As Long, _
    Optional ByVal RBCColumnOffset As Variant = "", Optional ByVal RBCRowOffset As Variant = "") As String

'calculate top-left cell address
Dim ColumnNumber As Long
ColumnNumber = BaseCellRange.Column + TLCColumnOffset

Dim ColumnLetter As String
ColumnLetter = ColumnNumberToLetter(ColumnNumber)

Dim TLCAddress As String
TLCAddress = ColumnLetter & CStr(BaseCellRange.Row + TLCRowOffset)
    
If TypeName(RBCColumnOffset) = "String" Or TypeName(RBCRowOffset) = "String" Then
    GetRangeStringFromOffsets = TLCAddress
Else
    'calculate right-bottom cell address
    ColumnNumber = BaseCellRange.Column + RBCColumnOffset
        
    ColumnLetter = ColumnNumberToLetter(ColumnNumber)
    
    Dim RBCAddress As String
    RBCAddress = ColumnLetter & CStr(BaseCellRange.Row + RBCRowOffset)
                            
    GetRangeStringFromOffsets = TLCAddress & ":" & RBCAddress

End If

End Function

Public Function ExtractFileNameFromPath(ByVal Path As String) As String

ExtractFileNameFromPath = Right(Path, Len(Path) - InStrRev(Path, "\"))

End Function

Public Function GetFormattedStringFromOffsets(ByVal TLCColumnOffset As Long, ByVal TLCRowOffset As Long, _
    Optional ByVal RBCColumnOffset As Variant = "", Optional ByVal RBCRowOffset As Variant = "", _
    Optional ByVal SRCColumnOffset As Variant = "", Optional ByVal SRCRowOffset As Variant = "") As String
    
Const OFFSETS_DELIMITER As String = ","

GetFormattedStringFromOffsets = "(" & CStr(TLCColumnOffset) & OFFSETS_DELIMITER & CStr(TLCRowOffset)

If TypeName(RBCColumnOffset) <> "String" And TypeName(RBCRowOffset) <> "String" Then

    GetFormattedStringFromOffsets = GetFormattedStringFromOffsets & OFFSETS_DELIMITER & CStr(RBCColumnOffset) & OFFSETS_DELIMITER & CStr(RBCRowOffset)
    
    If TypeName(SRCColumnOffset) <> "String" And TypeName(SRCRowOffset) <> "String" Then
        GetFormattedStringFromOffsets = GetFormattedStringFromOffsets & OFFSETS_DELIMITER & CStr(SRCColumnOffset) & OFFSETS_DELIMITER & CStr(SRCRowOffset)
    End If

End If

GetFormattedStringFromOffsets = GetFormattedStringFromOffsets & ")"
    
End Function

Public Function SetAddressesAndValuesFromChosenColumn(ByVal MainStorage As CellsStorage, ByVal ColumnName As String, ByVal ws As String, Optional wb As Workbook) As Boolean

If wb Is Nothing Then
    Set wb = ThisWorkbook
End If

If DoesSheetExist(ws, wb) = False Then
    SetAddressesAndValuesFromChosenColumn = False
    Exit Function
End If

If MainStorage Is Nothing Or Len(ColumnName) = 0 Then
    SetAddressesAndValuesFromChosenColumn = False
    Exit Function
End If

Dim ColumnCell As Range

With wb.Worksheets(ws)
    For Each ColumnCell In .UsedRange.Columns(ColumnName).Cells
        If ColumnCell.Value <> vbNullString And IsNumeric(ColumnCell.Value) = True Then
            MainStorage.CellsValues.Add ColumnCell.Value
            MainStorage.CellsAddresses.Add ColumnCell.Address(RowAbsolute:=False, _
                                                    ColumnAbsolute:=False)
        End If
    Next
End With

If MainStorage.CellsValues.Count < 1 Or MainStorage.CellsAddresses.Count < 1 Then
    SetAddressesAndValuesFromChosenColumn = False
    Exit Function
End If

SetAddressesAndValuesFromChosenColumn = True

End Function
    
'10092543 - Light Yellow
Public Function ShowEditColorDialog(Optional InitialColor As Long = 10092543, Optional ByVal PaletteIndex As Long = 1) As Long

Dim OriginalColorInPalette As Long
Dim intR As Integer, intG As Integer, intB As Integer
OriginalColorInPalette = ThisWorkbook.Colors(PaletteIndex)

'get the RGB values of lngInitialColor
intR = InitialColor And 255
intG = InitialColor \ 256 And 255
intB = InitialColor \ 256 ^ 2 And 255
        
If Application.Dialogs(xlDialogEditColor).Show(PaletteIndex, intR, intG, intB) = True Then
    ShowEditColorDialog = ThisWorkbook.Colors(PaletteIndex)
    
    'restore original color in palette
    ThisWorkbook.Colors(PaletteIndex) = OriginalColorInPalette
Else
    ShowEditColorDialog = -1
End If

End Function
