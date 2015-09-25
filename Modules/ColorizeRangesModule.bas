Attribute VB_Name = "ColorizeRangesModule"
Option Explicit
Option Private Module

'[CONSTANTS]
Private Const p_ERROR_FUNCTION_NAME As String = ERROR_FUNCTION_NAME & "[ColorizeRanges]"
Private Const p_INCORRECT_ARGS_ERROR As String = ERROR_TITLE & INCORRECT_ARGS_ERROR_MSG & p_ERROR_FUNCTION_NAME

Public Sub ColorizeRanges(ByVal ColoringOffsets As Collection, ByVal ColoringColumn As String, ByVal BaseRange As String, ByVal WorksheetName As String, ByVal Color As Long, Optional ByVal wb As Workbook)

'[1] errors-checking
If Len(ColoringColumn) = 0 Or Len(WorksheetName) = 0 Or Len(BaseRange) = 0 Or ColoringOffsets Is Nothing Then
    ProgressBarForm.AddMessageToDetailsBox p_INCORRECT_ARGS_ERROR
    Exit Sub
End If

If DoesCollectionContainKey(ColoringOffsets, COLORING_SOUGHTFORRANGE_TOP_LEFT_CELL_ROW_OFFSET_KEY) = False Or _
    DoesCollectionContainKey(ColoringOffsets, COLORING_SOUGHTFORRANGE_TOP_LEFT_CELL_COLUMN_OFFSET_KEY) = False Or _
    DoesCollectionContainKey(ColoringOffsets, COLORING_SOUGHTFORRANGE_RIGHT_BOTTOM_CELL_ROW_OFFSET_KEY) = False Or _
    DoesCollectionContainKey(ColoringOffsets, COLORING_SOUGHTFORRANGE_RIGHT_BOTTOM_CELL_COLUMN_OFFSET_KEY) = False Then
    
    ProgressBarForm.AddMessageToDetailsBox p_INCORRECT_ARGS_ERROR
    Exit Sub
End If

If wb Is Nothing Then
    Set wb = ThisWorkbook
End If

If DoesSheetExist(WorksheetName, wb) = False Then
    ProgressBarForm.AddMessageToDetailsBox p_INCORRECT_ARGS_ERROR
    Exit Sub
End If

'[2] store numbers from the chosen column
Dim MainColoringStorage As New CellsStorage
If CommonHelpers.SetAddressesAndValuesFromChosenColumn(MainColoringStorage, ColoringColumn, WorksheetName, wb) = False Then
    ProgressBarForm.AddMessageToDetailsBox ERROR_TITLE & CANNOT_FIND_NECESSARY_VALUES_IN_CHOSEN_COLUMN_ERROR_MSG & p_ERROR_FUNCTION_NAME
    Exit Sub
End If

'[3] find all cells in the BaseRange with the required color
Dim br As Range
Set br = wb.Worksheets(WorksheetName).Range(BaseRange)
Dim i As Long

Dim CellsForColorizing As New Collection

For i = 1 To br.Cells.Count
    If br.Cells(i).Interior.Color = Color Then
        CellsForColorizing.Add i
    End If
Next

If CellsForColorizing.Count = 0 Then
    ProgressBarForm.AddMessageToDetailsBox WARNING_TITLE & THERE_ARE_NO_CELLS_WITH_COLOR_ERROR_MSG
    Exit Sub
End If

'[4] colorizing
Dim j As Long

For j = 1 To MainColoringStorage.CellsAddresses.Count

    Dim SoughtForRangeStr As String
    
    SoughtForRangeStr = CommonHelpers.GetRangeStringFromOffsets(wb.Worksheets(WorksheetName).Range(MainColoringStorage.CellsAddresses.item(j)), _
        ColoringOffsets.item(COLORING_SOUGHTFORRANGE_TOP_LEFT_CELL_COLUMN_OFFSET_KEY), ColoringOffsets.item(COLORING_SOUGHTFORRANGE_TOP_LEFT_CELL_ROW_OFFSET_KEY), _
        ColoringOffsets.item(COLORING_SOUGHTFORRANGE_RIGHT_BOTTOM_CELL_COLUMN_OFFSET_KEY), ColoringOffsets.item(COLORING_SOUGHTFORRANGE_RIGHT_BOTTOM_CELL_ROW_OFFSET_KEY))
    
    Dim SoughtForRange As Range
    Set SoughtForRange = wb.Worksheets(WorksheetName).Range(SoughtForRangeStr)
    
    Dim ci As Variant
    
    For Each ci In CellsForColorizing
        SoughtForRange.Cells(CLng(ci)).Interior.Color = Color
    Next
    
Next j

'[5] memory clean-up
Set MainColoringStorage = Nothing
Set CellsForColorizing = Nothing

End Sub
