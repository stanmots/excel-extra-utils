Attribute VB_Name = "SortRangesModule"
Option Explicit
Option Private Module

'[CONSTANTS]
Private Const p_ERROR_FUNCTION_NAME As String = ERROR_FUNCTION_NAME & "[SortRangesByNumberInColumn]"
Private Const p_INCORRECT_ARGS_ERROR As String = ERROR_TITLE & INCORRECT_ARGS_ERROR_MSG & p_ERROR_FUNCTION_NAME

'the main sorting function that can sort ranges by the number found in sorting column
Public Sub SortRangesByNumberInColumn(ByVal SortingColumn As String, ByVal SortingOffsets As Collection, ByVal WorksheetName As String, Optional ByVal wb As Workbook)

'[1] errors-checking
If Len(SortingColumn) = 0 Or Len(WorksheetName) = 0 Or SortingOffsets Is Nothing Then
    ProgressBarForm.AddMessageToDetailsBox p_INCORRECT_ARGS_ERROR
    Exit Sub
End If

If DoesCollectionContainKey(SortingOffsets, SORTING_TOP_LEFT_CELL_ROW_OFFSET_KEY) = False Or _
    DoesCollectionContainKey(SortingOffsets, SORTING_RIGHT_BOTTOM_CELL_ROW_OFFSET_KEY) = False Or _
    DoesCollectionContainKey(SortingOffsets, SORTING_TOP_LEFT_CELL_COLUMN_OFFSET_KEY) = False Or _
    DoesCollectionContainKey(SortingOffsets, SORTING_RIGHT_BOTTOM_CELL_COLUMN_OFFSET_KEY) = False Then
    
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

'[2] block with all the required vars and consts
Const TEMP_SHEET_NAME As String = "TempSheetForSortVbaProgram"

Dim MainSortingStorage As New CellsStorage
Dim SortedValues As Variant
Dim i As Long, j As Long

Dim DestRangeAddress As String, SourceRangeAddress As String

'[3] store sorting numbers from the chosen column
If CommonHelpers.SetAddressesAndValuesFromChosenColumn(MainSortingStorage, SortingColumn, WorksheetName, wb) = False Then
    ProgressBarForm.AddMessageToDetailsBox ERROR_TITLE & CANNOT_FIND_NECESSARY_VALUES_IN_CHOSEN_COLUMN_ERROR_MSG & p_ERROR_FUNCTION_NAME
    Exit Sub
End If

'[4] original values sorting
SortedValues = CollectionToArray(MainSortingStorage.CellsValues)
RandomizedQuickSort SortedValues, LBound(SortedValues), UBound(SortedValues)

'[5] create temp sheet
If DoesSheetExist(TEMP_SHEET_NAME, wb) = False Then
    AddHiddenSheet TEMP_SHEET_NAME, wb
End If

'[6] sort current worksheet
ProgressBarForm.SetLoopsParameters 90, UBound(SortedValues)
For i = 0 To UBound(SortedValues)
    ProgressBarForm.SetCurrentOperationLabelText CURRENT_SORTING_NUMBER_NAME_MSG & CStr(SortedValues(i))
    'i+1 - we need plus 1 because a collection's first index is 1,
    'while an array's first index is 0
    If SortedValues(i) <> MainSortingStorage.CellsValues.item(i + 1) Then
    
        For j = i + 1 To UBound(SortedValues)

            If SortedValues(i) = MainSortingStorage.CellsValues.item(j + 1) Then
            
                '[6][1] swap items in the collection with original values
                'i, j - array's indexies
                'i+1, j+1 - corresponding collection's indexies
                SwapItemsInCollection MainSortingStorage.CellsValues, i + 1, j + 1
                
                '[6][2] swap items in the worksheet
                With wb.Worksheets(WorksheetName)
                
                    '[6][2][1] store the first swapping item to the temp sheet
                    DestRangeAddress = CommonHelpers.GetRangeStringFromOffsets(.Range(MainSortingStorage.CellsAddresses.item(i + 1)), _
                        SortingOffsets.item(SORTING_TOP_LEFT_CELL_COLUMN_OFFSET_KEY), SortingOffsets.item(SORTING_TOP_LEFT_CELL_ROW_OFFSET_KEY), _
                        SortingOffsets.item(SORTING_RIGHT_BOTTOM_CELL_COLUMN_OFFSET_KEY), SortingOffsets.item(SORTING_RIGHT_BOTTOM_CELL_ROW_OFFSET_KEY))
                       
                    .Range(DestRangeAddress).Copy Destination:=wb.Worksheets(TEMP_SHEET_NAME).Range(DestRangeAddress)
                                   
                    '[6][2][2] copy the range from the(source, j-th) to the(dest, i-th)
                    SourceRangeAddress = CommonHelpers.GetRangeStringFromOffsets(.Range(MainSortingStorage.CellsAddresses.item(j + 1)), _
                        SortingOffsets.item(SORTING_TOP_LEFT_CELL_COLUMN_OFFSET_KEY), SortingOffsets.item(SORTING_TOP_LEFT_CELL_ROW_OFFSET_KEY), _
                        SortingOffsets.item(SORTING_RIGHT_BOTTOM_CELL_COLUMN_OFFSET_KEY), SortingOffsets.item(SORTING_RIGHT_BOTTOM_CELL_ROW_OFFSET_KEY))
                                   
                    .Range(SourceRangeAddress).Copy Destination:=.Range(DestRangeAddress)
                    
                    '[6][2][3] copy from the temp sheet to the source range
                    wb.Worksheets(TEMP_SHEET_NAME).Range(DestRangeAddress).Copy _
                        Destination:=.Range(SourceRangeAddress)
                End With
                Exit For
            End If
        Next j
    End If
    
    ProgressBarForm.IncreaseProgressInsideLoop
    
Next i

'[7] delete temporary sheet
DeleteSheetWithoutAlerts TEMP_SHEET_NAME, wb

'[8](optional) paste the serial numbers in the current worksheet
If DoesCollectionContainKey(SortingOffsets, SORTING_SERIAL_CELL_ROW_OFFSET_KEY) = True And _
    DoesCollectionContainKey(SortingOffsets, SORTING_SERIAL_CELL_COLUMN_OFFSET_KEY) = True Then
    
    Dim SerialCellAddress As String
        
    For i = 1 To MainSortingStorage.CellsAddresses.Count
        With wb.Worksheets(WorksheetName)
            
            SerialCellAddress = CommonHelpers.GetRangeStringFromOffsets(.Range(MainSortingStorage.CellsAddresses.item(i)), _
                SortingOffsets.item(SORTING_SERIAL_CELL_COLUMN_OFFSET_KEY), SortingOffsets.item(SORTING_SERIAL_CELL_ROW_OFFSET_KEY))
            
            .Range(SerialCellAddress).Value = i
                
        End With
    Next i
End If

Set MainSortingStorage = Nothing

ProgressBarForm.IncreaseProgressByPercent 10

End Sub
