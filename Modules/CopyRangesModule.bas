Attribute VB_Name = "CopyRangesModule"
Option Explicit
Option Private Module

'[CACHES]
'a collection that holds paths of the all workbooks that was open by CopyRangesBetweenWorksheets procedure
'you must create and release it from the calling function
Public OpenWorkbooksPaths As Collection

Public FromNumbersCache As Collection
Public ToNumbersCache As Collection

'[CONSTANTS]
Private Const p_ERROR_FUNCTION_NAME As String = ERROR_FUNCTION_NAME & "[CopyRangesBetweenWorksheets]"
Private Const p_INCORRECT_ARGS_ERROR As String = ERROR_TITLE & INCORRECT_ARGS_ERROR_MSG & p_ERROR_FUNCTION_NAME

'the main copying function that can copy ranges from one worksheet to another
'it searches for the appropriate cells according to the number in a copying column that is set in the CopyingSettings Collection
'so the number in a column is an identifier of the particular range (it must be the same in each of the worksheets)
Public Sub CopyRangesBetweenWorksheets(ByVal CopyingSettings As Collection)

'[1] errors-checking
'[1][1] check the arguments
If CopyingSettings Is Nothing Then
    ProgressBarForm.AddMessageToDetailsBox p_INCORRECT_ARGS_ERROR
    Exit Sub
End If

If DoesCollectionContainKey(CopyingSettings, COPYING_FROMWORKSHEET_KEY) = False Or _
    DoesCollectionContainKey(CopyingSettings, COPYING_FROMCOLUMN_KEY) = False Or _
    DoesCollectionContainKey(CopyingSettings, COPYING_FROM_TOP_LEFT_CELL_ROW_OFFSET_KEY) = False Or _
    DoesCollectionContainKey(CopyingSettings, COPYING_FROM_TOP_LEFT_CELL_COLUMN_OFFSET_KEY) = False Then
    
    ProgressBarForm.AddMessageToDetailsBox p_INCORRECT_ARGS_ERROR
    Exit Sub
End If

If DoesCollectionContainKey(CopyingSettings, COPYING_TOWORKSHEET_KEY) = False Or _
    DoesCollectionContainKey(CopyingSettings, COPYING_TOCOLUMN_KEY) = False Or _
    DoesCollectionContainKey(CopyingSettings, COPYING_TO_TOP_LEFT_CELL_ROW_OFFSET_KEY) = False Or _
    DoesCollectionContainKey(CopyingSettings, COPYING_TO_TOP_LEFT_CELL_COLUMN_OFFSET_KEY) = False Then
    
    ProgressBarForm.AddMessageToDetailsBox p_INCORRECT_ARGS_ERROR
    Exit Sub
End If

'[1][2] check workbooks
Dim fwb As Workbook, twb As Workbook
Dim fwbFullName As String, twbFullName As String

fwbFullName = CopyingSettings.item(COPYING_FROMWORKBOOK_KEY)
twbFullName = CopyingSettings.item(COPYING_TOWORKBOOK_KEY)

If SetCopyingWorkbook(fwbFullName, COPYING_FROMWORKBOOK_KEY, fwb, True) = False Or _
    SetCopyingWorkbook(twbFullName, COPYING_TOWORKBOOK_KEY, twb, False) = False Then
    
    ProgressBarForm.AddMessageToDetailsBox ERROR_TITLE & INCORRECT_WORKBOOKNAME_ERROR_MSG & p_ERROR_FUNCTION_NAME
    Exit Sub
End If

'[1][3] check worksheets
Dim fws As String, tws As String
fws = CopyingSettings.item(COPYING_FROMWORKSHEET_KEY)
tws = CopyingSettings.item(COPYING_TOWORKSHEET_KEY)

If DoesSheetExist(fws, fwb) = False Or DoesSheetExist(tws, twb) = False Then
    ProgressBarForm.AddMessageToDetailsBox p_INCORRECT_ARGS_ERROR
    Exit Sub
End If

ProgressBarForm.SetCurrentOperationLabelText COPYING_FROM_WORKSHEET_CO & fws & " (" & WORKBOOK_NAME & fwb.Name & ") " _
    & TO_WORKSHEET_CO & tws & " (" & WORKBOOK_NAME & twb.Name & ") "

'[1][4] check copying ranges
Dim CopyingRangesMethodName As String

Select Case CopyingSettingsHelpers.AreCopyingRangesValid(CopyingSettings.item(COPYING_FROMRANGE_KEY), CopyingSettings.item(COPYING_TORANGE_KEY))
Case CopyingRangesType.SIMILAR_RANGES_TYPE:
    CopyingRangesMethodName = "CopySimilarRanges"
Case CopyingRangesType.FROM_RANGE_TO_CELL_TYPE:
    CopyingRangesMethodName = "CopyFromRangeToCell"
Case CopyingRangesType.INCORRECT_RANGES:
    ProgressBarForm.AddMessageToDetailsBox p_INCORRECT_ARGS_ERROR
    Exit Sub
End Select

'[2] store the numbers from the chosen column
Dim FromNumbers As New CellsStorage, ToNumbers As New CellsStorage

'[2][1] retrieve from-numbers
Dim CacheKey As String
CacheKey = fwb.Name & fws

If DoesCollectionContainKey(FromNumbersCache, CacheKey) = True Then
    Set FromNumbers = FromNumbersCache.item(CacheKey)
Else
    If CommonHelpers.SetAddressesAndValuesFromChosenColumn(FromNumbers, CopyingSettings.item(COPYING_FROMCOLUMN_KEY), fws, fwb) = False Then
        ProgressBarForm.AddMessageToDetailsBox ERROR_TITLE & CANNOT_FIND_NECESSARY_VALUES_IN_CHOSEN_COLUMN_ERROR_MSG & p_ERROR_FUNCTION_NAME
        Exit Sub
    Else
        If Not FromNumbersCache Is Nothing Then FromNumbersCache.Add FromNumbers, CacheKey
    End If
End If

'[2][2] retrieve to-numbers
CacheKey = twb.Name & tws
If DoesCollectionContainKey(ToNumbersCache, CacheKey) = True Then
    Set ToNumbers = ToNumbersCache.item(CacheKey)
Else
    If CommonHelpers.SetAddressesAndValuesFromChosenColumn(ToNumbers, CopyingSettings.item(COPYING_TOCOLUMN_KEY), tws, twb) = False Then
        ProgressBarForm.AddMessageToDetailsBox ERROR_TITLE & CANNOT_FIND_NECESSARY_VALUES_IN_CHOSEN_COLUMN_ERROR_MSG & p_ERROR_FUNCTION_NAME
        Exit Sub
    Else
        If Not ToNumbersCache Is Nothing Then ToNumbersCache.Add ToNumbers, CacheKey
    End If
End If

'[3] copy the ranges that were found by the numbers between the worksheets (fws, tws)
Dim i As Long, j As Long

For i = 1 To FromNumbers.CellsAddresses.Count
    For j = 1 To ToNumbers.CellsAddresses.Count
        If FromNumbers.CellsValues.item(i) = ToNumbers.CellsValues.item(j) Then
            Application.Run CopyingRangesMethodName, CopyingSettings, _
                fwb.Worksheets(fws).Range(FromNumbers.CellsAddresses.item(i)), _
                twb.Worksheets(tws).Range(ToNumbers.CellsAddresses.item(j)), _
                fwb, twb
            Exit For
        End If
    Next j
Next i

'[4] memory clean-up
Set FromNumbers = Nothing
Set ToNumbers = Nothing

End Sub

Private Sub CopySimilarRanges(ByVal CopyingSettings As Collection, _
        ByVal FromBaseCellRange As Range, ByVal ToBaseCellRange As Range, _
        ByVal fwb As Workbook, ByVal twb As Workbook)

Dim FromRangeAddress As String, ToRangeAddress As String

FromRangeAddress = CommonHelpers.GetRangeStringFromOffsets(FromBaseCellRange, _
    CopyingSettings.item(COPYING_FROM_TOP_LEFT_CELL_COLUMN_OFFSET_KEY), CopyingSettings.item(COPYING_FROM_TOP_LEFT_CELL_ROW_OFFSET_KEY), _
    CopyingSettings.item(COPYING_FROM_RIGHT_BOTTOM_CELL_COLUMN_OFFSET_KEY), CopyingSettings.item(COPYING_FROM_RIGHT_BOTTOM_CELL_ROW_OFFSET_KEY))
                      
ToRangeAddress = CommonHelpers.GetRangeStringFromOffsets(ToBaseCellRange, _
    CopyingSettings.item(COPYING_TO_TOP_LEFT_CELL_COLUMN_OFFSET_KEY), CopyingSettings.item(COPYING_TO_TOP_LEFT_CELL_ROW_OFFSET_KEY), _
    CopyingSettings.item(COPYING_TO_RIGHT_BOTTOM_CELL_COLUMN_OFFSET_KEY), CopyingSettings.item(COPYING_TO_RIGHT_BOTTOM_CELL_ROW_OFFSET_KEY))
      
'clear previous values
'ClearContents doesn't work with merged cells
'twb.Worksheets(CopyingSettings.item(COPYING_TOWORKSHEET_KEY)).Range(ToRangeAddress).ClearContents
If CopyingSettings.item(XL_SPECIAL_OPERATION_KEY) = -4142 Then
    twb.Worksheets(CopyingSettings.item(COPYING_TOWORKSHEET_KEY)).Range(ToRangeAddress).Value = ""
End If

fwb.Worksheets(CopyingSettings.item(COPYING_FROMWORKSHEET_KEY)).Range(FromRangeAddress).Copy

'for merged cells special operation must be "PasteValuesAndFormats", not just "PasteValues"
twb.Worksheets(CopyingSettings.item(COPYING_TOWORKSHEET_KEY)).Range(ToRangeAddress).PasteSpecial _
    Paste:=CopyingSettings.item(XL_PASTE_TYPE_KEY), _
    Operation:=CopyingSettings.item(XL_SPECIAL_OPERATION_KEY)

End Sub

Private Sub CopyFromRangeToCell(ByVal CopyingSettings As Collection, _
        ByVal FromBaseCellRange As Range, ByVal ToBaseCellRange As Range, _
        ByVal fwb As Workbook, ByVal twb As Workbook)

Dim FromRangeAddress As String, ToRangeAddress As String

FromRangeAddress = CommonHelpers.GetRangeStringFromOffsets(FromBaseCellRange, _
    CopyingSettings.item(COPYING_FROM_TOP_LEFT_CELL_COLUMN_OFFSET_KEY), CopyingSettings.item(COPYING_FROM_TOP_LEFT_CELL_ROW_OFFSET_KEY), _
    CopyingSettings.item(COPYING_FROM_RIGHT_BOTTOM_CELL_COLUMN_OFFSET_KEY), CopyingSettings.item(COPYING_FROM_RIGHT_BOTTOM_CELL_ROW_OFFSET_KEY))
                      
ToRangeAddress = CommonHelpers.GetRangeStringFromOffsets(ToBaseCellRange, _
    CopyingSettings.item(COPYING_TO_TOP_LEFT_CELL_COLUMN_OFFSET_KEY), CopyingSettings.item(COPYING_TO_TOP_LEFT_CELL_ROW_OFFSET_KEY), _
    CopyingSettings.item(COPYING_TO_RIGHT_BOTTOM_CELL_COLUMN_OFFSET_KEY), CopyingSettings.item(COPYING_TO_RIGHT_BOTTOM_CELL_ROW_OFFSET_KEY))

Dim Cell As Range
Dim CellsColor As Long

If CopyingSettings.item(COPYING_COLOR_KEY) = -1 Then
    '16777215 - white (default color)
    CellsColor = 16777215
Else
    CellsColor = CLng(CopyingSettings.item(COPYING_COLOR_KEY))
End If

'clear previous values
'twb.Worksheets(CopyingSettings.item(COPYING_TOWORKSHEET_KEY)).Range(ToRangeAddress).ClearContents
If CopyingSettings.item(XL_SPECIAL_OPERATION_KEY) = -4142 Then
    twb.Worksheets(CopyingSettings.item(COPYING_TOWORKSHEET_KEY)).Range(ToRangeAddress).Value = ""
End If


With fwb.Worksheets(CopyingSettings.item(COPYING_FROMWORKSHEET_KEY))
    For Each Cell In .Range(FromRangeAddress).Cells
        If Cell.Interior.Color = CellsColor Then
            Cell.Copy
            twb.Worksheets(CopyingSettings.item(COPYING_TOWORKSHEET_KEY)).Range(ToRangeAddress).PasteSpecial _
                Paste:=CopyingSettings.item(XL_PASTE_TYPE_KEY), _
                Operation:=CopyingSettings.item(XL_SPECIAL_OPERATION_KEY)
        End If
    Next
End With

End Sub

Private Function SetCopyingWorkbook(ByVal wbFullName As String, ByVal wbKey As String, _
    ByRef wb As Workbook, ByVal ReadOnlyFlag As Boolean) As Boolean

SetCopyingWorkbook = True

'if workbook name was not passed it is supposed to be ThisWorkbook
If Len(wbFullName) = 0 Then
    Set wb = ThisWorkbook
Else
    If DoesWorkbookExist(wbFullName) = True Then
        If StrComp(wbFullName, ThisWorkbook.FullName, vbTextCompare) <> 0 Then
            If IsWorkBookOpen(wbFullName) = False Then
                'open workbook
                Set wb = Workbooks.Open(FileName:=wbFullName, ReadOnly:=ReadOnlyFlag)
                Dim ExtractedFileName As String
                ExtractedFileName = CommonHelpers.ExtractFileNameFromPath(wbFullName)
                
                If Not OpenWorkbooksPaths Is Nothing And DoesCollectionContainKey(OpenWorkbooksPaths, ExtractedFileName) = False Then
                    Dim PathReadOnlyFlagPair As Pair
                    Set PathReadOnlyFlagPair = New Pair
                    PathReadOnlyFlagPair.First = ExtractedFileName
                    PathReadOnlyFlagPair.Second = Not ReadOnlyFlag
                    
                    OpenWorkbooksPaths.Add PathReadOnlyFlagPair, ExtractedFileName
                    
                    Set PathReadOnlyFlagPair = Nothing
                End If
            Else
                wbFullName = CommonHelpers.ExtractFileNameFromPath(wbFullName)
                Set wb = Workbooks(wbFullName)
            End If
        Else: Set wb = ThisWorkbook
        End If
    Else
        SetCopyingWorkbook = False
        Exit Function
    End If
End If

End Function

