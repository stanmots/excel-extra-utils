Attribute VB_Name = "MainKeysModule"
Option Explicit
Option Private Module

Public Const ERROR_FLAG_KEY As String = "ErrorFlag"

'sorting keys
Public Const SORTING_SECTION_KEY As String = "SortingSection"
Public Const SORTING_COLUMN_KEY As String = "Sorting Column"

Public Const EXCLUDED_WORKSHEETS_KEY As String = "Excluded WorkSheets"

Public Const SORTING_TOP_LEFT_CELL_ROW_OFFSET_KEY As String = "SortingTopLeftCellRowOffset"
Public Const SORTING_RIGHT_BOTTOM_CELL_ROW_OFFSET_KEY As String = "SortingRightBottomCellRowOffset"
Public Const SORTING_TOP_LEFT_CELL_COLUMN_OFFSET_KEY As String = "SortingTopLeftCellColumnOffset"
Public Const SORTING_RIGHT_BOTTOM_CELL_COLUMN_OFFSET_KEY As String = "SortingRightBottomCellColumnOffset"
Public Const SORTING_SERIAL_CELL_ROW_OFFSET_KEY As String = "SortingSerialCellRowOffset"
Public Const SORTING_SERIAL_CELL_COLUMN_OFFSET_KEY As String = "SortingSerialCellColumnOffset"

Public Const SORTING_RANGE_KEY As String = "SortingRange"
Public Const SORTING_BASECELL_KEY As String = "SortingBaseCell"
Public Const SORTING_SERIAL_CELL_KEY As String = "SortingSerialCell"

'copying keys
Public Const COPYING_SECTION_KEY As String = "CopyingSection"

Public Const COPYING_FROM_TOP_LEFT_CELL_ROW_OFFSET_KEY As String = "CopyingFromTopLeftCellRowOffset"
Public Const COPYING_FROM_RIGHT_BOTTOM_CELL_ROW_OFFSET_KEY As String = "CopyingFromRightBottomCellRowOffset"
Public Const COPYING_FROM_TOP_LEFT_CELL_COLUMN_OFFSET_KEY As String = "CopyingFromTopLeftCellColumnOffset"
Public Const COPYING_FROM_RIGHT_BOTTOM_CELL_COLUMN_OFFSET_KEY As String = "CopyingFromRightBottomCellColumnOffset"

Public Const COPYING_TO_TOP_LEFT_CELL_ROW_OFFSET_KEY As String = "CopyingToTopLeftCellRowOffset"
Public Const COPYING_TO_RIGHT_BOTTOM_CELL_ROW_OFFSET_KEY As String = "CopyingToRightBottomCellRowOffset"
Public Const COPYING_TO_TOP_LEFT_CELL_COLUMN_OFFSET_KEY As String = "CopyingToTopLeftCellColumnOffset"
Public Const COPYING_TO_RIGHT_BOTTOM_CELL_COLUMN_OFFSET_KEY As String = "CopyingToRightBottomCellColumnOffset"

Public Const COPYING_FROMWORKSHEET_KEY As String = "CopyingFromWorksheet"
Public Const COPYING_FROMWORKBOOK_KEY As String = "CopyingFromWorkbook"
Public Const COPYING_FROMCOLUMN_KEY As String = "CopyingFromColumn"
Public Const COPYING_FROMBASECELL_KEY As String = "CopyingFromBaseCell"
Public Const COPYING_FROMRANGE_KEY As String = "CopyingFromRange"

Public Const COPYING_TOWORKSHEET_KEY As String = "CopyingToWorksheet"
Public Const COPYING_TOWORKBOOK_KEY As String = "CopyingToWorkbook"
Public Const COPYING_TOCOLUMN_KEY As String = "CopyingToColumn"
Public Const COPYING_TOBASECELL_KEY As String = "CopyingToBaseCell"
Public Const COPYING_TORANGE_KEY As String = "CopyingToRange"

Public Const XL_SPECIAL_OPERATION_KEY As String = "xlPasteSpecialOperation"
Public Const XL_PASTE_TYPE_KEY As String = "xlPasteType"
Public Const COPYING_COLOR_KEY As String = "CopyingColorKey"

'global copying settings
Public Const GLOBAL_FWB_IS_USED_KEY As String = "GlobalFromWorkbookNameIsUsed"
Public Const GLOBAL_FWS_IS_USED_KEY As String = "GlobalFromWorksheetNameIsUsed"
Public Const GLOBAL_TWB_IS_USED_KEY As String = "GlobalToWorkbookNameIsUsed"
Public Const GLOBAL_TWS_IS_USED_KEY As String = "GlobalToWorksheetNameIsUsed"

Public Const GLOBAL_FWB_IS_REMOVED_AFTER_COPYING_KEY As String = "GlobalFWBIsRemovedAfterCopying"
Public Const GLOBAL_TWB_IS_REMOVED_AFTER_COPYING_KEY As String = "GlobalTWBIsRemovedAfterCopying"

Public Const GLOBAL_FROMWORKBOOK_NAME_KEY As String = "GlobalFromWorkbookName"
Public Const GLOBAL_FROMWORKSHEET_NAME_KEY As String = "GlobalFromWorksheetName"
Public Const GLOBAL_TOWORKBOOK_NAME_KEY As String = "GlobalToWorkbookName"
Public Const GLOBAL_TOWORKSHEET_NAME_KEY As String = "GlobalToWorksheetName"

'coloring section
Public Const COLORING_SECTION_KEY As String = "ColoringSection"
Public Const COLORING_COLUMN_KEY As String = "ColoringColumn"
Public Const COLORING_BASECELL_KEY As String = "ColoringBaseCell"
Public Const COLORING_BASERANGE_KEY As String = "ColoringBaseRange"
Public Const COLORING_SOUGHTFORRANGE_KEY As String = "ColoringSoughtForRange"
Public Const COLORING_BASECOLOR_KEY As String = "ColoringBaseColor"

Public Const COLORING_SOUGHTFORRANGE_TOP_LEFT_CELL_ROW_OFFSET_KEY As String = "ColoringSoughtForRangeTLCRowOffset"
Public Const COLORING_SOUGHTFORRANGE_TOP_LEFT_CELL_COLUMN_OFFSET_KEY As String = "ColoringSoughtForRangeTLCColumnOffset"
Public Const COLORING_SOUGHTFORRANGE_RIGHT_BOTTOM_CELL_ROW_OFFSET_KEY As String = "ColoringSoughtForRangeRBCRowOffset"
Public Const COLORING_SOUGHTFORRANGE_RIGHT_BOTTOM_CELL_COLUMN_OFFSET_KEY As String = "ColoringSoughtForRangeRBCColumnOffset"

