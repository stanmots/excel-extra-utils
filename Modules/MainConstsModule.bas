Attribute VB_Name = "MainConstsModule"
Option Explicit
Option Private Module

'Constants
Public Const VBA_PROGRAMM_VERSION As String = "v1.0"
Public Const SPLITTER As String = ": "
Public Const PASTE_PARAMETERS_DELIMITER As String = ", "
Public Const SECTION_DELIMITER As String = "/"
Public Const EXCEL_FILE_FILTER As String = "Excel files (*.xls; *.xlsx), *.xls; *.xlsx"

'enums
Public Enum SettingsType
    FROM_WORKSHEET_TYPE = 0
    TO_WORKSHEET_TYPE = 1
End Enum

Public Enum CopyingRangesType
    INCORRECT_RANGES = 10
    SIMILAR_RANGES_TYPE = 11
    FROM_RANGE_TO_CELL_TYPE = 12
End Enum

Public Enum GlobalSetting
    IS_USED = 20
    IS_NOT_USED = 21
End Enum
