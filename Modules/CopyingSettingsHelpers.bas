Attribute VB_Name = "CopyingSettingsHelpers"
'*****************************
'* The CopyingSettingsHelpers Module
'*
'* Short description:
'*
'*  Contains the functions specific to the Copying Worksheets Vba program.
'*
'* Basic usage:
'*
'*  CopyingSettingsHelpers.FunctionNameHere(arg1, arg2...)
'*
'*****************************

Option Explicit
Option Private Module

'special operations combo-box constants
Private Const xlPasteSpecialOperationNone_IntValue As Integer = -4142
Private Const xlPasteSpecialOperationAdd_IntValue As Integer = 2
Private Const xlPasteSpecialOperationSubtract_IntValue As Integer = 3
Private Const xlPasteSpecialOperationMultiply_IntValue As Integer = 4
Private Const xlPasteSpecialOperationDivide_IntValue As Integer = 5

'paste types combo-box constants
Private Const xlPasteAll_IntValue As Integer = -4104
Private Const xlPasteAllExceptBorders_IntValue As Integer = 7
Private Const xlPasteAllMergingConditionalFormats_IntValue As Integer = 14
Private Const xlPasteAllUsingSourceTheme_IntValue As Integer = 13
Private Const xlPasteColumnWidths_IntValue As Integer = 8
Private Const xlPasteComments_IntValue As Integer = -4144
Private Const xlPasteFormats_IntValue As Integer = -4122
Private Const xlPasteFormulas_IntValue As Integer = -4123
Private Const xlPasteFormulasAndNumberFormats_IntValue As Integer = 11
Private Const xlPasteValidation_IntValue As Integer = 6
Private Const xlPasteValues_IntValue As Integer = -4163
Private Const xlPasteValuesAndNumberFormats_IntValue As Integer = 12

'special operations combo-box
Public Function GetXlPasteSpecialOperationNumber(ByVal xlPasteSpecialOperationType As String) As Integer
 
Select Case xlPasteSpecialOperationType
    Case "xlPasteSpecialOperationNone"
        GetXlPasteSpecialOperationNumber = xlPasteSpecialOperationNone_IntValue
    Case "xlPasteSpecialOperationAdd"
        GetXlPasteSpecialOperationNumber = xlPasteSpecialOperationAdd_IntValue
    Case "xlPasteSpecialOperationSubtract"
        GetXlPasteSpecialOperationNumber = xlPasteSpecialOperationSubtract_IntValue
    Case "xlPasteSpecialOperationMultiply"
        GetXlPasteSpecialOperationNumber = xlPasteSpecialOperationMultiply_IntValue
    Case "xlPasteSpecialOperationDivide"
        GetXlPasteSpecialOperationNumber = xlPasteSpecialOperationDivide_IntValue
    Case Else
        GetXlPasteSpecialOperationNumber = 0
End Select
    
End Function

Public Function GetXlPasteSpecialOperationString(ByVal xlPasteSpecialOperationType As Integer) As String
 
Select Case xlPasteSpecialOperationType
    Case -4142
        GetXlPasteSpecialOperationString = "xlPasteSpecialOperationNone"
    Case 2
        GetXlPasteSpecialOperationString = "xlPasteSpecialOperationAdd"
    Case 3
        GetXlPasteSpecialOperationString = "xlPasteSpecialOperationSubtract"
    Case 4
        GetXlPasteSpecialOperationString = "xlPasteSpecialOperationMultiply"
    Case 5
        GetXlPasteSpecialOperationString = "xlPasteSpecialOperationDivide"
    Case Else
        GetXlPasteSpecialOperationString = vbNullString
End Select
    
End Function

Public Function GetSpecialOperationsComboBoxItemSerialNumberFromString(ByVal item As String) As Integer
 
Select Case item
    Case "xlPasteSpecialOperationNone"
        GetSpecialOperationsComboBoxItemSerialNumberFromString = 0
    Case "xlPasteSpecialOperationAdd"
        GetSpecialOperationsComboBoxItemSerialNumberFromString = 1
    Case "xlPasteSpecialOperationSubtract"
        GetSpecialOperationsComboBoxItemSerialNumberFromString = 2
    Case "xlPasteSpecialOperationMultiply"
        GetSpecialOperationsComboBoxItemSerialNumberFromString = 3
    Case "xlPasteSpecialOperationDivide"
        GetSpecialOperationsComboBoxItemSerialNumberFromString = 4
    Case Else
        GetSpecialOperationsComboBoxItemSerialNumberFromString = -1
End Select
    
End Function

'paste types combo-box
Public Function GetXlPasteTypeNumber(ByVal xlPasteType As String) As Integer
 
Select Case xlPasteType
    Case "xlPasteAll"
        GetXlPasteTypeNumber = xlPasteAll_IntValue
    Case "xlPasteAllExceptBorders"
        GetXlPasteTypeNumber = xlPasteAllExceptBorders_IntValue
    Case "xlPasteAllMergingConditionalFormats"
        GetXlPasteTypeNumber = xlPasteAllMergingConditionalFormats_IntValue
    Case "xlPasteAllUsingSourceTheme"
        GetXlPasteTypeNumber = xlPasteAllUsingSourceTheme_IntValue
    Case "xlPasteColumnWidths"
        GetXlPasteTypeNumber = xlPasteColumnWidths_IntValue
    Case "xlPasteComments"
        GetXlPasteTypeNumber = xlPasteComments_IntValue
    Case "xlPasteFormats"
        GetXlPasteTypeNumber = xlPasteFormats_IntValue
    Case "xlPasteFormulas"
        GetXlPasteTypeNumber = xlPasteFormulas_IntValue
    Case "xlPasteFormulasAndNumberFormats"
        GetXlPasteTypeNumber = xlPasteFormulasAndNumberFormats_IntValue
    Case "xlPasteValidation"
        GetXlPasteTypeNumber = xlPasteValidation_IntValue
    Case "xlPasteValues"
        GetXlPasteTypeNumber = xlPasteValues_IntValue
    Case "xlPasteValuesAndNumberFormats"
        GetXlPasteTypeNumber = xlPasteValuesAndNumberFormats_IntValue
    Case Else
        GetXlPasteTypeNumber = 0
End Select
    
End Function

Public Function GetXlPasteTypeString(ByVal xlPasteType As Integer) As String
 
Select Case xlPasteType
    Case -4104
        GetXlPasteTypeString = "xlPasteAll"
    Case 7
        GetXlPasteTypeString = "xlPasteAllExceptBorders"
    Case 14
        GetXlPasteTypeString = "xlPasteAllMergingConditionalFormats"
    Case 13
        GetXlPasteTypeString = "xlPasteAllUsingSourceTheme"
    Case 8
        GetXlPasteTypeString = "xlPasteColumnWidths"
    Case -4144
        GetXlPasteTypeString = "xlPasteComments"
    Case -4122
        GetXlPasteTypeString = "xlPasteFormats"
    Case -4123
        GetXlPasteTypeString = "xlPasteFormulas"
    Case 11
        GetXlPasteTypeString = "xlPasteFormulasAndNumberFormats"
    Case 6
        GetXlPasteTypeString = "xlPasteValidation"
    Case -4163
        GetXlPasteTypeString = "xlPasteValues"
    Case 12
        GetXlPasteTypeString = "xlPasteValuesAndNumberFormats"
    Case Else
        GetXlPasteTypeString = vbNullString
End Select
    
End Function

Public Function GetPasteTypesComboBoxItemSerialNumberFromString(ByVal item As String) As Integer
 
Select Case item
    Case "xlPasteAll"
        GetPasteTypesComboBoxItemSerialNumberFromString = 0
    Case "xlPasteAllExceptBorders"
        GetPasteTypesComboBoxItemSerialNumberFromString = 1
    Case "xlPasteAllMergingConditionalFormats"
        GetPasteTypesComboBoxItemSerialNumberFromString = 2
    Case "xlPasteAllUsingSourceTheme"
        GetPasteTypesComboBoxItemSerialNumberFromString = 3
    Case "xlPasteColumnWidths"
        GetPasteTypesComboBoxItemSerialNumberFromString = 4
    Case "xlPasteComments"
        GetPasteTypesComboBoxItemSerialNumberFromString = 5
    Case "xlPasteFormats"
        GetPasteTypesComboBoxItemSerialNumberFromString = 6
    Case "xlPasteFormulas"
        GetPasteTypesComboBoxItemSerialNumberFromString = 7
    Case "xlPasteFormulasAndNumberFormats"
        GetPasteTypesComboBoxItemSerialNumberFromString = 8
    Case "xlPasteValidation"
        GetPasteTypesComboBoxItemSerialNumberFromString = 9
    Case "xlPasteValues"
        GetPasteTypesComboBoxItemSerialNumberFromString = 10
    Case "xlPasteValuesAndNumberFormats"
        GetPasteTypesComboBoxItemSerialNumberFromString = 11
    Case Else
        GetPasteTypesComboBoxItemSerialNumberFromString = -1
End Select
    
End Function

Public Function GetPasteParametersString(ByVal PasteTypeStr As String, SpecialOperationStr As String) As String

Dim NestedPasteType() As String
NestedPasteType = Split(PasteTypeStr, DELIMITER:="xlPaste")

Dim NestedPasteSpecialOperation() As String
NestedPasteSpecialOperation = Split(SpecialOperationStr, DELIMITER:="xlPasteSpecialOperation")

GetPasteParametersString = NestedPasteType(1) & PASTE_PARAMETERS_DELIMITER & NestedPasteSpecialOperation(1)

End Function

Public Function AreCopyingRangesValid(ByVal FromRangeStr As String, ByVal ToRangeStr As String) As CopyingRangesType

AreCopyingRangesValid = CopyingRangesType.INCORRECT_RANGES

If Len(FromRangeStr) = 0 Or Len(ToRangeStr) = 0 Then
    Exit Function
End If

Dim frg As Range, trg As Range
Set frg = Range(FromRangeStr)
Set trg = Range(ToRangeStr)

'there are two possible situations:
'1) 1st and 2nd ranges have the same amount of cells, so the source range will be copied to the destination range as is.
'2) 1st range has several cells, while 2nd range consists of just one cell. In this case the cells in the 1st range will be chosed according to the copying color
' and then they will be copied to the 2nd range according to the XL_SPECIAL_OPERATION_TYPE.
If frg.Count = trg.Count Then
    AreCopyingRangesValid = CopyingRangesType.SIMILAR_RANGES_TYPE
Else
    If frg.Count > 1 And trg.Count = 1 Then
        AreCopyingRangesValid = CopyingRangesType.FROM_RANGE_TO_CELL_TYPE
    End If
End If

End Function

'ConfigName must be without subsections
Public Function GetSavedCopyingSettings(ByVal ConfigName As String, ByVal CopyingConfigs As Collection, Optional ByVal IsNeededGlobalSettingsChecking As Boolean = True) As Collection

Set GetSavedCopyingSettings = New Collection

If CopyingConfigs Is Nothing Or Len(ConfigName) = 0 Then
    SetErrorFlagInCopyingSettingsCollection GetSavedCopyingSettings, INCORRECT_ARGS_ERROR_MSG & vbCrLf & ERROR_FUNCTION_NAME & "[GetSavedCopyingSettings]"
    Exit Function
End If

Dim i As Long
Dim SectionName As String
Dim GlobalSectionName As String
Dim CurrentSettingsCollection As Collection

For i = 1 To CopyingConfigs.Count

    GlobalSectionName = COPYING_SECTION_KEY & SECTION_DELIMITER & ConfigName
    SectionName = COPYING_SECTION_KEY & SECTION_DELIMITER & ConfigName & SECTION_DELIMITER & CopyingConfigs.item(i)
    Set CurrentSettingsCollection = New Collection
    
    '[1] required settings
    'from-worksheet name and from-column
    Dim FromWorksheetName As String
    
    Dim GlobalFWSIsUsed As Boolean
    GlobalFWSIsUsed = ExcelUtilsMainWindow.MainStorage.RetrieveValue(GlobalSectionName, GLOBAL_FWS_IS_USED_KEY, False)
    
    If GlobalFWSIsUsed = True Then
        FromWorksheetName = ExcelUtilsMainWindow.MainStorage.RetrieveValue(GlobalSectionName, GLOBAL_FROMWORKSHEET_NAME_KEY, "")
    Else
        FromWorksheetName = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, COPYING_FROMWORKSHEET_KEY, "")
    End If
    
    If FromWorksheetName = "" Then
        SetErrorFlagInCopyingSettingsCollection GetSavedCopyingSettings, WORKSHEETNAME_IS_NOT_SET_ERROR_MSG
        Exit Function
    End If
    
    CurrentSettingsCollection.Add FromWorksheetName, COPYING_FROMWORKSHEET_KEY
    
    Dim FromColumn As String
    FromColumn = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, COPYING_FROMCOLUMN_KEY, "")
    If FromColumn = "" Then
        SetErrorFlagInCopyingSettingsCollection GetSavedCopyingSettings, COLUMN_IS_NOT_SET_ERROR_MSG
        Exit Function
    End If
    
    CurrentSettingsCollection.Add FromColumn, COPYING_FROMCOLUMN_KEY
    
    'from-topleft-offsets
    Dim FromTopLeftCellRowOffset As Variant, FromTopLeftCellColumnOffset As Variant
    
    FromTopLeftCellRowOffset = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, COPYING_FROM_TOP_LEFT_CELL_ROW_OFFSET_KEY)
    FromTopLeftCellColumnOffset = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, COPYING_FROM_TOP_LEFT_CELL_COLUMN_OFFSET_KEY)
    
    If IsNull(FromTopLeftCellRowOffset) = True Or IsNull(FromTopLeftCellColumnOffset) = True Then
        SetErrorFlagInCopyingSettingsCollection GetSavedCopyingSettings, TL_OFFSETS_ARE_NOT_SET_ERROR_MSG
        Exit Function
    End If
    
    CurrentSettingsCollection.Add FromTopLeftCellRowOffset, COPYING_FROM_TOP_LEFT_CELL_ROW_OFFSET_KEY
    CurrentSettingsCollection.Add FromTopLeftCellColumnOffset, COPYING_FROM_TOP_LEFT_CELL_COLUMN_OFFSET_KEY
    
    'from base cell and range
    Dim FromBaseCell As String, FromRange As String
    FromBaseCell = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, COPYING_FROMBASECELL_KEY, "")
    FromRange = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, COPYING_FROMRANGE_KEY, "")
    
    If FromRange = "" Then
        SetErrorFlagInCopyingSettingsCollection GetSavedCopyingSettings, COPYING_RANGE_IS_NOT_SET_ERROR_MSG
        Exit Function
    End If
    
    CurrentSettingsCollection.Add FromRange, COPYING_FROMRANGE_KEY
        
    If FromBaseCell = "" Then
        SetErrorFlagInCopyingSettingsCollection GetSavedCopyingSettings, COPYING_BASECELL_IS_NOT_SET_ERROR_MSG
        Exit Function
    End If
    
    CurrentSettingsCollection.Add FromBaseCell, COPYING_FROMBASECELL_KEY
    
    'to-worksheet name and to-column
    Dim ToWorksheetName As String
    
    Dim GlobalTWSIsUsed As Boolean
    GlobalTWSIsUsed = ExcelUtilsMainWindow.MainStorage.RetrieveValue(GlobalSectionName, GLOBAL_TWS_IS_USED_KEY, False)
    
    If GlobalTWSIsUsed = True Then
        ToWorksheetName = ExcelUtilsMainWindow.MainStorage.RetrieveValue(GlobalSectionName, GLOBAL_TOWORKSHEET_NAME_KEY, "")
    Else
        ToWorksheetName = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, COPYING_TOWORKSHEET_KEY, "")
    End If
    
    If ToWorksheetName = "" Then
        SetErrorFlagInCopyingSettingsCollection GetSavedCopyingSettings, WORKSHEETNAME_IS_NOT_SET_ERROR_MSG
        Exit Function
    End If
    
    CurrentSettingsCollection.Add ToWorksheetName, COPYING_TOWORKSHEET_KEY
    
    Dim ToColumn As String
    ToColumn = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, COPYING_TOCOLUMN_KEY, "")
    If ToColumn = "" Then
        SetErrorFlagInCopyingSettingsCollection GetSavedCopyingSettings, COLUMN_IS_NOT_SET_ERROR_MSG
        Exit Function
    End If
    
    CurrentSettingsCollection.Add ToColumn, COPYING_TOCOLUMN_KEY
    
    'to base cell and range
    Dim ToBaseCell As String, ToRange As String
    
    ToBaseCell = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, COPYING_TOBASECELL_KEY, "")
    ToRange = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, COPYING_TORANGE_KEY, "")
    
    If ToRange = "" Then
        SetErrorFlagInCopyingSettingsCollection GetSavedCopyingSettings, COPYING_RANGE_IS_NOT_SET_ERROR_MSG
        Exit Function
    End If
    
    CurrentSettingsCollection.Add ToRange, COPYING_TORANGE_KEY
        
    If ToBaseCell = "" Then
        SetErrorFlagInCopyingSettingsCollection GetSavedCopyingSettings, COPYING_BASECELL_IS_NOT_SET_ERROR_MSG
        Exit Function
    End If
        
    CurrentSettingsCollection.Add ToBaseCell, COPYING_TOBASECELL_KEY
 
    'to-offsets
    Dim ToTopLeftCellRowOffset As Variant, ToTopLeftCellColumnOffset As Variant
    
    ToTopLeftCellRowOffset = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, COPYING_TO_TOP_LEFT_CELL_ROW_OFFSET_KEY)
    ToTopLeftCellColumnOffset = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, COPYING_TO_TOP_LEFT_CELL_COLUMN_OFFSET_KEY)
    
    If IsNull(ToTopLeftCellRowOffset) = True Or IsNull(ToTopLeftCellColumnOffset) = True Then
        SetErrorFlagInCopyingSettingsCollection GetSavedCopyingSettings, TL_OFFSETS_ARE_NOT_SET_ERROR_MSG
        Exit Function
    End If
    
    CurrentSettingsCollection.Add ToTopLeftCellRowOffset, COPYING_TO_TOP_LEFT_CELL_ROW_OFFSET_KEY
    CurrentSettingsCollection.Add ToTopLeftCellColumnOffset, COPYING_TO_TOP_LEFT_CELL_COLUMN_OFFSET_KEY
    
    '[2] optional settings
    'from-workbook name
    Dim FromWorkbookName As String
    
    Dim GlobalFWBIsUsed As Boolean
    GlobalFWBIsUsed = ExcelUtilsMainWindow.MainStorage.RetrieveValue(GlobalSectionName, GLOBAL_FWB_IS_USED_KEY, False)
    
    If GlobalFWBIsUsed = True Then
        
        FromWorkbookName = ExcelUtilsMainWindow.MainStorage.RetrieveValue(GlobalSectionName, GLOBAL_FROMWORKBOOK_NAME_KEY, "")
        
        If IsNeededGlobalSettingsChecking = True And FromWorkbookName = "" Then
            SetErrorFlagInCopyingSettingsCollection GetSavedCopyingSettings, COPYING_FROMWORKBOOK_IS_NOT_SET_ERROR_MSG
            Exit Function
        End If
    
    Else
        FromWorkbookName = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, COPYING_FROMWORKBOOK_KEY, "")
    End If
    
    CurrentSettingsCollection.Add FromWorkbookName, COPYING_FROMWORKBOOK_KEY
    
    'from-rightbottom-offsets
    Dim FromRightBottomCellRowOffset As Variant, FromRightBottomCellColumnOffset As Variant
    FromRightBottomCellRowOffset = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, COPYING_FROM_RIGHT_BOTTOM_CELL_ROW_OFFSET_KEY, "")
    FromRightBottomCellColumnOffset = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, COPYING_FROM_RIGHT_BOTTOM_CELL_COLUMN_OFFSET_KEY, "")
 
    CurrentSettingsCollection.Add FromRightBottomCellRowOffset, COPYING_FROM_RIGHT_BOTTOM_CELL_ROW_OFFSET_KEY
    CurrentSettingsCollection.Add FromRightBottomCellColumnOffset, COPYING_FROM_RIGHT_BOTTOM_CELL_COLUMN_OFFSET_KEY
    
    'to-workbook name
    Dim ToWorkbookName As String
    
    Dim GlobalTWBIsUsed As Boolean
    GlobalTWBIsUsed = ExcelUtilsMainWindow.MainStorage.RetrieveValue(GlobalSectionName, GLOBAL_TWB_IS_USED_KEY, False)
    
    If GlobalTWBIsUsed = True Then
        
        ToWorkbookName = ExcelUtilsMainWindow.MainStorage.RetrieveValue(GlobalSectionName, GLOBAL_TOWORKBOOK_NAME_KEY, "")
        
        If IsNeededGlobalSettingsChecking = True And ToWorkbookName = "" Then
            SetErrorFlagInCopyingSettingsCollection GetSavedCopyingSettings, COPYING_TOWORKBOOK_IS_NOT_SET_ERROR_MSG
            Exit Function
        End If
    
    Else
        ToWorkbookName = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, COPYING_TOWORKBOOK_KEY, "")
    End If
    
    CurrentSettingsCollection.Add ToWorkbookName, COPYING_TOWORKBOOK_KEY
    
    'to-rightbottom-offsets
    Dim ToRightBottomCellRowOffset As Variant, ToRightBottomCellColumnOffset As Variant
    ToRightBottomCellRowOffset = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, COPYING_TO_RIGHT_BOTTOM_CELL_ROW_OFFSET_KEY, "")
    ToRightBottomCellColumnOffset = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, COPYING_TO_RIGHT_BOTTOM_CELL_COLUMN_OFFSET_KEY, "")
    
    CurrentSettingsCollection.Add ToRightBottomCellRowOffset, COPYING_TO_RIGHT_BOTTOM_CELL_ROW_OFFSET_KEY
    CurrentSettingsCollection.Add ToRightBottomCellColumnOffset, COPYING_TO_RIGHT_BOTTOM_CELL_COLUMN_OFFSET_KEY
    
    '[3] common settings
    'get special operation
    Dim SpecialOperationType As Integer
    SpecialOperationType = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, XL_SPECIAL_OPERATION_KEY, -4142)
    CurrentSettingsCollection.Add SpecialOperationType, XL_SPECIAL_OPERATION_KEY
    
    'get paste type
    Dim PasteType As Integer
    PasteType = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, XL_PASTE_TYPE_KEY, -4104)
    CurrentSettingsCollection.Add PasteType, XL_PASTE_TYPE_KEY
    
    'color
    Dim CopyingColor As Long
    CopyingColor = ExcelUtilsMainWindow.MainStorage.RetrieveValue(SectionName, COPYING_COLOR_KEY, -1)
     
    CurrentSettingsCollection.Add CopyingColor, COPYING_COLOR_KEY
    
    GetSavedCopyingSettings.Add CurrentSettingsCollection
Next i

End Function

Private Sub SetErrorFlagInCopyingSettingsCollection(ByRef csCol As Collection, ByVal ErrorMsg As String)

Set csCol = New Collection

csCol.Add ErrorMsg, ERROR_FLAG_KEY

End Sub




