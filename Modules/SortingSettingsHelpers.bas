Attribute VB_Name = "SortingSettingsHelpers"
'*****************************
'* The SortingSettingsHelpers Module
'*
'* Short description:
'*
'*  Contains the functions specific to the Sorting Worksheets Vba program.
'*
'* Basic usage:
'*
'*  SortingSettingsHelpers.FunctionNameHere(arg1, arg2...)
'*
'*****************************

Option Explicit
Option Private Module

'return false if there are no settings for the current worksheet
Public Function RetrieveCurrentWorksheetSettings(ByVal WorksheetName As String, Optional ByRef SortingColumn As String = "COLUMN_MISSING", _
    Optional ByVal SortingOffsets As Collection = Nothing, Optional ByRef SortingRange As String = "RANGE_MISSING", _
    Optional ByRef BaseCell As String = "BASECELL_MISSING", _
    Optional ByRef SerialCell As String = "SERIALCELL_MISSING") As Boolean

Dim TestVar As Variant
RetrieveCurrentWorksheetSettings = False

Dim SectionName As String
SectionName = SORTING_SECTION_KEY & SECTION_DELIMITER & WorksheetName

With ExcelUtilsMainWindow.MainStorage
    
    'logical optional settings we can retrieve just for checking the existence (without storing)
    TestVar = .RetrieveValue(SectionName, SORTING_COLUMN_KEY)
    If IsNull(TestVar) = False Then
        If SortingColumn <> "COLUMN_MISSING" Then SortingColumn = CStr(TestVar)
    Else: Exit Function
    End If
    
    TestVar = .RetrieveValue(SectionName, SORTING_TOP_LEFT_CELL_ROW_OFFSET_KEY)
    If IsNull(TestVar) = False Then
        If Not SortingOffsets Is Nothing Then SortingOffsets.Add CLng(TestVar), SORTING_TOP_LEFT_CELL_ROW_OFFSET_KEY
    Else: Exit Function
    End If
    
    TestVar = .RetrieveValue(SectionName, SORTING_TOP_LEFT_CELL_COLUMN_OFFSET_KEY)
    If IsNull(TestVar) = False Then
        If Not SortingOffsets Is Nothing Then SortingOffsets.Add CLng(TestVar), SORTING_TOP_LEFT_CELL_COLUMN_OFFSET_KEY
    Else: Exit Function
    End If
    
    TestVar = .RetrieveValue(SectionName, SORTING_RIGHT_BOTTOM_CELL_ROW_OFFSET_KEY)
    If IsNull(TestVar) = False Then
        If Not SortingOffsets Is Nothing Then SortingOffsets.Add CLng(TestVar), SORTING_RIGHT_BOTTOM_CELL_ROW_OFFSET_KEY
    Else: Exit Function
    End If
    
    TestVar = .RetrieveValue(SectionName, SORTING_RIGHT_BOTTOM_CELL_COLUMN_OFFSET_KEY)
    If IsNull(TestVar) = False Then
        If Not SortingOffsets Is Nothing Then SortingOffsets.Add CLng(TestVar), SORTING_RIGHT_BOTTOM_CELL_COLUMN_OFFSET_KEY
    Else: Exit Function
    End If
    
    'real optional settings (do not require exit function)
    If Not SortingOffsets Is Nothing Then
    
        TestVar = .RetrieveValue(SectionName, SORTING_SERIAL_CELL_ROW_OFFSET_KEY)
        If IsNull(TestVar) = False Then
            SortingOffsets.Add CLng(TestVar), SORTING_SERIAL_CELL_ROW_OFFSET_KEY
        Else
            SortingOffsets.Add "", SORTING_SERIAL_CELL_ROW_OFFSET_KEY
        End If
        
        TestVar = .RetrieveValue(SectionName, SORTING_SERIAL_CELL_COLUMN_OFFSET_KEY)
        If IsNull(TestVar) = False Then
            SortingOffsets.Add CLng(TestVar), SORTING_SERIAL_CELL_COLUMN_OFFSET_KEY
        Else
            SortingOffsets.Add "", SORTING_SERIAL_CELL_COLUMN_OFFSET_KEY
        End If
        
    End If

          
    If SortingRange <> "RANGE_MISSING" Then
    
        TestVar = .RetrieveValue(SectionName, SORTING_RANGE_KEY)
        If IsNull(TestVar) = False Then
            SortingRange = CStr(TestVar)
        End If
    
    End If
    
    If BaseCell <> "BASECELL_MISSING" Then
    
        TestVar = .RetrieveValue(SectionName, SORTING_BASECELL_KEY)
        If IsNull(TestVar) = False Then
            BaseCell = CStr(TestVar)
        End If
    
    End If
    
    If SerialCell <> "SERIALCELL_MISSING" Then

    TestVar = .RetrieveValue(SectionName, SORTING_SERIAL_CELL_KEY)
    If IsNull(TestVar) = False Then
        SerialCell = CStr(TestVar)
    End If

    End If
    
End With

'if all the settings were found
RetrieveCurrentWorksheetSettings = True

End Function

Public Function DoSelectedWorksheetsHaveValidSettings(ByVal aListBox As MSForms.ListBox) As Boolean

DoSelectedWorksheetsHaveValidSettings = True

Dim i As Long

For i = 0 To aListBox.ListCount - 1
    If aListBox.Selected(i) = True Then
        If SortingSettingsHelpers.RetrieveCurrentWorksheetSettings(aListBox.List(i, 0)) = False Then
            DoSelectedWorksheetsHaveValidSettings = False
            Exit Function
        End If
    End If
Next i

End Function



