Attribute VB_Name = "ColoringSettingsHelpers"
Option Explicit
Option Private Module

Public Function RetrieveCurrentWorksheetSettings(ByVal WorksheetName As String, _
    Optional ByVal ColoringOffsets As Collection = Nothing, _
    Optional ByRef ColoringColumn As String = "COLORINGCOLUMN_MISSING", _
    Optional ByRef BaseRange As String = "BASERANGE_MISSING", _
    Optional ByRef BaseCell As String = "BASECELL_MISSING", _
    Optional ByRef SoughtForRange As String = "SOUGHTFORRANGE_MISSING", _
    Optional ByRef Color As Long = -1) As Boolean

Dim TestVar As Variant
RetrieveCurrentWorksheetSettings = False

Dim SectionName As String
SectionName = COLORING_SECTION_KEY & SECTION_DELIMITER & WorksheetName

With ExcelUtilsMainWindow.MainStorage
    
    TestVar = .RetrieveValue(SectionName, COLORING_SOUGHTFORRANGE_TOP_LEFT_CELL_ROW_OFFSET_KEY)
    If IsNull(TestVar) = False Then
        If Not ColoringOffsets Is Nothing Then ColoringOffsets.Add CLng(TestVar), COLORING_SOUGHTFORRANGE_TOP_LEFT_CELL_ROW_OFFSET_KEY
    Else: Exit Function
    End If
    
    TestVar = .RetrieveValue(SectionName, COLORING_SOUGHTFORRANGE_TOP_LEFT_CELL_COLUMN_OFFSET_KEY)
    If IsNull(TestVar) = False Then
        If Not ColoringOffsets Is Nothing Then ColoringOffsets.Add CLng(TestVar), COLORING_SOUGHTFORRANGE_TOP_LEFT_CELL_COLUMN_OFFSET_KEY
    Else: Exit Function
    End If
    
    TestVar = .RetrieveValue(SectionName, COLORING_SOUGHTFORRANGE_RIGHT_BOTTOM_CELL_ROW_OFFSET_KEY)
    If IsNull(TestVar) = False Then
        If Not ColoringOffsets Is Nothing Then ColoringOffsets.Add CLng(TestVar), COLORING_SOUGHTFORRANGE_RIGHT_BOTTOM_CELL_ROW_OFFSET_KEY
    Else: Exit Function
    End If
    
    TestVar = .RetrieveValue(SectionName, COLORING_SOUGHTFORRANGE_RIGHT_BOTTOM_CELL_COLUMN_OFFSET_KEY)
    If IsNull(TestVar) = False Then
        If Not ColoringOffsets Is Nothing Then ColoringOffsets.Add CLng(TestVar), COLORING_SOUGHTFORRANGE_RIGHT_BOTTOM_CELL_COLUMN_OFFSET_KEY
    Else: Exit Function
    End If
    
    TestVar = .RetrieveValue(SectionName, COLORING_COLUMN_KEY)
    If IsNull(TestVar) = False Then
       If ColoringColumn <> "COLORINGCOLUMN_MISSING" Then ColoringColumn = CStr(TestVar)
    Else: Exit Function
    End If
    
    TestVar = .RetrieveValue(SectionName, COLORING_BASECELL_KEY)
    If IsNull(TestVar) = False Then
        If BaseCell <> "BASECELL_MISSING" Then BaseCell = CStr(TestVar)
    Else: Exit Function
    End If
    
    TestVar = .RetrieveValue(SectionName, COLORING_BASERANGE_KEY)
    If IsNull(TestVar) = False Then
        If BaseRange <> "BASERANGE_MISSING" Then BaseRange = CStr(TestVar)
    Else: Exit Function
    End If
    
    TestVar = .RetrieveValue(SectionName, COLORING_SOUGHTFORRANGE_KEY)
    If IsNull(TestVar) = False Then
       If SoughtForRange <> "SOUGHTFORRANGE_MISSING" Then SoughtForRange = CStr(TestVar)
    Else: Exit Function
    End If

    If Color <> -1 Then

        TestVar = .RetrieveValue(SectionName, COLORING_BASECOLOR_KEY)
        If IsNull(TestVar) = False Then
            Color = CLng(TestVar)
        Else: Color = 0
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
        If ColoringSettingsHelpers.RetrieveCurrentWorksheetSettings(aListBox.List(i, 0)) = False Then
            DoSelectedWorksheetsHaveValidSettings = False
            Exit Function
        End If
    End If
Next i

End Function
