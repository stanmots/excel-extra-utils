VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbaSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************
'* The VbaSettings class
'*
'* Short description:
'*
'*  Provides persistent storage for the vba applications' settings.
'*  The data is stored in the special worksheet which is hidden. You can unhide it from the Format/List menu.
'*  The name of the worksheet will be taken from the argument you pass to the constructor.
'*
'* Basic usage:
'*
'*  You can create VbaSettings object using the method from the Factory Module:
'*
'*      Dim MainSettingsStorage As VbaSettings
'*      Set MainSettingsStorage = FactoryModule.CreateObjectOfTypeVBASettings(SETTINGS_WORKSHEET_NAME)
'*
'*  To store a value in the persistend storage use this function(all the basic types
'*  and one-dimensional arrays are supported):
'*
'*      MainSettingsStorage.StoreValue(SectionName, SettingsEntryName, Value)
'*
'*  To retrieve a value:
'*
'*      Dim value As Variant
'*      value = MainSettingsStorage.RetrieveValue(SectionName, SettingsEntryNam, DefaultValue)
'*
'*  Section also could be nested like this: "TOP_LEVEL_SECION/NEXT_LEVEL_SECTION"
'*
'*****************************

Option Explicit

'We need these marks for the proper namespace handling
Private Const START_SECTION_MARK As String = "Start Section" & SPLITTER
Private Const END_SECTION_MARK As String = "End Section" & SPLITTER

Private m_SettingsWorksheetName As String
 
Public Sub InitiateProperties(ByVal WorksheetName As String)

m_SettingsWorksheetName = WorksheetName

If DoesSheetExist(WorksheetName) = False Then
    CreateSettingsWorksheet
End If

End Sub

Public Property Get SettingsWorksheetName() As String

SettingsWorksheetName = m_SettingsWorksheetName
    
End Property

Private Sub CreateSettingsWorksheet()

AddHiddenSheet m_SettingsWorksheetName

With ThisWorkbook.Worksheets(m_SettingsWorksheetName)
    .Range(.Cells(1, 1), .Cells(1, 10)).Merge
    .Range("A1").Value = SETTINGS_HEADER
    .Range("A1").Font.Bold = True
    .Range("A1").HorizontalAlignment = xlCenter
End With

End Sub

Private Function DoesSectionExist(ByVal SectionName As String, Optional ByVal SourceRange As Range) As Range

If SourceRange Is Nothing Then
    Set SourceRange = ThisWorkbook.Worksheets(m_SettingsWorksheetName).UsedRange.Columns(1)
End If

Dim StartSectionCell As Range, EndSectionCell As Range
Dim NestedSections() As String
Dim i As Long

NestedSections = Split(SectionName, DELIMITER:=SECTION_DELIMITER)

Set DoesSectionExist = Nothing
Const CurrentSectionColumn As String = "A"

For i = LBound(NestedSections) To UBound(NestedSections)
    If Len(CStr(NestedSections(i))) > 0 Then
        Set StartSectionCell = SourceRange.Find(What:=START_SECTION_MARK & CStr(NestedSections(i)), _
                    After:=SourceRange.Cells(SourceRange.Cells.Count), _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    SearchOrder:=xlByColumns, _
                    SearchDirection:=xlNext, _
                    MatchByte:=False, _
                    MatchCase:=True, _
                    SearchFormat:=False)
                    
        If Not StartSectionCell Is Nothing Then
            Dim NewSourceRange As Range
            Dim StartRangeIndex As Long, EndRangeIndex As Long

            StartRangeIndex = StartSectionCell.Cells(1).Row - SourceRange.Cells(1).Row + 1
            EndRangeIndex = SourceRange.End(xlDown).Row - SourceRange.Cells(1).Row + 1
            
            Set NewSourceRange = SourceRange.Range(CurrentSectionColumn & CStr(StartRangeIndex) & ":" & _
                CurrentSectionColumn & CStr(EndRangeIndex))
    
            Set EndSectionCell = NewSourceRange.Find(What:=END_SECTION_MARK & CStr(NestedSections(i)), _
                        After:=NewSourceRange.Cells(NewSourceRange.Cells.Count), _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByColumns, _
                        SearchDirection:=xlNext, _
                        MatchByte:=False, _
                        MatchCase:=True, _
                        SearchFormat:=False)
            If Not EndSectionCell Is Nothing Then
                Set SourceRange = SourceRange.Range(CurrentSectionColumn & CStr(StartRangeIndex) & ":" & _
                    CurrentSectionColumn & CStr(EndSectionCell.Cells(1).Row - SourceRange.Cells(1).Row + 1))
            Else: Exit Function
            End If
        Else: Exit Function
        End If
    End If
Next i

Set DoesSectionExist = SourceRange

End Function

Private Function DoesSettingsEntryExist(ByVal SettingsEntryName As String, ByVal SectionRange As Range) As Range

If SectionRange Is Nothing Then
    Set DoesSettingsEntryExist = Nothing
    Exit Function
End If

Set DoesSettingsEntryExist = SectionRange.Find(What:=SettingsEntryName, _
                After:=SectionRange.Cells(SectionRange.Cells.Count), _
                LookIn:=xlValues, _
                LookAt:=xlWhole, _
                SearchOrder:=xlByColumns, _
                SearchDirection:=xlNext, _
                MatchByte:=False, _
                MatchCase:=True, _
                SearchFormat:=False)
                    
End Function

Private Function AddSettingsEntry(ByVal SettingsEntryName As String, ByVal SectionRange As Range) As Range

Dim SettingsEntryIndex As Long
SettingsEntryIndex = SectionRange.Count
SectionRange.Cells(SectionRange.Count).EntireRow.Insert

SectionRange.Cells(SettingsEntryIndex).Value = SettingsEntryName

Set AddSettingsEntry = SectionRange.Cells(SettingsEntryIndex)

End Function

Private Function GetSettingsEntryValuesRange(ByVal SettingsEntryCell As Range) As Range

Set GetSettingsEntryValuesRange = Nothing

Dim LastUsedCellIndex As Long
With ThisWorkbook.Worksheets(m_SettingsWorksheetName)
    LastUsedCellIndex = .Cells(SettingsEntryCell.Row, .Columns.Count).End(xlToLeft).Column
    
    If LastUsedCellIndex > SettingsEntryCell.Column Then
        Set GetSettingsEntryValuesRange = .Range(.Cells(SettingsEntryCell.Row, SettingsEntryCell.Column + 1), _
                                        .Cells(SettingsEntryCell.Row, LastUsedCellIndex))
    End If
End With

End Function

Private Sub ClearSettingsEntryValues(ByVal SettingsEntryCell As Range)

Dim SettingsEntryValuesRange As Range
Set SettingsEntryValuesRange = GetSettingsEntryValuesRange(SettingsEntryCell)

If Not SettingsEntryValuesRange Is Nothing Then
    SettingsEntryValuesRange.Clear
End If

End Sub

Private Function WriteDataInPersistentStorage(ByRef Data As Variant, ByVal SettingsEntryCell As Range) As Boolean

WriteDataInPersistentStorage = False

If IsObject(Data) = True Then
    Debug.Print "Error! The variable "; Data; " in WriteDataInPersistentStorage function is an object!"
    Exit Function
End If

Dim SettingsEntryValuesRange As Range
Set SettingsEntryValuesRange = SettingsEntryCell.Offset(0, 1)

If IsArray(Data) = True Then
    If GetArrayDimensionsNumber(Data) = 1 Then
        Dim ColumnsCount As Long
        ColumnsCount = UBound(Data) - LBound(Data) + 1
        
        Set SettingsEntryValuesRange = SettingsEntryValuesRange.Resize(1, ColumnsCount)
    Else
        Debug.Print "Error! Uncorrect Array has been passed to the WriteDataInPersistentStorage function."
        Exit Function
    End If
End If

SettingsEntryValuesRange.Value = Data
SettingsEntryValuesRange.HorizontalAlignment = xlCenter
ThisWorkbook.Worksheets(m_SettingsWorksheetName).UsedRange.Columns.AutoFit

WriteDataInPersistentStorage = True

End Function

Private Function ReadFromPersistentStorage(ByVal SettingsEntryCell As Range) As Variant

ReadFromPersistentStorage = Null

If SettingsEntryCell Is Nothing Then Exit Function

Dim SettingsEntryValuesRange As Range
Set SettingsEntryValuesRange = GetSettingsEntryValuesRange(SettingsEntryCell)

If Not SettingsEntryValuesRange Is Nothing Then
    ReadFromPersistentStorage = SettingsEntryValuesRange
End If

End Function

Public Sub StoreValue(ByVal SectionName As String, ByVal SettingsEntryName As String, ByRef Value As Variant)

Dim SectionRange As Range, SettingsEntryCell As Range

Set SectionRange = DoesSectionExist(SectionName)
If SectionRange Is Nothing Then
    Set SectionRange = AddSection(SectionName)
End If
    
Set SettingsEntryCell = DoesSettingsEntryExist(SettingsEntryName, SectionRange)
If SettingsEntryCell Is Nothing Then
    Set SettingsEntryCell = AddSettingsEntry(SettingsEntryName, SectionRange)
End If

'clear previous values
ClearSettingsEntryValues SettingsEntryCell

WriteDataInPersistentStorage Value, SettingsEntryCell

End Sub

Public Function RetrieveValue(ByVal SectionName As String, ByVal SettingsEntryName As String, _
                                Optional ByRef DefaultValue As Variant) As Variant

RetrieveValue = Null

Dim SectionRange As Range, SettingsEntryCell As Range

Set SectionRange = DoesSectionExist(SectionName)
If SectionRange Is Nothing Then
    If IsMissing(DefaultValue) = False Then
        Set SectionRange = AddSection(SectionName)
    Else: Exit Function
    End If
End If

Set SettingsEntryCell = DoesSettingsEntryExist(SettingsEntryName, SectionRange)
If SettingsEntryCell Is Nothing Then
    If IsMissing(DefaultValue) = False Then
        Set SettingsEntryCell = AddSettingsEntry(SettingsEntryName, SectionRange)
    Else: Exit Function
    End If
End If

RetrieveValue = ReadFromPersistentStorage(SettingsEntryCell)

If IsNull(RetrieveValue) = True And IsMissing(DefaultValue) = False Then
    WriteDataInPersistentStorage DefaultValue, SettingsEntryCell
    RetrieveValue = DefaultValue
End If

End Function

'SectionName can contain subsections, delimited by "/"
Public Function AddSection(ByVal SectionName As String) As Range

Dim SectionRange As Range

'if a section already exists return its range
Set SectionRange = DoesSectionExist(SectionName)
If Not SectionRange Is Nothing Then
    Set AddSection = SectionRange
    Exit Function
End If

Set AddSection = Nothing

Dim CellIndexForInsert As Long
With ThisWorkbook.Worksheets(m_SettingsWorksheetName).UsedRange.Columns(1)
    CellIndexForInsert = .Cells(.Cells.Count).Row
End With

Dim NestedSections() As String
Dim i As Long

NestedSections = Split(SectionName, DELIMITER:=SECTION_DELIMITER)

Dim IsNeededAddSection As Boolean
IsNeededAddSection = False

For i = LBound(NestedSections) To UBound(NestedSections)
    If IsNeededAddSection = False Then
        Set SectionRange = DoesSectionExist(CStr(NestedSections(i)), SectionRange)
        If SectionRange Is Nothing Then
            IsNeededAddSection = True
        Else
            CellIndexForInsert = SectionRange.Cells(SectionRange.Count - 1).Row
        End If
    End If
    
    If IsNeededAddSection = True Then
        With ThisWorkbook.Worksheets(m_SettingsWorksheetName).UsedRange.Columns(1)
        
            'insert 2 rows
            .Cells(CellIndexForInsert + 1).EntireRow.Resize(2).Insert
            
            'select range as the first two cells of the inserted rows
            Set SectionRange = .Range(.Cells(CellIndexForInsert + 1), .Cells(CellIndexForInsert + 2))
            
            'set the section marks
            SectionRange.Cells(1).Value = START_SECTION_MARK & CStr(NestedSections(i))
            SectionRange.Cells(SectionRange.Cells.Count).Value = END_SECTION_MARK & CStr(NestedSections(i))
    
            'tuning the appearance
            SectionRange.Font.Bold = True
            SectionRange.HorizontalAlignment = xlCenter
            
            CellIndexForInsert = CellIndexForInsert + 1
        End With
    End If
Next i

ThisWorkbook.Worksheets(m_SettingsWorksheetName).UsedRange.Columns(1).AutoFit

Set AddSection = SectionRange

End Function

Public Sub DeleteSection(ByVal SectionName As String)

Dim SectionRange As Range

Set SectionRange = DoesSectionExist(SectionName)
If Not SectionRange Is Nothing Then
    SectionRange.EntireRow.Delete Shift:=xlShiftUp
End If

End Sub

Public Sub DeleteSettingsEntry(ByVal SectionName As String, ByVal SettingsEntryName As String)

Dim SectionRange As Range, SettingsEntryCell As Range

Set SectionRange = DoesSectionExist(SectionName)
If SectionRange Is Nothing Then Exit Sub
    
Set SettingsEntryCell = DoesSettingsEntryExist(SettingsEntryName, SectionRange)
If Not SettingsEntryCell Is Nothing Then
    SettingsEntryCell.EntireRow.Delete Shift:=xlShiftUp
End If

End Sub

Public Function RenameSection(ByVal FromSectionName As String, ByVal ToSectionName As String) As Boolean

RenameSection = True

Dim SectionRange As Range

Set SectionRange = DoesSectionExist(FromSectionName)
If Not SectionRange Is Nothing Then
    SectionRange.Cells(1).Value = START_SECTION_MARK & ToSectionName
    SectionRange.Cells(SectionRange.Cells.Count).Value = END_SECTION_MARK & ToSectionName
Else
    Debug.Print "The required section cannot be renamed!"
    RenameSection = False
    Exit Function
End If

End Function

'SectionName can contain nested subsections
'if SectionName does not exist an empty collection will be returned
Public Function GetAllSubSections(ByVal SectionName As String) As Collection

Set GetAllSubSections = New Collection

Dim SectionRange As Range

Set SectionRange = DoesSectionExist(SectionName)
If SectionRange Is Nothing Then Exit Function

Set SectionRange = SectionRange.Range("A2:A" & _
    CStr(SectionRange.Cells(SectionRange.Cells.Count).Row - SectionRange.Cells(1).Row + 1))
    
If SectionRange.Count = 1 Then Exit Function

Dim StartSectionCell As Range, EndSectionCell As Range
Dim IsSearchFinished As Boolean
IsSearchFinished = False

Do While IsSearchFinished = False
    Set StartSectionCell = SectionRange.Find(What:=START_SECTION_MARK, _
                After:=SectionRange.Cells(SectionRange.Cells.Count), _
                LookIn:=xlValues, _
                LookAt:=xlPart, _
                SearchOrder:=xlByColumns, _
                SearchDirection:=xlNext, _
                MatchByte:=False, _
                MatchCase:=True, _
                SearchFormat:=False)
                
    If Not StartSectionCell Is Nothing Then
        Dim NewSectionRange As Range
        Dim StartRangeIndex As Long
        Dim SectionSubName As String
        StartRangeIndex = StartSectionCell.Cells(1).Row - SectionRange.Cells(1).Row + 1
        SectionSubName = Split(StartSectionCell.Value, SPLITTER)(1)
        
        Set NewSectionRange = SectionRange.Range("A" & CStr(StartRangeIndex) & ":A" & _
            CStr(SectionRange.Cells(SectionRange.Cells.Count).Row - SectionRange.Cells(1).Row + 1))

        Set EndSectionCell = NewSectionRange.Find(What:=END_SECTION_MARK & CStr(SectionSubName), _
                    After:=NewSectionRange.Cells(NewSectionRange.Cells.Count), _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    SearchOrder:=xlByColumns, _
                    SearchDirection:=xlNext, _
                    MatchByte:=False, _
                    MatchCase:=True, _
                    SearchFormat:=False)
        If Not EndSectionCell Is Nothing Then
            GetAllSubSections.Add SectionSubName
            Set SectionRange = SectionRange.Range("A" & CStr(EndSectionCell.Cells(1).Row - SectionRange.Cells(1).Row + 1) & _
             ":A" & CStr(SectionRange.Cells(SectionRange.Cells.Count).Row - SectionRange.Cells(1).Row + 1))
        Else: IsSearchFinished = True
        End If
    Else: IsSearchFinished = True
    End If
Loop

End Function
