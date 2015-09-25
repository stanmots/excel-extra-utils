Attribute VB_Name = "DoesSheetExistModule"
Option Explicit
Option Private Module

Public Function DoesSheetExist(ByVal sheetName As String, Optional ByVal wb As Workbook) As Boolean

Dim ws As Worksheet

If wb Is Nothing Then
    Set wb = ThisWorkbook
End If

Err.Clear
On Error Resume Next

'The SHEETS object is a parent object for: Worksheets, Chart Sheets
Set ws = wb.sheets(sheetName)

On Error GoTo 0

DoesSheetExist = Not ws Is Nothing
     
End Function

Private Sub UnitTest_DoesSheetExist()
  
Debug.Print DoesSheetExist(ThisWorkbook.sheets(1).Name, ThisWorkbook)

End Sub


