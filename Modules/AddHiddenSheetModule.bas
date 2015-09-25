Attribute VB_Name = "AddHiddenSheetModule"
Option Explicit
Option Private Module

Public Sub AddHiddenSheet(ByVal sheetName As String, Optional ByVal wb As Workbook)

If DoesSheetExist(sheetName) = True Then
    Debug.Print "Cannot add the sheet " & sheetName & ", because it already exists."
    Exit Sub
End If

If wb Is Nothing Then
    Set wb = ThisWorkbook
End If

Dim OldSheet As Worksheet
Set OldSheet = ActiveSheet

'check if excel can hide the sheet (if current visible sheets number is > 0)
Dim ws As Worksheet, isAbleHideSheet As Boolean
For Each ws In wb.Worksheets
    If ws.Visible = True Then
        isAbleHideSheet = True
        Exit For
    End If
Next

'the sheet will be added after the last one
wb.sheets.Add(After:=wb.sheets(wb.sheets.Count)).Name = sheetName
OldSheet.Activate

If isAbleHideSheet = True Then
    wb.sheets(wb.sheets.Count).Visible = xlSheetHidden
End If

End Sub

Private Sub UnitTest_AddSheetWithoutActivating()

AddHiddenSheet SETTINGS_WORKSHEET_NAME
'DeleteSheetWithoutAlerts "TestSheetName"

End Sub
