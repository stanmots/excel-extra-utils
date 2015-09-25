Attribute VB_Name = "DeleteSheetWithoutAlertsModule"
Option Explicit
Option Private Module

Public Sub DeleteSheetWithoutAlerts(ByVal sheetName As String, Optional ByVal wb As Workbook)

If DoesSheetExist(sheetName) = False Then
    Debug.Print "Cannot delete the sheet(" & sheetName & "), because it is not found."
    Exit Sub
End If

If wb Is Nothing Then
    Set wb = ThisWorkbook
End If

Application.DisplayAlerts = False
wb.sheets(sheetName).Delete
Application.DisplayAlerts = True

End Sub

Private Sub UnitTest_DeleteSheetWithoutAlerts()

DeleteSheetWithoutAlerts "SomeWorkSheetName"

End Sub
