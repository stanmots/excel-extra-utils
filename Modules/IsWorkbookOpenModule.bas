Attribute VB_Name = "IsWorkbookOpenModule"
Option Explicit
Option Private Module

Public Function IsWorkBookOpen(ByVal FileName As String) As Boolean

Dim fileNumber As Long, errCode As Long

Err.Clear
On Error Resume Next

'Returns an Integer value representing the next file number available for use by the FileOpen function
fileNumber = FreeFile()

'Open pathname For mode [Access access] [lock] As [#]filenumber [Len=reclength]
Open FileName For Input Lock Read As #fileNumber

'The sharp sign (#) is optional(some previous versions of the BASIC language required that)
Close #fileNumber

errCode = Err.Number

'error handler disabled, next error will not be handled
On Error GoTo 0

Select Case errCode

    ' File is NOT already open by another user
    Case 0:    IsWorkBookOpen = False
    
    ' Error number for "Permission Denied."
    Case 70:   IsWorkBookOpen = True
    
    'File NOT found error
    'it is better to exit a program with Error
    'Case 53, 75: MsgBox WORKBOOK_NOT_FOUND_ERROR_MSG, vbOKOnly, ERROR_TITLE
    
    ' Raise some other run-time error
    Case Else: Error errCode
End Select
    
End Function

Private Sub UnitTest_IsWorkBookOpen()

Debug.Print IsWorkBookOpen(ThisWorkbook.FullName)

End Sub
