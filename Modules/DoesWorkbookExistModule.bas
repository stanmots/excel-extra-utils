Attribute VB_Name = "DoesWorkbookExistModule"
Option Explicit
Option Private Module

Public Function DoesWorkbookExist(ByVal FullFileName As String) As Boolean
    'returns TRUE if the file exists
    DoesWorkbookExist = Len(Dir(FullFileName)) > 0
End Function

Private Sub UnitTest_DoesWorkbookExist()
  
Debug.Print DoesWorkbookExist(ThisWorkbook.FullName)

End Sub
