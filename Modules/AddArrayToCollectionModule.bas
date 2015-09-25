Attribute VB_Name = "AddArrayToCollectionModule"
Option Explicit
Option Private Module

'collection cannot contain the same key
Public Function AddArrayToCollection(ByVal ToCollection As Collection, FromArray As Variant) As Boolean
    
    AddArrayToCollection = False
    
    If IsObject(FromArray) = True Then
        Debug.Print "Error! The variable "; FromArray; " in [AddArrayToCollection] function is an object!"
        Exit Function
    End If
   
    Const COLLECTION_CONTAINS_SAME_KEY_ERROR As String = "Error! The variable ToCollection in [AddArrayToCollection] function contains the same key."
    
    If IsArray(FromArray) = False Then
        If DoesCollectionContainKey(ToCollection, CStr(FromArray)) = False Then
            ToCollection.Add FromArray, CStr(FromArray)
        Else
            Debug.Print COLLECTION_CONTAINS_SAME_KEY_ERROR
            Exit Function
        End If
    Else
        Dim Element As Variant
        For Each Element In FromArray
            If DoesCollectionContainKey(ToCollection, CStr(Element)) = False Then
                ToCollection.Add Element, CStr(Element)
            Else
                Debug.Print COLLECTION_CONTAINS_SAME_KEY_ERROR
                Exit Function
            End If
        Next
    End If
    
    AddArrayToCollection = True
    
End Function

Private Sub UnitTest_AddArrayToCollection()
  
Dim col As New Collection
Dim TestArray() As Variant
TestArray = Array(23, 12, 123, 16, 57, 48, 78, 3, 149, 963, 28, 125, 37, 77, 77)

AddArrayToCollection col, TestArray

Dim TwoDimArray(1 To 1, 1 To 3) As Integer
TwoDimArray(1, 1) = 10
TwoDimArray(1, 2) = 11
TwoDimArray(1, 3) = 12

AddArrayToCollection col, TwoDimArray

PrintArrayToConsole col

Set col = Nothing

End Sub
