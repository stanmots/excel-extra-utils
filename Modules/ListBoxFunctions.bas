Attribute VB_Name = "ListBoxFunctions"
Option Explicit
Option Private Module

Public Sub UpdateListBoxWithAllSheetsExceptExcluded(ByVal ListBoxForUpdate As MSForms.ListBox, _
                Optional ByVal ExcludedNames As Collection, Optional ByVal wb As Workbook)

If wb Is Nothing Then
    Set wb = ThisWorkbook
End If

ListBoxForUpdate.Clear

'fill the listbox with the worksheets names
Dim ws As Worksheet
For Each ws In wb.Worksheets
    If DoesCollectionContainKey(ExcludedNames, ws.Name) = False Then
        ListBoxForUpdate.AddItem ws.Name
    End If
Next

End Sub

'can be used with collections too
Public Sub UpdateListboxFromArray(ByVal ListBoxForUpdate As MSForms.ListBox, _
                 ByRef LBEntriesArray As Variant, Optional ByVal wb As Workbook)

If wb Is Nothing Then
    Set wb = ThisWorkbook
End If

ListBoxForUpdate.Clear

Const LISTBOX_ELEMENT_IS_NOT_STRING_ERROR As String = "There was an error during updating the listbox. The element in the collection is not a string!"

If IsArray(LBEntriesArray) = False And TypeName(LBEntriesArray) <> "Collection" Then
    If TypeName(LBEntriesArray) = "String" Then
        ListBoxForUpdate.AddItem LBEntriesArray
    Else: Debug.Print LISTBOX_ELEMENT_IS_NOT_STRING_ERROR
    End If
    Exit Sub
End If

Dim i As Variant
For Each i In LBEntriesArray
    If TypeName(i) = "String" Then
        ListBoxForUpdate.AddItem i
    Else: Debug.Print LISTBOX_ELEMENT_IS_NOT_STRING_ERROR
    End If
Next

End Sub

Public Sub SetSelectionModeOfListBox(ByVal aListBox As MSForms.ListBox, ByVal modeFlag As Boolean)

Dim i As Long
For i = 0 To aListBox.ListCount - 1
    aListBox.Selected(i) = modeFlag
Next i

End Sub

'returns true if at least one selected item was copied
Public Function CopyAllSelectedItems(ByVal FromListBox As MSForms.ListBox, ByVal ToListBox As MSForms.ListBox) As Boolean

Dim i As Long
CopyAllSelectedItems = False

For i = 0 To FromListBox.ListCount - 1
    If FromListBox.Selected(i) = True Then
        If CopyAllSelectedItems <> True Then CopyAllSelectedItems = True
  
        Dim SortingSettingsListBoxRowIndex As Variant
        
        ToListBox.AddItem
        SortingSettingsListBoxRowIndex = ToListBox.ListCount - 1
        
        'copy selected item
        Dim ColumnIndex As Long
        ColumnIndex = 0
        Do While ColumnIndex < FromListBox.ColumnCount And ColumnIndex < ToListBox.ColumnCount
            ToListBox.List(SortingSettingsListBoxRowIndex, ColumnIndex) = FromListBox.List(i, ColumnIndex)
            ColumnIndex = ColumnIndex + 1
        Loop
        
    End If
Next i

End Function

Public Function IsListBoxHasSelectedItems(ByVal aListBox As MSForms.ListBox) As Boolean

IsListBoxHasSelectedItems = False

Dim i As Long
For i = 0 To aListBox.ListCount - 1
    If aListBox.Selected(i) = True Then
        IsListBoxHasSelectedItems = True
        Exit Function
    End If
Next i

End Function

Public Function SelectedCount(ByVal aListBox As MSForms.ListBox) As Long

SelectedCount = 0

Dim i As Long
For i = 0 To aListBox.ListCount - 1
    If aListBox.Selected(i) = True Then
        SelectedCount = SelectedCount + 1
    End If
Next i

End Function

Public Function IsListBoxHasItem(ByVal ItemName As String, ByVal aListBox As MSForms.ListBox) As Boolean

IsListBoxHasItem = False

Dim i As Long, j As Long
For i = 0 To aListBox.ListCount - 1
    For j = 0 To aListBox.ColumnCount - 1
        If aListBox.List(i, j) = ItemName Then
            IsListBoxHasItem = True
            Exit Function
        End If
    Next j
Next i

End Function

Public Function RenameItem(ByVal OldItemName As String, ByVal NewItemName As String, ByVal aListBox As MSForms.ListBox) As Boolean

RenameItem = False

Dim i As Long, j As Long
For i = 0 To aListBox.ListCount - 1
    For j = 0 To aListBox.ColumnCount - 1
        If aListBox.List(i, j) = OldItemName Then
            aListBox.List(i, j) = NewItemName
            RenameItem = True
            Exit Function
        End If
    Next j
Next i

End Function

'get the first items in a row if a listbox is multi-column
Public Function GetSelectedItems(ByVal aListBox As MSForms.ListBox) As Collection

Set GetSelectedItems = New Collection

Dim i As Long
For i = 0 To aListBox.ListCount - 1
    If aListBox.Selected(i) = True Then
        GetSelectedItems.Add aListBox.List(i, 0)
    End If
Next i

End Function
