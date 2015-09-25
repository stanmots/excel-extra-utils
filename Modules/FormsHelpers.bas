Attribute VB_Name = "FormsHelpers"
Option Explicit
Option Private Module

'it is used only for forms that can ADD and EDIT the copying settings
Public Sub PrepareAndShowCopyingForm(ByVal Form As Object, ByVal ParentForm As MSForms.UserForm, _
    ByVal IsNeededAddingFlag As Boolean, ByVal ParentFormStateFlag As Boolean)
 
Form.IsNeededAdding = IsNeededAddingFlag
FormsHelpers.ChangeStateOfAllControlsOnForm ParentForm, ParentFormStateFlag
If Form.Visible = False Then Form.Show
    
End Sub

Public Sub ChangeStateOfAllControlsOnForm(ByVal Form As MSForms.UserForm, ByVal StateFlag As Boolean)

Dim CurrentControl As Control

For Each CurrentControl In Form.Controls
    CurrentControl.Enabled = StateFlag
Next CurrentControl

End Sub

Public Function IsUserFormLoaded(ByVal UFName As String) As Boolean
    Dim UForm As Object
    IsUserFormLoaded = False
    For Each UForm In VBA.UserForms
        If UForm.Name = UFName Then
            IsUserFormLoaded = True
            Exit For
        End If
    Next
End Function
