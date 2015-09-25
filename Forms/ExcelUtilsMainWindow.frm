VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExcelUtilsMainWindow 
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   OleObjectBlob   =   "ExcelUtilsMainWindow.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExcelUtilsMainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private m_MainSettingsStorage As VbaSettings

Private Sub UserForm_Initialize()

SetMainWindowCaptions
Set m_MainSettingsStorage = FactoryModule.CreateObjectOfTypeVbaSettings(SETTINGS_WORKSHEET_NAME)

'restore settings of the:
'[1] sorting page
Dim StoredSortingExcludedWorkSheets As Variant, ExcludedSortingWorkSheetsCollection As New Collection

'an array with default excluded worksheets
'TODO: Create a UserForm by which users can edit excluded worksheets
Dim ExcludedSortingWorksheetsArray() As Variant
ExcludedSortingWorksheetsArray = Array(SETTINGS_WORKSHEET_NAME)

StoredSortingExcludedWorkSheets = m_MainSettingsStorage.RetrieveValue(SORTING_SECTION_KEY, EXCLUDED_WORKSHEETS_KEY, ExcludedSortingWorksheetsArray)
AddArrayToCollection ExcludedSortingWorkSheetsCollection, StoredSortingExcludedWorkSheets
ListBoxFunctions.UpdateListBoxWithAllSheetsExceptExcluded Me.MainMultiPageObject.Pages("SortingPage").WorksheetsNamesListBox, ExcludedSortingWorkSheetsCollection

Set ExcludedSortingWorkSheetsCollection = Nothing

'[2] copying page
Dim StoredCopyingConfigs As Collection

Set StoredCopyingConfigs = m_MainSettingsStorage.GetAllSubSections(COPYING_SECTION_KEY)
If StoredCopyingConfigs.Count > 0 Then
    ListBoxFunctions.UpdateListboxFromArray Me.MainMultiPageObject.Pages("CopyingPage").CopyingConfigsListBox, StoredCopyingConfigs
End If

'TODO: implement new form for editing persistent storage
'[3] coloring page
Dim StoredColoringExcludedWorkSheets As Variant, ExcludedColoringWorkSheetsCollection As New Collection

'an array with default excluded worksheets
'TODO: Create a UserForm by which users can edit excluded worksheets
Dim ExcludedColoringWorksheetsArray() As Variant
ExcludedColoringWorksheetsArray = Array(SETTINGS_WORKSHEET_NAME)

StoredColoringExcludedWorkSheets = m_MainSettingsStorage.RetrieveValue(COLORING_SECTION_KEY, EXCLUDED_WORKSHEETS_KEY, ExcludedColoringWorksheetsArray)
AddArrayToCollection ExcludedColoringWorkSheetsCollection, StoredColoringExcludedWorkSheets
ListBoxFunctions.UpdateListBoxWithAllSheetsExceptExcluded Me.MainMultiPageObject.Pages("ColoringPage").ColoringWorksheetsListBox, ExcludedColoringWorkSheetsCollection

Set ExcludedColoringWorkSheetsCollection = Nothing

End Sub

Private Sub SetMainWindowCaptions()

Me.Caption = EXCEL_UTILS_MAIN_WINDOW_TITLE
Me.SelectAllButton.Caption = SELECTALL_BUTTON_TITLE
Me.UnselectAllButton.Caption = UNSELECTALL_BUTTON_TITLE
Me.CloseExcelUtilsMainWindowButton.Caption = CANCEL_BUTTON_TITLE

'labels
Me.MainMultiPageObject.Pages("SortingPage").WorksheetsNamesListBoxLabel.Caption = WORKSHEETS_LIST_LABEL
Me.MainMultiPageObject.Pages("CopyingPage").CopyingConfigsLabel.Caption = COPYING_CONFIGS_DESCRIPTION_LABEL
Me.MainMultiPageObject.Pages("ColoringPage").ColoringWorksheetsLabel.Caption = WORKSHEETS_LIST_LABEL

'pages
Me.MainMultiPageObject.Pages("SortingPage").Caption = SORTING_PAGE_TITLE
Me.MainMultiPageObject.Pages("CopyingPage").Caption = COPYING_PAGE_TITLE
Me.MainMultiPageObject.Pages("ColoringPage").Caption = COLORING_PAGE_TITLE

'buttons
Me.MainMultiPageObject.Pages("SortingPage").SortingSettingsButton.Caption = SORTING_SETTINGS_BUTTON_TITLE
Me.MainMultiPageObject.Pages("SortingPage").StartSortingButton.Caption = START_SORTING_BUTTON_TITLE
Me.MainMultiPageObject.Pages("CopyingPage").StartCopyingButton.Caption = START_COPYING_BUTTON_TITLE
Me.MainMultiPageObject.Pages("CopyingPage").AddCopyingConfigButton.Caption = ADD_BUTTON_TITLE
Me.MainMultiPageObject.Pages("CopyingPage").DeleteCopyingConfigButton.Caption = DELETE_BUTTON_TITLE
Me.MainMultiPageObject.Pages("CopyingPage").EditCopyingConfigButton.Caption = EDIT_BUTTON_TITLE
Me.MainMultiPageObject.Pages("ColoringPage").ColoringSettingsButton.Caption = SETTINGS_BUTTON_TITLE
Me.MainMultiPageObject.Pages("ColoringPage").StartColoringButton.Caption = START_COLORING_BUTTON_TITLE

End Sub

Private Sub StartSortingButton_Click()

With Me.MainMultiPageObject.Pages("SortingPage")

    If ListBoxFunctions.IsListBoxHasSelectedItems(.WorksheetsNamesListBox) = False Then
        MsgBox NO_SELECTED_ITEMS_ERROR_MSG, vbOKOnly, ERROR_TITLE
        Exit Sub
    End If
    
    If SortingSettingsHelpers.DoSelectedWorksheetsHaveValidSettings(.WorksheetsNamesListBox) = True Then
        
        Dim SortingColumn As String
        Dim SortingOffsets As Collection
        Dim RowIndex As Long
        
        'set Excel performance properties before sorting
        Application.ScreenUpdating = False
        
        'start sorting
        FormsHelpers.ChangeStateOfAllControlsOnForm Me, False
        
        If ProgressBarForm.Visible = False Then
            ProgressBarForm.Show
        End If
        
        ProgressBarForm.AddMessageToDetailsBox SORTING_STARTED_MSG
        
        For RowIndex = 0 To .WorksheetsNamesListBox.ListCount - 1
            If .WorksheetsNamesListBox.Selected(RowIndex) = True Then
                
                ProgressBarForm.ResetProgress
                ProgressBarForm.SetMainLabelText CURRENT_SORTING_WORKSHEET_NAME_MSG & CStr(.WorksheetsNamesListBox.List(RowIndex, 0))
                
                Set SortingOffsets = New Collection
                If SortingSettingsHelpers.RetrieveCurrentWorksheetSettings(.WorksheetsNamesListBox.List(RowIndex, 0), SortingColumn, SortingOffsets) = True Then
                    SortRangesByNumberInColumn SortingColumn, SortingOffsets, CStr(.WorksheetsNamesListBox.List(RowIndex, 0))
                Else
                    ProgressBarForm.AddMessageToDetailsBox ERROR_TITLE & CANNOT_FIND_WORKSHEET_SETTINGS_ERROR_MSG
                    Exit For
                End If
                Set SortingOffsets = Nothing
                
            End If
        Next RowIndex
        
        ProgressBarForm.AddMessageToDetailsBox SORTING_FINISHED_MSG
        ProgressBarForm.FinishProgress
        
        FormsHelpers.ChangeStateOfAllControlsOnForm Me, True
        
        'restore Excel performance properties
        Application.ScreenUpdating = True
        
    Else
        MsgBox INCORRECT_SORTING_SETTINGS_ERROR_MSG, vbOKOnly, ERROR_TITLE
    End If
    
End With

End Sub

'you must select some configs in CopyingConfigsListBox before proceeding further
'each config must have correct settings stored in a persistent storage (you set these settings in CopyingSettingsForm)
Private Sub StartCopyingButton_Click()

With Me.MainMultiPageObject.Pages("CopyingPage")
    If ListBoxFunctions.IsListBoxHasSelectedItems(.CopyingConfigsListBox) = False Then
        MsgBox NO_SELECTED_ITEMS_ERROR_MSG, vbOKOnly, ERROR_TITLE
        Exit Sub
    End If
    
    Dim SelectedConfigName As String
    Dim SelectedConfigsInListBox As Collection
    Set SelectedConfigsInListBox = ListBoxFunctions.GetSelectedItems(Me.CopyingConfigsListBox)
    
    Dim CurrentConfig As Variant
    Dim SubSectionsOfCurrentConfig As Collection, StoredSettings As Collection
    
    'set initial properties of the CopyRangesModule
    Set CopyRangesModule.OpenWorkbooksPaths = New Collection
    Set CopyRangesModule.FromNumbersCache = New Collection
    Set CopyRangesModule.ToNumbersCache = New Collection
    
    'set Excel performance properties before copying
    Application.ScreenUpdating = False
    
    'start copying configs
    FormsHelpers.ChangeStateOfAllControlsOnForm Me, False
    
    If ProgressBarForm.Visible = False Then
        ProgressBarForm.Show
    End If
    
    ProgressBarForm.AddMessageToDetailsBox COPYING_STARTED_MSG
    
    For Each CurrentConfig In SelectedConfigsInListBox
    
        ProgressBarForm.ResetProgress
        ProgressBarForm.SetMainLabelText CURRENT_COPYING_CONFIG_NAME_MSG & CStr(CurrentConfig)
        
        Set SubSectionsOfCurrentConfig = m_MainSettingsStorage.GetAllSubSections(COPYING_SECTION_KEY & SECTION_DELIMITER & CurrentConfig)
        
        If SubSectionsOfCurrentConfig.Count = 0 Then
            ProgressBarForm.AddMessageToDetailsBox ERROR_TITLE & EMPTY_SETTINGS_ERROR_MSG
            Exit For
        End If
        
        Set StoredSettings = CopyingSettingsHelpers.GetSavedCopyingSettings(CurrentConfig, SubSectionsOfCurrentConfig)
        
        Set SubSectionsOfCurrentConfig = Nothing
        
        If DoesCollectionContainKey(StoredSettings, ERROR_FLAG_KEY) = False Then
        
            Dim CurrentSettings As Collection
            
            'update progress for each config
            ProgressBarForm.SetLoopsParameters 100, StoredSettings.Count
            
            For Each CurrentSettings In StoredSettings
                CopyRangesBetweenWorksheets CurrentSettings
                ProgressBarForm.IncreaseProgressInsideLoop
            Next
            
            Set CurrentSettings = Nothing
            
            'remove current global settings if needed
            Dim GlobalFWBIsUsed As Boolean
            GlobalFWBIsUsed = m_MainSettingsStorage.RetrieveValue(COPYING_SECTION_KEY & SECTION_DELIMITER & CurrentConfig, GLOBAL_FWB_IS_USED_KEY, False)
    
            If GlobalFWBIsUsed = True Then
                Dim GlobalFWBIsRemoved As Boolean
                GlobalFWBIsRemoved = m_MainSettingsStorage.RetrieveValue(COPYING_SECTION_KEY & SECTION_DELIMITER & CurrentConfig, GLOBAL_FWB_IS_REMOVED_AFTER_COPYING_KEY, False)
                If GlobalFWBIsRemoved = True Then
                    m_MainSettingsStorage.StoreValue COPYING_SECTION_KEY & SECTION_DELIMITER & CurrentConfig, GLOBAL_FROMWORKBOOK_NAME_KEY, ""
                End If
            End If
            
            Dim GlobalTWBIsUsed As Boolean
            GlobalTWBIsUsed = m_MainSettingsStorage.RetrieveValue(COPYING_SECTION_KEY & SECTION_DELIMITER & CurrentConfig, GLOBAL_TWB_IS_USED_KEY, False)
    
            If GlobalTWBIsUsed = True Then
                Dim GlobalTWBIsRemoved As Boolean
                GlobalTWBIsRemoved = m_MainSettingsStorage.RetrieveValue(COPYING_SECTION_KEY & SECTION_DELIMITER & CurrentConfig, GLOBAL_TWB_IS_REMOVED_AFTER_COPYING_KEY, False)
                If GlobalTWBIsRemoved = True Then
                    m_MainSettingsStorage.StoreValue COPYING_SECTION_KEY & SECTION_DELIMITER & CurrentConfig, GLOBAL_TOWORKBOOK_NAME_KEY, ""
                End If
            End If
                    
        Else
            ProgressBarForm.AddMessageToDetailsBox ERROR_TITLE & CANNOT_RESTORE_SETTINGS_ERROR_MSG & vbCrLf & _
                ERROR_DETAILS & StoredSettings.item(ERROR_FLAG_KEY)
            Exit For
        End If
    
        Set StoredSettings = Nothing
    Next
    
    'end copying configs
    FormsHelpers.ChangeStateOfAllControlsOnForm Me, True
    ProgressBarForm.AddMessageToDetailsBox COPYING_FINISHED_MSG
    ProgressBarForm.FinishProgress

    'clean-up
    Set SelectedConfigsInListBox = Nothing
    Set CopyRangesModule.FromNumbersCache = Nothing
    Set CopyRangesModule.ToNumbersCache = Nothing
    
    'close workbooks that were open by CopyRangesBetweenWorksheets procedure
    Dim p As Pair
    For Each p In CopyRangesModule.OpenWorkbooksPaths
        Workbooks(p.First).Close SaveChanges:=p.Second
    Next
    
    Set CopyRangesModule.OpenWorkbooksPaths = Nothing
    
    'restore Excel performance properties
    Application.ScreenUpdating = True
    
End With

End Sub

Private Sub StartColoringButton_Click()

With Me.MainMultiPageObject.Pages("ColoringPage")
    If ListBoxFunctions.IsListBoxHasSelectedItems(.ColoringWorksheetsListBox) = False Then
        MsgBox NO_SELECTED_ITEMS_ERROR_MSG, vbOKOnly, ERROR_TITLE
        Exit Sub
    End If
    
    If ColoringSettingsHelpers.DoSelectedWorksheetsHaveValidSettings(.ColoringWorksheetsListBox) = True Then
        
        Dim ColoringColumn As String, BaseRange As String
        Dim ColoringOffsets As Collection
        Dim RowIndex As Long, Color As Long
        
        'set Excel performance properties before sorting
        Application.ScreenUpdating = False
        
        'start sorting
        FormsHelpers.ChangeStateOfAllControlsOnForm Me, False
        
        If ProgressBarForm.Visible = False Then
            ProgressBarForm.Show
        End If
        
        ProgressBarForm.AddMessageToDetailsBox COLORING_STARTED_MSG
        
        For RowIndex = 0 To .ColoringWorksheetsListBox.ListCount - 1
            If .ColoringWorksheetsListBox.Selected(RowIndex) = True Then
                
                ProgressBarForm.ResetProgress
                ProgressBarForm.SetMainLabelText CURRENT_WORKSHEET_NAME_MSG & CStr(.ColoringWorksheetsListBox.List(RowIndex, 0))
                
                Set ColoringOffsets = New Collection
                If ColoringSettingsHelpers.RetrieveCurrentWorksheetSettings(.ColoringWorksheetsListBox.List(RowIndex, 0), ColoringOffsets, ColoringColumn, _
                    BaseRange:=BaseRange, Color:=Color) = True Then
                    
                    ColorizeRanges ColoringOffsets, ColoringColumn, BaseRange, CStr(.ColoringWorksheetsListBox.List(RowIndex, 0)), Color
                
                Else
                    ProgressBarForm.AddMessageToDetailsBox ERROR_TITLE & CANNOT_FIND_WORKSHEET_SETTINGS_ERROR_MSG
                    Exit For
                End If
                Set ColoringOffsets = Nothing
                
            End If
        Next RowIndex
        
        ProgressBarForm.AddMessageToDetailsBox COLORING_FINISHED_MSG
        ProgressBarForm.FinishProgress
        
        FormsHelpers.ChangeStateOfAllControlsOnForm Me, True
        
        'restore Excel performance properties
        Application.ScreenUpdating = True
        
    Else
        MsgBox CANNOT_RESTORE_SETTINGS_ERROR_MSG, vbOKOnly, ERROR_TITLE
    End If
    
End With

End Sub

Private Sub ColoringSettingsButton_Click()

If ColoringSettingsForm.Visible = False Then
    
    Dim SelectedFoundFlag As Boolean
    SelectedFoundFlag = ListBoxFunctions.CopyAllSelectedItems(Me.MainMultiPageObject.Pages("ColoringPage").ColoringWorksheetsListBox, ColoringSettingsForm.ColoringSettingsListBox)
 
    If SelectedFoundFlag = True Then
        ColoringSettingsForm.RetrieveWorksheetsSettings
        FormsHelpers.ChangeStateOfAllControlsOnForm ExcelUtilsMainWindow, False
        ColoringSettingsForm.Show
    Else
        MsgBox NO_SELECTED_ITEMS_ERROR_MSG, vbOKOnly, ERROR_TITLE
        Exit Sub
    End If
    
End If

End Sub

Private Sub AddCopyingConfigButton_Click()

FormsHelpers.PrepareAndShowCopyingForm CopyingSettingsForm, ExcelUtilsMainWindow, _
        True, False
        
End Sub

Private Sub EditCopyingConfigButton_Click()

If ListBoxFunctions.SelectedCount(Me.CopyingConfigsListBox) > 1 Then
    MsgBox TOO_MUCH_SELECTED_ITEMS_ERROR_MSG, vbOKOnly, ERROR_TITLE
    Exit Sub
Else
    If ListBoxFunctions.IsListBoxHasSelectedItems(Me.CopyingConfigsListBox) = False Then
        MsgBox NO_SELECTED_ITEMS_ERROR_MSG, vbOKOnly, ERROR_TITLE
        Exit Sub
    End If
End If

Dim SelectedConfigNameInListBox As String
Dim SelectedConfigsInListBox As Collection
Set SelectedConfigsInListBox = ListBoxFunctions.GetSelectedItems(Me.CopyingConfigsListBox)
SelectedConfigNameInListBox = SelectedConfigsInListBox.item(1)

Set SelectedConfigsInListBox = Nothing

If UpdateCopyingSettingsListBox(SelectedConfigNameInListBox) = False Then
    MsgBox CANNOT_RESTORE_SETTINGS_ERROR_MSG, vbOKOnly, ERROR_TITLE
    Exit Sub
End If

CopyingSettingsForm.CurrentStoredConfigName = SelectedConfigNameInListBox
FormsHelpers.PrepareAndShowCopyingForm CopyingSettingsForm, ExcelUtilsMainWindow, _
        False, False
    
End Sub

'SelectedConfigName must be without subsections
Private Function UpdateCopyingSettingsListBox(ByVal SelectedConfigName As String) As Boolean

UpdateCopyingSettingsListBox = False

Dim SubSectionsOfCurrentConfig As Collection, StoredSettings As Collection
Set SubSectionsOfCurrentConfig = m_MainSettingsStorage.GetAllSubSections(COPYING_SECTION_KEY & SECTION_DELIMITER & SelectedConfigName)

Set StoredSettings = CopyingSettingsHelpers.GetSavedCopyingSettings(SelectedConfigName, SubSectionsOfCurrentConfig, False)

Set SubSectionsOfCurrentConfig = Nothing

If DoesCollectionContainKey(StoredSettings, ERROR_FLAG_KEY) = False Then

    UpdateCopyingSettingsListBox = True
    
    CopyingSettingsForm.InputCopyingConfigNameTextBox.text = SelectedConfigName
    
    'update copying list box
    'StoredSettings is a collection of collections with configs
    Dim c As Collection
    For Each c In StoredSettings
        'fill the first row
        CopyingSettingsForm.CopyingSettingsListBox.AddItem
        
        Dim rCount As Long
        rCount = CopyingSettingsForm.CopyingSettingsListBox.ListCount - 1
        CopyingSettingsForm.CopyingSettingsListBox.List(rCount, 0) = CLng(rCount / 2 + 1)
        CopyingSettingsForm.CopyingSettingsListBox.List(rCount, 1) = COPYING_DESTINATION_FROMTYPE
        CopyingSettingsForm.CopyingSettingsListBox.List(rCount, 2) = c.item(COPYING_FROMCOLUMN_KEY)
    
        Dim FromOffsetsStr As String
        
        FromOffsetsStr = CommonHelpers.GetFormattedStringFromOffsets(c.item(COPYING_FROM_TOP_LEFT_CELL_COLUMN_OFFSET_KEY), c.item(COPYING_FROM_TOP_LEFT_CELL_ROW_OFFSET_KEY), _
        c.item(COPYING_FROM_RIGHT_BOTTOM_CELL_COLUMN_OFFSET_KEY), c.item(COPYING_FROM_RIGHT_BOTTOM_CELL_ROW_OFFSET_KEY))
                    
        CopyingSettingsForm.CopyingSettingsListBox.List(rCount, 3) = c.item(COPYING_FROMBASECELL_KEY) & ";" & c.item(COPYING_FROMRANGE_KEY) & FromOffsetsStr
        CopyingSettingsForm.CopyingSettingsListBox.List(rCount, 4) = c.item(COPYING_FROMWORKSHEET_KEY)
        CopyingSettingsForm.CopyingSettingsListBox.List(rCount, 5) = CommonHelpers.ExtractFileNameFromPath(c.item(COPYING_FROMWORKBOOK_KEY))
        
        'recover comboboxes settings
        Dim PasteParameters As String
        PasteParameters = CopyingSettingsHelpers.GetPasteParametersString(CopyingSettingsHelpers.GetXlPasteTypeString(c.item(XL_PASTE_TYPE_KEY)), _
            CopyingSettingsHelpers.GetXlPasteSpecialOperationString(c.item(XL_SPECIAL_OPERATION_KEY)))
        
        CopyingSettingsForm.CopyingSettingsListBox.List(rCount, 6) = PasteParameters
        
        If c.item(COPYING_COLOR_KEY) <> -1 Then
            CopyingSettingsForm.CopyingSettingsListBox.List(rCount, 7) = c.item(COPYING_COLOR_KEY)
        End If
     
       'fill the second row
        CopyingSettingsForm.CopyingSettingsListBox.AddItem
        
        rCount = rCount + 1
        CopyingSettingsForm.CopyingSettingsListBox.List(rCount, 1) = COPYING_DESTINATION_TOTYPE
        CopyingSettingsForm.CopyingSettingsListBox.List(rCount, 2) = c.item(COPYING_TOCOLUMN_KEY)
        
        Dim ToOffsetsStr As String
        
        ToOffsetsStr = CommonHelpers.GetFormattedStringFromOffsets(c.item(COPYING_TO_TOP_LEFT_CELL_COLUMN_OFFSET_KEY), c.item(COPYING_TO_TOP_LEFT_CELL_ROW_OFFSET_KEY), _
        c.item(COPYING_TO_RIGHT_BOTTOM_CELL_COLUMN_OFFSET_KEY), c.item(COPYING_TO_RIGHT_BOTTOM_CELL_ROW_OFFSET_KEY))
       
        CopyingSettingsForm.CopyingSettingsListBox.List(rCount, 3) = c.item(COPYING_TOBASECELL_KEY) & ";" & c.item(COPYING_TORANGE_KEY) & ToOffsetsStr
        CopyingSettingsForm.CopyingSettingsListBox.List(rCount, 4) = c.item(COPYING_TOWORKSHEET_KEY)
        CopyingSettingsForm.CopyingSettingsListBox.List(rCount, 5) = CommonHelpers.ExtractFileNameFromPath(c.item(COPYING_TOWORKBOOK_KEY))
    Next

End If

Set StoredSettings = Nothing

End Function

Private Sub DeleteCopyingConfigButton_Click()

Dim ReturnedChoice As Integer

ReturnedChoice = MsgBox(DELETE_CONFIRMATION, vbOKCancel, ATTENTION_TITLE)
If ReturnedChoice = vbCancel Then
    Exit Sub
End If

If ListBoxFunctions.IsListBoxHasSelectedItems(Me.CopyingConfigsListBox) = False Then
    MsgBox NO_SELECTED_ITEMS_ERROR_MSG, vbOKOnly, ERROR_TITLE
    Exit Sub
Else
    Dim i As Long
    Dim NewConfigNames As New Collection
    
    For i = 0 To Me.CopyingConfigsListBox.ListCount - 1
        If Me.CopyingConfigsListBox.Selected(i) = True Then
           DeleteConfig Me.CopyingConfigsListBox.List(i, 0)
        Else
            NewConfigNames.Add Me.CopyingConfigsListBox.List(i, 0)
        End If
    Next i
    
    ListBoxFunctions.UpdateListboxFromArray Me.CopyingConfigsListBox, NewConfigNames
    Set NewConfigNames = Nothing
End If

End Sub

Private Sub DeleteConfig(ByVal ConfigName As String)

m_MainSettingsStorage.DeleteSection COPYING_SECTION_KEY & SECTION_DELIMITER & ConfigName

End Sub

Public Property Get MainStorage() As VbaSettings

Set MainStorage = m_MainSettingsStorage
    
End Property

Private Sub SortingSettingsButton_Click()

If SortingSettingsForm.Visible = False Then
    
    Dim SelectedFoundFlag As Boolean
    SelectedFoundFlag = ListBoxFunctions.CopyAllSelectedItems(Me.MainMultiPageObject.Pages("SortingPage").WorksheetsNamesListBox, SortingSettingsForm.SortingSettingsListBox)
 
    If SelectedFoundFlag = True Then
        SortingSettingsForm.RetrieveWorksheetsSettings
        FormsHelpers.ChangeStateOfAllControlsOnForm ExcelUtilsMainWindow, False
        SortingSettingsForm.Show
    Else
        MsgBox NO_SELECTED_ITEMS_ERROR_MSG, vbOKOnly, ERROR_TITLE
        Exit Sub
    End If
    
End If

End Sub

Private Sub SelectAllButton_Click()

SetSelectionModeOfCurrentListBox True

End Sub

Private Sub UnselectAllButton_Click()

SetSelectionModeOfCurrentListBox False

End Sub

Private Sub SetSelectionModeOfCurrentListBox(ByVal SelMode As Boolean)

With Me.MainMultiPageObject
    Select Case .Pages(.Value).Name
        Case "SortingPage"
            ListBoxFunctions.SetSelectionModeOfListBox .Pages("SortingPage").WorksheetsNamesListBox, SelMode
        Case "CopyingPage"
            ListBoxFunctions.SetSelectionModeOfListBox .Pages("CopyingPage").CopyingConfigsListBox, SelMode
        Case "ColoringPage"
                ListBoxFunctions.SetSelectionModeOfListBox .Pages("ColoringPage").ColoringWorksheetsListBox, SelMode
    End Select
End With

End Sub

Private Sub DestroyObject()

Set m_MainSettingsStorage = Nothing

'unload child forms
If FormsHelpers.IsUserFormLoaded("SelectedSheetsSortSettingsForm") = True Then Unload SelectedSheetsSortSettingsForm
If FormsHelpers.IsUserFormLoaded("SortingSettingsForm") = True Then Unload SortingSettingsForm
If FormsHelpers.IsUserFormLoaded("CopyingSettingsForm") = True Then Unload CopyingSettingsForm
If FormsHelpers.IsUserFormLoaded("EditCopyingConfigForm") = True Then Unload EditCopyingConfigForm
If FormsHelpers.IsUserFormLoaded("CopyingGlobalSettingsForm") = True Then Unload CopyingGlobalSettingsForm

'unload current form
Unload Me

End Sub

Private Sub CloseExcelUtilsMainWindowButton_Click()

DestroyObject

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

'If CloseMode = 1 the Unload statement is invoked from code
If CloseMode <> 1 Then DestroyObject

End Sub

