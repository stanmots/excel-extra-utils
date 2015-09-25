Attribute VB_Name = "tr_ExcelUtilsMainWindow_en_US"
'*****************************
'* The ExcelUtilsMainWindow_tr_(your_country_codes_here) Module
'*
'* Short description:
'*
'*  Contains all string constants that are used in the current Vba program.
'*  Provides the easy way for translating the strings into the required language.
'*
'* Basic usage:
'*
'*  You can just replace the contents of this file with the contents of another file
'*  which has all the required translations.
'*
'*****************************

Option Explicit
Option Private Module

'base translations
Public Const SETTINGS_HEADER As String = "This worksheet contains settings for the Vba Program."
Public Const SETTINGS_WORKSHEET_NAME As String = "VbaSettings"
Public Const COPYING_DESTINATION_FROMTYPE As String = "From"
Public Const COPYING_DESTINATION_TOTYPE As String = "To"
Public Const ATTENTION_TITLE As String = "Attention!"
Public Const DELETE_CONFIRMATION As String = "Do you really want to delete?"
Public Const WORKBOOK_NAME As String = "workbook "
Public Const COPYING_STARTED_MSG As String = "The copying process has been started..."
Public Const COPYING_FINISHED_MSG As String = "The copying process has been finished."
Public Const SORTING_STARTED_MSG As String = "The sorting process has been started..."
Public Const SORTING_FINISHED_MSG As String = "The sorting process has been finished."
Public Const CURRENT_COPYING_CONFIG_NAME_MSG As String = "Current copying config name is: "
Public Const CURRENT_SORTING_WORKSHEET_NAME_MSG As String = "Current sorting worksheet name is: "
Public Const CURRENT_SORTING_NUMBER_NAME_MSG As String = "Current sorting number is: "
Public Const COLORING_STARTED_MSG As String = "The coloring process has been started..."
Public Const COLORING_FINISHED_MSG As String = "The coloring process has been finished."
Public Const CURRENT_WORKSHEET_NAME_MSG As String = "Current worksheet name is: "

'progressbar current operation texts
Public Const COPYING_FROM_WORKSHEET_CO As String = "Copying from worksheet "
Public Const TO_WORKSHEET_CO As String = "to worksheet "


'forms titles
Public Const EXCEL_UTILS_MAIN_WINDOW_TITLE As String = "Excel Additional Tools and Utilities " & VBA_PROGRAMM_VERSION
Public Const SORTING_SETTINGS_FORM_TITLE As String = "Sorting Settings"
Public Const COPYING_SETTINGS_FORM_TITLE As String = "Copying Setttings"
Public Const COLORING_SETTINGS_FORM_TITLE As String = "Coloring Setttings"
Public Const EDIT_COPYING_CONFIG_FORM_TITLE As String = "Edit current copying config"
Public Const SELECTED_SHEETS_SORTING_SETTINGS_FORM_TITLE As String = "Settings of the selected worksheets"
Public Const SELECTED_SHEETS_COLORING_SETTINGS_FORM_TITLE As String = "Settings of the selected worksheets"
Public Const SORTING_PAGE_TITLE As String = "Sorting"
Public Const COPYING_PAGE_TITLE As String = "Copying"
Public Const COLORING_PAGE_TITLE As String = "Coloring"
Public Const COPYING_GLOBAL_SETTINGS_TITLE As String = "Global Settings of the Current Config"
Public Const GET_PATH_TO_FILE_TITLE As String = "Get the path to a file"

'buttons
Public Const SELECTALL_BUTTON_TITLE As String = "Select All"
Public Const UNSELECTALL_BUTTON_TITLE As String = "Unselect All"
Public Const SORTING_SETTINGS_BUTTON_TITLE As String = "Settings"
Public Const START_SORTING_BUTTON_TITLE As String = "Start Sorting"
Public Const START_COPYING_BUTTON_TITLE As String = "Start Copying"
Public Const EDIT_BUTTON_TITLE As String = "Edit"
Public Const ADD_BUTTON_TITLE As String = "Add"
Public Const DELETE_BUTTON_TITLE As String = "Delete"
Public Const CANCEL_BUTTON_TITLE As String = "Cancel"
Public Const SAVE_BUTTON_TITLE As String = "Save"
Public Const SET_BUTTON_TITLE As String = "Set"
Public Const CLEAR_BUTTON_TITLE As String = "Clear"
Public Const BROWSE_BUTTON_TITLE As String = "Browse"
Public Const GLOBAL_SETTINGS_BUTTON_TITLE As String = "Global Settings"
Public Const SETTINGS_BUTTON_TITLE As String = "Settings"
Public Const START_COLORING_BUTTON_TITLE As String = "Start Coloring"

'labels
Public Const SORTING_LIST_BOX_DESCRIPTION_LABEL As String = "List with the sorting settings of the selected worksheets:"
Public Const COPYING_LIST_BOX_DESCRIPTION_LABEL As String = "List with the copying settings of the selected config:"
Public Const COLORING_LIST_BOX_DESCRIPTION_LABEL As String = "List with the coloring settings of the selected worksheets:"

Public Const COPYING_CONFIGS_DESCRIPTION_LABEL As String = "List with the copying configurations:"
Public Const WORKSHEETS_LIST_LABEL As String = "List of all the worksheets of the current workbook:"
Public Const INPUT_SORTING_COLUMN_LABEL As String = "Please, enter any address of the cell which has the sorting number:"
Public Const INPUT_SORTING_OFFSETS_LABEL As String = "Please, enter the top-left and right-bottom cells which define the sorting range of the one entry:"
Public Const INPUT_SERIAL_NUMBERS_LABEL As String = "(Optional) Please, enter the cell with the first serial number"
Public Const INPUT_COPYING_CONFIG_NAME_LABEL As String = "Please, enter the copying config name:"
Public Const SORTING_WORKSHEETS_NAME_LABEL As String = "Worksheet Name"
Public Const SORTING_COLUMN_LABEL As String = "Sorting Column"
Public Const SORTING_OFFSETS_LABEL As String = "Sorting Offsets"
Public Const COPYING_DIRECTION_LABEL As String = "Direction"
Public Const COPYING_COLUMN_LABEL As String = "Column"
Public Const COPYING_OFFSETS_LABEL As String = "Offsets"
Public Const COPYING_WORKSHEET_LABEL As String = "Worksheet"
Public Const COPYING_WORKBOOK_LABEL As String = "Workbook"
Public Const COPYING_SPECIAL_OPERATION_LABEL As String = "Operation"
Public Const COPYING_PASTE_TYPE_LABEL As String = "Paste Type"
Public Const COPYING_PASTE_PARAMETERS_LABEL As String = "Paste Modes"
Public Const COPYING_COLOR_LABEL As String = "Color"
Public Const WORKSHEET_NAME_LABEL As String = "Worksheet name:"
Public Const WORKBOOK_NAME_LABEL As String = "Workbook path:"
Public Const BASE_CELL_LABEL As String = "Base cell:"
Public Const COPYING_RANGE_LABEL As String = "Copying range:"
Public Const COMMON_COPYING_LABEL As String = "Common copying settings:"
Public Const FROM_WORKSHEET_COPYING_SETTINGS_LABEL As String = "Settings of the worksheet from which you need copying:"
Public Const TO_WORKSHEET_COPYING_SETTINGS_LABEL As String = "Settings of the worksheet to which you need copying:"
Public Const USE_CURRENT_GLOBAL_SETTING_LABEL As String = "Use current global setting"
Public Const IS_REMOVED_GLOBAL_SETTING_LABEL As String = "Remove setting after copying"

Public Const COLORING_BASERANGE_LABEL As String = "Base Range"
Public Const WORKSHEETNAME_LABEL As String = "Worksheet"
Public Const OFFSETS_LABEL As String = "Offsets"
Public Const COLOR_LABEL As String = "Color"

Public Const BASECELL_LABEL As String = "Input the base cell: "
Public Const SOUGHTFOR_RANGE_LABEL As String = "Input the sought-for range: "
Public Const BASERANGE_LABEL As String = "Input the base range: "
Public Const INPUT_COLOR_LABEL As String = "Input the color code: "

'errors
Public Const ERROR_TITLE As String = "Error! "
Public Const ERROR_FUNCTION_NAME As String = "Function name: "
Public Const ERROR_DETAILS As String = "Details about the error: "
Public Const WARNING_TITLE As String = "Warning! "

'[1] incorrect function arguments
Public Const INCORRECT_ARGS_ERROR_MSG As String = "Incorrect arguments were passed into the function. "
Public Const INCORRECT_WORKBOOKNAME_ERROR_MSG As String = "An incorrect workbookname was passed! "
Public Const INCORRECT_INPUT_VALUES_ERROR_MSG As String = "Incorrect input values have been found!"

'[2] settings erros
Public Const INCORRECT_SORTING_SETTINGS_ERROR_MSG As String = "Incorrect sorting settings of the selected worksheets have been found."
Public Const CANNOT_RESTORE_SETTINGS_ERROR_MSG As String = "Cannot restore the settings! Please, check your input values."
Public Const CANNOT_FIND_WORKSHEET_SETTINGS_ERROR_MSG As String = "Cannot find the settings for the worksheet."
Public Const INCORRECT_COPYING_CONFIG_NAME_ERROR_MSG As String = "You must input a proper copying config name!"
Public Const CANNOT_FIND_NECESSARY_VALUES_IN_CHOSEN_COLUMN_ERROR_MSG As String = "Cannot find the necessary values in the chosen column!"
Public Const EMPTY_SETTINGS_ERROR_MSG As String = "There are no settings for the specified config."

Public Const COLUMN_IS_NOT_SET_ERROR_MSG As String = "The copying column is not set."
Public Const WORKSHEETNAME_IS_NOT_SET_ERROR_MSG As String = "The copying worksheet name is not set."
Public Const TL_OFFSETS_ARE_NOT_SET_ERROR_MSG As String = "The top-left offsets are not set."
Public Const COPYING_RANGE_IS_NOT_SET_ERROR_MSG As String = "The copying range is not set."
Public Const COPYING_BASECELL_IS_NOT_SET_ERROR_MSG As String = "The copying basecell is not set."
Public Const COPYING_FROMWORKBOOK_IS_NOT_SET_ERROR_MSG As String = "The copying from-workbook name is not set."
Public Const COPYING_TOWORKBOOK_IS_NOT_SET_ERROR_MSG As String = "The copying to-workbook name is not set."

'[3] other errors
Public Const NO_SELECTED_ITEMS_ERROR_MSG As String = "You must select some items before proceeding further!"
Public Const LISTBOX_ALREADY_HAS_ITEM_ERROR_MSG As String = "List box already has such name of the config!"
Public Const CANNOT_REPLACE_ITEM_ERROR_MSG As String = "Cannot replace the current stored copying config!"
Public Const TOO_MUCH_SELECTED_ITEMS_ERROR_MSG As String = "The ListBox has too much selected items!"
Public Const WORKBOOK_NOT_FOUND_ERROR_MSG As String = "The Workbook you have specified was not found!"
Public Const THERE_ARE_NO_CELLS_WITH_COLOR_ERROR_MSG As String = "There are no cells with required color!"



