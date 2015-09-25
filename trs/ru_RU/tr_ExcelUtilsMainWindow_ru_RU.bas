Attribute VB_Name = "tr_ExcelUtilsMainWindow_ru_RU"
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
Public Const SETTINGS_HEADER As String = "������ ���� �������� ����������� ��������� ��� ������� Vba-���������."
Public Const SETTINGS_WORKSHEET_NAME As String = "Vba���������"

Public Const COPYING_DESTINATION_FROMTYPE As String = "��"
Public Const COPYING_DESTINATION_TOTYPE As String = "�"
Public Const ATTENTION_TITLE As String = "��������!"
Public Const DELETE_CONFIRMATION As String = "�� ������������� ������ �������?"
Public Const WORKBOOK_NAME As String = "����� "
Public Const COPYING_STARTED_MSG As String = "������� ����������� �������..."
Public Const COPYING_FINISHED_MSG As String = "����������� ���� ���������."
Public Const SORTING_STARTED_MSG As String = "������� ���������� �������..."
Public Const SORTING_FINISHED_MSG As String = "���������� ���� ���������."
Public Const CURRENT_COPYING_CONFIG_NAME_MSG As String = "������� ��� ������������: "
Public Const CURRENT_SORTING_WORKSHEET_NAME_MSG As String = "������� ��� ������������ �����: "
Public Const CURRENT_SORTING_NUMBER_NAME_MSG As String = "������� ��������� �����: "
Public Const COLORING_STARTED_MSG As String = "������� ��������� ��� �������..."
Public Const COLORING_FINISHED_MSG As String = "��������� ���������."
Public Const CURRENT_WORKSHEET_NAME_MSG As String = "������� ����: "

'progressbar current operation texts
Public Const COPYING_FROM_WORKSHEET_CO As String = "����������� �� ����� "
Public Const TO_WORKSHEET_CO As String = "� ���� "

'forms titles
Public Const EXCEL_UTILS_MAIN_WINDOW_TITLE As String = "�������������� ������� Excel " & VBA_PROGRAMM_VERSION
Public Const SORTING_SETTINGS_FORM_TITLE As String = "��������� ����������"
Public Const COPYING_SETTINGS_FORM_TITLE As String = "��������� �����������"
Public Const COLORING_SETTINGS_FORM_TITLE As String = "��������� ���������"
Public Const EDIT_COPYING_CONFIG_FORM_TITLE As String = "�������� ������� ������ �����������"
Public Const SELECTED_SHEETS_SORTING_SETTINGS_FORM_TITLE As String = "��������� ��������� ������"
Public Const SELECTED_SHEETS_COLORING_SETTINGS_FORM_TITLE As String = "��������� ��������� ������"
Public Const SORTING_PAGE_TITLE As String = "����������"
Public Const COPYING_PAGE_TITLE As String = "�����������"
Public Const COLORING_PAGE_TITLE As String = "���������"
Public Const COPYING_GLOBAL_SETTINGS_TITLE As String = "���������� ��������� �������� �������"
Public Const GET_PATH_TO_FILE_TITLE As String = "�������� ���� � �����"

'buttons
Public Const SELECTALL_BUTTON_TITLE As String = "�������� ���"
Public Const UNSELECTALL_BUTTON_TITLE As String = "����� ���������"
Public Const SORTING_SETTINGS_BUTTON_TITLE As String = "���������"
Public Const START_SORTING_BUTTON_TITLE As String = "������ ����������"
Public Const START_COPYING_BUTTON_TITLE As String = "������ �����������"
Public Const EDIT_BUTTON_TITLE As String = "��������"
Public Const ADD_BUTTON_TITLE As String = "��������"
Public Const DELETE_BUTTON_TITLE As String = "�������"
Public Const CANCEL_BUTTON_TITLE As String = "�������"
Public Const SAVE_BUTTON_TITLE As String = "���������"
Public Const SET_BUTTON_TITLE As String = "������"
Public Const CLEAR_BUTTON_TITLE As String = "��������"
Public Const BROWSE_BUTTON_TITLE As String = "�����"
Public Const GLOBAL_SETTINGS_BUTTON_TITLE As String = "���������� ���������"
Public Const SETTINGS_BUTTON_TITLE As String = "���������"
Public Const START_COLORING_BUTTON_TITLE As String = "������ ���������"

'labels
Public Const SORTING_LIST_BOX_DESCRIPTION_LABEL As String = "������ ���������� ���������� ��������� ������: "
Public Const COPYING_LIST_BOX_DESCRIPTION_LABEL As String = "������ ���������� ����������� ���������� �������: "
Public Const COLORING_LIST_BOX_DESCRIPTION_LABEL As String = "������ ���������� ��������� ��������� ������: "

Public Const COPYING_CONFIGS_DESCRIPTION_LABEL As String = "������ ������������ �����������: "
Public Const WORKSHEETS_LIST_LABEL As String = "������ ���� ������ ������� ����� Excel: "
Public Const INPUT_SORTING_COLUMN_LABEL As String = "������� ����� ����� ������, ������� �������� ����� ��� ����������: "
Public Const INPUT_SORTING_OFFSETS_LABEL As String = "������� ������� ����� � ������ ������ ������, ������� ��������� �������� ��� ����������: "
Public Const INPUT_SERIAL_NUMBERS_LABEL As String = "(�����������) �������� ����� ������ � ������ ���������� �������: "
Public Const INPUT_COPYING_CONFIG_NAME_LABEL As String = "������� ��� ������������ �����������:"
Public Const SORTING_WORKSHEETS_NAME_LABEL As String = "�������� �����"
Public Const SORTING_COLUMN_LABEL As String = "������� ����������"
Public Const SORTING_OFFSETS_LABEL As String = "������ ����������"
Public Const COPYING_DIRECTION_LABEL As String = "�����������"
Public Const COPYING_COLUMN_LABEL As String = "�������"
Public Const COPYING_OFFSETS_LABEL As String = "������"
Public Const COPYING_WORKSHEET_LABEL As String = "����"
Public Const COPYING_WORKBOOK_LABEL As String = "�����"
Public Const COPYING_SPECIAL_OPERATION_LABEL As String = "��������"
Public Const COPYING_PASTE_TYPE_LABEL As String = "��� �������"
Public Const COPYING_PASTE_PARAMETERS_LABEL As String = "������ �������"
Public Const COPYING_COLOR_LABEL As String = "����"
Public Const WORKSHEET_NAME_LABEL As String = "����: "
Public Const WORKBOOK_NAME_LABEL As String = "���� � �����: "
Public Const BASE_CELL_LABEL As String = "������� ������: "
Public Const COPYING_RANGE_LABEL As String = "�������� �����������: "
Public Const COMMON_COPYING_LABEL As String = "����� ��������� �����������: "
Public Const FROM_WORKSHEET_COPYING_SETTINGS_LABEL As String = "��������� �����, �� �������� ������������ �����������: "
Public Const TO_WORKSHEET_COPYING_SETTINGS_LABEL As String = "��������� �����, � ������� ������������ �����������: "
Public Const USE_CURRENT_GLOBAL_SETTING_LABEL As String = "������������ ������� ���������� ���������"
Public Const IS_REMOVED_GLOBAL_SETTING_LABEL As String = "������� �������� ����� �����������"

Public Const COLORING_BASERANGE_LABEL As String = "������� ��������"
Public Const WORKSHEETNAME_LABEL As String = "����"
Public Const OFFSETS_LABEL As String = "������"
Public Const COLOR_LABEL As String = "����"

Public Const BASECELL_LABEL As String = "������� ������� ������: "
Public Const SOUGHTFOR_RANGE_LABEL As String = "������� ������� ��������: "
Public Const BASERANGE_LABEL As String = "������� ������� ��������: "
Public Const INPUT_COLOR_LABEL As String = "������� ������ ����: "

'errors
Public Const ERROR_TITLE As String = "������! "
Public Const ERROR_FUNCTION_NAME As String = "��� �������: "
Public Const ERROR_DETAILS As String = "��������� �� ������: "
Public Const WARNING_TITLE As String = "��������������! "

'[1] incorrect function arguments
Public Const INCORRECT_ARGS_ERROR_MSG As String = "� ������� ���� �������� ������������ ���������. "
Public Const INCORRECT_WORKBOOKNAME_ERROR_MSG As String = "��� ��������� ������������ ���� � �����! "
Public Const INCORRECT_INPUT_VALUES_ERROR_MSG As String = "���� ���������� ������������ ������� ������!"

'[2] settings erros
Public Const INCORRECT_SORTING_SETTINGS_ERROR_MSG As String = "���� ���������� ������������ ��������� ���������� ��� ��������� ������. "
Public Const CANNOT_RESTORE_SETTINGS_ERROR_MSG As String = "���������� ������������ ���������! ���������, ����������, �������� ��������."
Public Const CANNOT_FIND_WORKSHEET_SETTINGS_ERROR_MSG As String = "���������� ����� ��������� ��� ���������� �����. "
Public Const INCORRECT_COPYING_CONFIG_NAME_ERROR_MSG As String = "�� ������ ������ ���������� ��� ������������ �����������! "
Public Const CANNOT_FIND_NECESSARY_VALUES_IN_CHOSEN_COLUMN_ERROR_MSG As String = "���������� ���������� ����������� �������� � ��������� �������! "
Public Const EMPTY_SETTINGS_ERROR_MSG As String = "��� ��������� ������������ �� ���������� ��������. "

Public Const COLUMN_IS_NOT_SET_ERROR_MSG As String = "�� ����� ������� �����������. "
Public Const WORKSHEETNAME_IS_NOT_SET_ERROR_MSG As String = "�� ������ ��� ����� �����������. "
Public Const TL_OFFSETS_ARE_NOT_SET_ERROR_MSG As String = "�� ������ ������ � ������� ����� ������. "
Public Const COPYING_RANGE_IS_NOT_SET_ERROR_MSG As String = "�� ����� �������� �����������. "
Public Const COPYING_BASECELL_IS_NOT_SET_ERROR_MSG As String = "�� ������ ������� ������ �����������. "
Public Const COPYING_FROMWORKBOOK_IS_NOT_SET_ERROR_MSG As String = "�� ����� ���� � �����, �� ������� ������������ �����������. "
Public Const COPYING_TOWORKBOOK_IS_NOT_SET_ERROR_MSG As String = "�� ����� ���� � �����, � ������� ������������ �����������. "

'[3] other errors
Public Const NO_SELECTED_ITEMS_ERROR_MSG As String = "��� ������ ���������� ������� �� ������ ����������� ������!"
Public Const LISTBOX_ALREADY_HAS_ITEM_ERROR_MSG As String = "������ ��� �������� ������ ��� ������������!"
Public Const CANNOT_REPLACE_ITEM_ERROR_MSG As String = "���������� �������� ������� ����������� ������!"
Public Const TOO_MUCH_SELECTED_ITEMS_ERROR_MSG As String = "���� ������� ������� ����� ��������� � ������!"
Public Const WORKBOOK_NOT_FOUND_ERROR_MSG As String = "��������� ����� �� ����� ���� �������!"
Public Const THERE_ARE_NO_CELLS_WITH_COLOR_ERROR_MSG As String = "����������� ������ � �������� ������!"

