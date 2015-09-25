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
Public Const SETTINGS_HEADER As String = "Данный лист содержит необходимые настройки для текущей Vba-программы."
Public Const SETTINGS_WORKSHEET_NAME As String = "VbaНастройки"

Public Const COPYING_DESTINATION_FROMTYPE As String = "Из"
Public Const COPYING_DESTINATION_TOTYPE As String = "В"
Public Const ATTENTION_TITLE As String = "Внимание!"
Public Const DELETE_CONFIRMATION As String = "Вы действительно хотите удалить?"
Public Const WORKBOOK_NAME As String = "книга "
Public Const COPYING_STARTED_MSG As String = "Процесс копирования запущен..."
Public Const COPYING_FINISHED_MSG As String = "Копирование было завершено."
Public Const SORTING_STARTED_MSG As String = "Процесс сортировки запущен..."
Public Const SORTING_FINISHED_MSG As String = "Сортировка была завершена."
Public Const CURRENT_COPYING_CONFIG_NAME_MSG As String = "Текущее имя конфигурации: "
Public Const CURRENT_SORTING_WORKSHEET_NAME_MSG As String = "Текущее имя сортируемого листа: "
Public Const CURRENT_SORTING_NUMBER_NAME_MSG As String = "Текущий табельный номер: "
Public Const COLORING_STARTED_MSG As String = "Процесс раскраски был запущен..."
Public Const COLORING_FINISHED_MSG As String = "Раскраска завершена."
Public Const CURRENT_WORKSHEET_NAME_MSG As String = "Текущий лист: "

'progressbar current operation texts
Public Const COPYING_FROM_WORKSHEET_CO As String = "Копирование из листа "
Public Const TO_WORKSHEET_CO As String = "в лист "

'forms titles
Public Const EXCEL_UTILS_MAIN_WINDOW_TITLE As String = "Дополнительные утилиты Excel " & VBA_PROGRAMM_VERSION
Public Const SORTING_SETTINGS_FORM_TITLE As String = "Настройки сортировки"
Public Const COPYING_SETTINGS_FORM_TITLE As String = "Настройки копирования"
Public Const COLORING_SETTINGS_FORM_TITLE As String = "Настройки раскраски"
Public Const EDIT_COPYING_CONFIG_FORM_TITLE As String = "Изменить текущий конфиг копирования"
Public Const SELECTED_SHEETS_SORTING_SETTINGS_FORM_TITLE As String = "Настройки выбранных листов"
Public Const SELECTED_SHEETS_COLORING_SETTINGS_FORM_TITLE As String = "Настройки выбранных листов"
Public Const SORTING_PAGE_TITLE As String = "Сортировка"
Public Const COPYING_PAGE_TITLE As String = "Копирование"
Public Const COLORING_PAGE_TITLE As String = "Раскраска"
Public Const COPYING_GLOBAL_SETTINGS_TITLE As String = "Глобальные настройки текущего конфига"
Public Const GET_PATH_TO_FILE_TITLE As String = "Получить путь к файлу"

'buttons
Public Const SELECTALL_BUTTON_TITLE As String = "Выделить все"
Public Const UNSELECTALL_BUTTON_TITLE As String = "Снять выделение"
Public Const SORTING_SETTINGS_BUTTON_TITLE As String = "Настройки"
Public Const START_SORTING_BUTTON_TITLE As String = "Начать сортировку"
Public Const START_COPYING_BUTTON_TITLE As String = "Начать копирование"
Public Const EDIT_BUTTON_TITLE As String = "Изменить"
Public Const ADD_BUTTON_TITLE As String = "Добавить"
Public Const DELETE_BUTTON_TITLE As String = "Удалить"
Public Const CANCEL_BUTTON_TITLE As String = "Закрыть"
Public Const SAVE_BUTTON_TITLE As String = "Сохранить"
Public Const SET_BUTTON_TITLE As String = "Задать"
Public Const CLEAR_BUTTON_TITLE As String = "Очистить"
Public Const BROWSE_BUTTON_TITLE As String = "Обзор"
Public Const GLOBAL_SETTINGS_BUTTON_TITLE As String = "Глобальные настройки"
Public Const SETTINGS_BUTTON_TITLE As String = "Настройки"
Public Const START_COLORING_BUTTON_TITLE As String = "Начать раскраску"

'labels
Public Const SORTING_LIST_BOX_DESCRIPTION_LABEL As String = "Список параметров сортировки выбранных листов: "
Public Const COPYING_LIST_BOX_DESCRIPTION_LABEL As String = "Список параметров копирования выбранного конфига: "
Public Const COLORING_LIST_BOX_DESCRIPTION_LABEL As String = "Список параметров раскраски выбранных листов: "

Public Const COPYING_CONFIGS_DESCRIPTION_LABEL As String = "Список конфигураций копирования: "
Public Const WORKSHEETS_LIST_LABEL As String = "Список всех листов текущей книги Excel: "
Public Const INPUT_SORTING_COLUMN_LABEL As String = "Введите адрес любой ячейки, которая содержит число для сортировки: "
Public Const INPUT_SORTING_OFFSETS_LABEL As String = "Задайте верхнюю левую и правую нижнюю ячейки, которые формируют диапазон для сортировки: "
Public Const INPUT_SERIAL_NUMBERS_LABEL As String = "(Опционально) Ввведите адрес ячейки с первым порядковым номером: "
Public Const INPUT_COPYING_CONFIG_NAME_LABEL As String = "Задайте имя конфигурации копирования:"
Public Const SORTING_WORKSHEETS_NAME_LABEL As String = "Название листа"
Public Const SORTING_COLUMN_LABEL As String = "Столбец сортировки"
Public Const SORTING_OFFSETS_LABEL As String = "Сдвиги сортировки"
Public Const COPYING_DIRECTION_LABEL As String = "Направление"
Public Const COPYING_COLUMN_LABEL As String = "Столбец"
Public Const COPYING_OFFSETS_LABEL As String = "Сдвиги"
Public Const COPYING_WORKSHEET_LABEL As String = "Лист"
Public Const COPYING_WORKBOOK_LABEL As String = "Книга"
Public Const COPYING_SPECIAL_OPERATION_LABEL As String = "Операция"
Public Const COPYING_PASTE_TYPE_LABEL As String = "Тип вставки"
Public Const COPYING_PASTE_PARAMETERS_LABEL As String = "Режимы вставки"
Public Const COPYING_COLOR_LABEL As String = "Цвет"
Public Const WORKSHEET_NAME_LABEL As String = "Лист: "
Public Const WORKBOOK_NAME_LABEL As String = "Путь к книге: "
Public Const BASE_CELL_LABEL As String = "Базовая ячейка: "
Public Const COPYING_RANGE_LABEL As String = "Диапазон копирования: "
Public Const COMMON_COPYING_LABEL As String = "Общие настройки копирования: "
Public Const FROM_WORKSHEET_COPYING_SETTINGS_LABEL As String = "Настройки листа, из которого производится копирование: "
Public Const TO_WORKSHEET_COPYING_SETTINGS_LABEL As String = "Настройки листа, в который производится копирование: "
Public Const USE_CURRENT_GLOBAL_SETTING_LABEL As String = "Использовать текущие глобальные настройки"
Public Const IS_REMOVED_GLOBAL_SETTING_LABEL As String = "Удалять параметр после копирования"

Public Const COLORING_BASERANGE_LABEL As String = "Базовый диапазон"
Public Const WORKSHEETNAME_LABEL As String = "Лист"
Public Const OFFSETS_LABEL As String = "Сдвиги"
Public Const COLOR_LABEL As String = "Цвет"

Public Const BASECELL_LABEL As String = "Задайте базовую ячейку: "
Public Const SOUGHTFOR_RANGE_LABEL As String = "Задайте искомый диапазон: "
Public Const BASERANGE_LABEL As String = "Задайте базовый диапазон: "
Public Const INPUT_COLOR_LABEL As String = "Задайте нужный цвет: "

'errors
Public Const ERROR_TITLE As String = "Ошибка! "
Public Const ERROR_FUNCTION_NAME As String = "Имя функции: "
Public Const ERROR_DETAILS As String = "Подробнее об ошибке: "
Public Const WARNING_TITLE As String = "Предупреждение! "

'[1] incorrect function arguments
Public Const INCORRECT_ARGS_ERROR_MSG As String = "В функцию были переданы некорректные аргументы. "
Public Const INCORRECT_WORKBOOKNAME_ERROR_MSG As String = "Был обнаружен некорректный путь к книге! "
Public Const INCORRECT_INPUT_VALUES_ERROR_MSG As String = "Были обнаружены некорректные входные данные!"

'[2] settings erros
Public Const INCORRECT_SORTING_SETTINGS_ERROR_MSG As String = "Были обнаружены некорректные настройки сортировки для выбранных листов. "
Public Const CANNOT_RESTORE_SETTINGS_ERROR_MSG As String = "Невозможно восстановить настройки! Проверьте, пожалуйста, заданные значения."
Public Const CANNOT_FIND_WORKSHEET_SETTINGS_ERROR_MSG As String = "Невозможно найти настройки для выбранного листа. "
Public Const INCORRECT_COPYING_CONFIG_NAME_ERROR_MSG As String = "Вы должны задать корректное имя конфигурации копирования! "
Public Const CANNOT_FIND_NECESSARY_VALUES_IN_CHOSEN_COLUMN_ERROR_MSG As String = "Невозможно обнаружить необходимые значения в выбранном столбце! "
Public Const EMPTY_SETTINGS_ERROR_MSG As String = "Для выбранной конфигурации не обнаружено настроек. "

Public Const COLUMN_IS_NOT_SET_ERROR_MSG As String = "Не задан столбец копирования. "
Public Const WORKSHEETNAME_IS_NOT_SET_ERROR_MSG As String = "Не задано имя листа копирования. "
Public Const TL_OFFSETS_ARE_NOT_SET_ERROR_MSG As String = "Не заданы сдвиги к верхней левой ячейке. "
Public Const COPYING_RANGE_IS_NOT_SET_ERROR_MSG As String = "Не задан диапазон копирования. "
Public Const COPYING_BASECELL_IS_NOT_SET_ERROR_MSG As String = "Не задана базовая ячейка копирования. "
Public Const COPYING_FROMWORKBOOK_IS_NOT_SET_ERROR_MSG As String = "Не задан путь к книге, из которой производится копирование. "
Public Const COPYING_TOWORKBOOK_IS_NOT_SET_ERROR_MSG As String = "Не задан путь к книге, в которую производится копирование. "

'[3] other errors
Public Const NO_SELECTED_ITEMS_ERROR_MSG As String = "Для начала необходимо выбрать из списка необходимые пункты!"
Public Const LISTBOX_ALREADY_HAS_ITEM_ERROR_MSG As String = "Список уже содержит данное имя конфигурации!"
Public Const CANNOT_REPLACE_ITEM_ERROR_MSG As String = "Невозможно заменить текущий сохраненный конфиг!"
Public Const TOO_MUCH_SELECTED_ITEMS_ERROR_MSG As String = "Было выбрано слишком много элементов в списке!"
Public Const WORKBOOK_NOT_FOUND_ERROR_MSG As String = "Выбранная книга не может быть найдена!"
Public Const THERE_ARE_NO_CELLS_WITH_COLOR_ERROR_MSG As String = "Отсутствуют ячейки с заданным цветом!"

