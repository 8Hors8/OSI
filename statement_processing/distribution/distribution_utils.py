
"""
distribution_utils.py
"""
import logging
from typing import Optional,Any
from openpyxl.worksheet.worksheet import Worksheet


logger = logging.getLogger(__file__)


def cell_values_sheet (sheet:Worksheet, row:int, column:int)-> Any:
    """
    Возвращает значение ячейки Excel-листа по указанным координатам.

    Функция является утилитой-обёрткой над openpyxl и используется
    для централизованного доступа к значениям ячеек с логированием.

    Args:
        sheet (Worksheet): Лист Excel (openpyxl), из которого читается значение.
        row (int): Номер строки (нумерация начинается с 1).
        column (int): Номер столбца (нумерация начинается с 1).

    Returns:
        Any: Значение ячейки (str, int, float, datetime или None).

    Notes:
        - Функция не выполняет проверку корректности координат.
        - Логирует факт чтения ячейки на уровне DEBUG.
    """
    result = sheet.cell(row=row, column=column).value
    logger.debug(f"Значения ячейки: ({row}:{column})-{result}")
    return result

def writing_cell (sheet:Worksheet, row:int, column:int, value:Any):
    """
    Записывает значение в указанную ячейку Excel-листа.

    Используется для централизованной записи данных в ведомость
    с единым стилем логирования.

    Args:
        sheet (Worksheet): Лист Excel (openpyxl), в который производится запись.
        row (int): Номер строки (нумерация начинается с 1).
        column (int): Номер столбца (нумерация начинается с 1).
        value (Any): Значение для записи в ячейку.

    Returns:
        None

    Notes:
        - Функция не выполняет проверку допустимости значения.
        - Логирует факт записи ячейки на уровне DEBUG.
    """
    sheet.cell(row=row, column=column).value  = value
    logger.debug(f"Запись в ячейку [{row}:{column}]: {value}")