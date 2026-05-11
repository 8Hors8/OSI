"""
distribution_utils.py
"""
import logging
import sys
from pathlib import Path
from typing import Optional, Any
from openpyxl.worksheet.worksheet import Worksheet

logger = logging.getLogger(__name__)


def cell_values_sheet(sheet: Worksheet, row: int, column: int, log=False) -> Any:
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

    # Получаем имя функции и строку, которая вызвала эту утилиту
    f = sys._getframe(1)
    caller_info = f"{f.f_code.co_name}:{f.f_lineno}"

    # Передаем это через 'extra'. SmartFormatter сам поймет, что это нужно логгировать подробно.
    if log:
        logger.debug(
            f"Ячейка ({row}:{column}) -> '{result}'",
            extra={'custom_caller': caller_info}
        )

    return result


def writing_cell(sheet: Worksheet, row: int, column: int, value: Any):
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
    sheet.cell(row=row, column=column).value = value
    logger.debug(f"Запись в ячейку [{row}:{column}]: {value}")


def get_sorted_month_starts(buffer: dict) -> list[tuple[str, int]]:
    """
        Извлекает названия месяцев и их начальные колонки из буфера и сортирует их по порядку расположения в Excel.

        Функция сканирует словарь с данными о структуре листов и формирует упорядоченный список
        точек входа для каждого месяца. Это необходимо, так как словари в Python не гарантируют
        порядок, соответствующий визуальному порядку столбцов в таблице Excel.

        Args:
            buffer (dict): Словарь, где ключами являются названия месяцев (str),
                           а значениями — словари с метаданными месяца.
                           Обязательно наличие ключа 'start_col' (int).
                           Пример: {'февраль': {'start_col': 18}, 'январь': {'start_col': 12}}

        Returns:
            list[tuple[str, int]]: Список кортежей, отсортированный по номеру колонки.
                                   Пример: [('январь', 12), ('февраль', 18)]
                                   Если 'start_col' равен None, месяц игнорируется.

        Note:
            Сортировка выполняется по значению `x[1]` (номеру колонки). Это гарантирует,
            что логика обработки всегда будет идти слева направо, как в физической таблице.
        """
    return sorted(
        ((key, data['start_col']) for key, data in buffer.items() if data.get('start_col') is not None),
        key=lambda x: x[1]
    )


def build_column_ranges(items: list[tuple[str, int]]) -> list[tuple[str, int, int]]:
    """
    Преобразует список начальных координат месяцев в список полных интервалов колонок.

    Логика работы:
    1. Для каждого месяца (кроме последнего) конечной колонкой (end_col) считается
       начальная колонка следующего месяца.
    2. Для определения ширины последнего месяца используется вычисленная ширина
       предыдущих месяцев (предполагается, что блоки месяцев в Excel имеют одинаковый размер).

    Args:
        items: Список кортежей (название_месяца, начальная_колонка),
               отсортированный по возрастанию колонок.
               Пример: [('январь', 12), ('февраль', 18)]

    Returns:
        list[tuple[str, int, int]]: Список кортежей с границами:
               (месяц, начальная_колонка, конечная_колонка).
               Пример: [('январь', 12, 18), ('февраль', 18, 24)]
               Возвращает пустой список, если входные данные пусты или невозможно
               определить ширину шага (если передан только один месяц).

    Note:
        Функция критически важна для итерации по ячейкам конкретного месяца,
        так как задает границы для `range(start_col, end_col)`.
    """
    if not items:
        return []

    ranges = []
    column_range = None

    for i in range(len(items)):
        month, start_col = items[i]

        if i + 1 < len(items):
            end_col = items[i + 1][1]
            column_range = end_col - start_col
        else:
            if column_range is None:
                return []
            end_col = start_col + column_range

        ranges.append((month, start_col, end_col))

    return ranges
