
"""
distribution_utils.py
"""
import logging
from typing import Optional,Any
from openpyxl.worksheet.worksheet import Worksheet


logger = logging.getLogger(__file__)


def cell_values_sheet (sheet:Worksheet, row:int, column:int)-> Any:
    result = sheet.cell(row=row, column=column).value
    logger.debug(f"Значения ячейки: ({row}:{column})-{result}")
    return result

def writing_cell (sheet:Worksheet, row:int, column:int, value:Any):
    sheet.cell(row=row, column=column).value  = value
    logger.debug(f"Запись в ячейку [{row}:{column}]: {value}")