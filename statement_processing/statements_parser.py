"""
statements_parser.py
"""

import logging
from typing import Optional
from openpyxl import Workbook, worksheet

logger = logging.getLogger(__name__)





def universal_scan(book: Optional[Workbook], schema: Optional[type]) -> list[str]:
    result = []
    if book is None or schema is None:
        return result
    sheet = book[getattr(schema, 'NAME_SHEET')]
    row_start = getattr(schema, 'ROW_START')
    column_start = getattr(schema, 'COLUMN_START')
    expected_value = getattr(schema, 'EXPECTED_VALUE').lower()
    start_cell = sheet.cell(row=row_start, column=column_start).value.lower()
    scan_type = getattr(schema,'SCAN_TYPE')
    logger.debug(f'получение начальной ячейки "{start_cell}" у объекта {sheet}')
    if expected_value in start_cell:
        logger.debug(f'Заголовок подтвержден "{expected_value}"')
        max_row = sheet.max_row
        max_column = sheet.max_column
        if scan_type == 'row':
            for row in range(row_start+1, max_row):
                value_cell = sheet.cell(row=row,column=column_start).value
                print(sheet.cell(row=row,column=column_start).value,f'строки {row}')
        elif scan_type == 'column':
            for column in range(column_start+1, max_column):
                value_cell = sheet.cell(row=row_start,column=column).value
                print(sheet.cell(row=row_start,column=column).value,f'строки {column}')
    else:
        logger.error(f"Ошибка: В ячейке {row_start}:{column_start} "
                     f"не найдено '{expected_value}'. Найдено: '{start_cell}'")
    return result

class UniversalScan:
    def __init__(self,book: Optional[Workbook], schema: Optional[type]):
        self.schema = schema
        self.sheet = book[getattr(schema, 'NAME_SHEET')]
        self.row_start = getattr(schema, 'ROW_START')
        self.column_start = getattr(schema, 'COLUMN_START')
        self.expected_value = getattr(schema, 'EXPECTED_VALUE').lower()
        self.scan_type = getattr(schema, 'SCAN_TYPE')

        self.sheet = book[getattr(schema, 'NAME_SHEET')]
        self.start_cell = self.sheet.cell(row=self.row_start, column=self.column_start).value.lower()

    def scan(self):

        if self.expected_value in self.start_cell:
            logger.debug(f'Заголовок подтвержден "{self.expected_value}"')

            if self.scan_type == 'row':
                return self._row_scan()
            elif self.scan_type == 'column':
                return self._column_skan()

        else:
            logger.error(f"Ошибка: В ячейке {self.row_start}:{self.column_start} "
                         f"не найдено '{self.expected_value}'. Найдено: '{self.start_cell}'")
        return []

    def _row_scan(self):
        result = []
        for row in range(self.row_start + 1, self.sheet.max_row):
            value_cell = self.sheet.cell(row=row, column=self.column_start).value
        return result

    def _column_skan(self):
        result = []
        for column in range(self.column_start + 1, self.sheet.max_column):
            value_cell = self.sheet.cell(row=self.row_start, column=column).value
        return result