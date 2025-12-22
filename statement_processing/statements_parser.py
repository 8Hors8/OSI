"""
statements_parser.py
"""

import logging
from openpyxl import Workbook

logger = logging.getLogger(__name__)


class UniversalScan:
    """
    Универсальный сканер Excel-листов по описательной схеме.

    Класс извлекает данные из листа Excel на основании схемы, которая
    описывает:
    - имя листа (NAME_SHEET);
    - координаты заголовка (ROW_START, COLUMN_START);
    - ожидаемое значение заголовка (EXPECTED_VALUE);
    - тип сканирования (SCAN_TYPE);
    - смещения для начала данных (ROW_OFFSET, COLUMN_OFFSET).

    Класс не занимается загрузкой файлов и не хранит состояние приложения.
    """

    def __init__(self, book: Workbook, schema: type):
        """
        Инициализирует сканер.

        :param book: Загруженная Excel-книга (Workbook).
        :param schema: Класс-схема с описанием структуры листа.
        """
        self.schema = schema

        self.sheet = book[getattr(schema, "NAME_SHEET")]
        self.row_start = getattr(schema, "ROW_START")
        self.column_start = getattr(schema, "COLUMN_START")
        self.expected_value = str(getattr(schema, "EXPECTED_VALUE")).lower()
        self.scan_type = getattr(schema, "SCAN_TYPE")

        header_value = self.sheet.cell(
            row=self.row_start,
            column=self.column_start
        ).value

        self.start_cell = (
            header_value.lower()
            if isinstance(header_value, str)
            else ""
        )

    def scan(self) -> list:
        """
        Выполняет сканирование листа в соответствии со схемой.

        Проверяет корректность заголовка и, в зависимости от типа сканирования,
        извлекает данные по строкам, столбцам или в табличном виде.

        :return: Список извлечённых значений.
        """
        if self.expected_value not in self.start_cell:
            logger.error(
                f"Ошибка: В ячейке {self.row_start}:{self.column_start} "
                f"не найдено '{self.expected_value}'. Найдено: '{self.start_cell}'"
            )
            return []

        logger.debug(f'Заголовок подтвержден "{self.expected_value}"')

        if self.scan_type == "row":
            logger.debug("Выбрано сканирование по строкам")
            return self._row_scan()

        if self.scan_type == "column":
            logger.debug("Выбрано сканирование по столбцам")
            return self._column_scan()

        if self.scan_type == "mixed":
            logger.debug("Выбрано смешанное сканирование")
            return self._mixed_scan()

        logger.error(f"Неизвестный тип сканирования: {self.scan_type}")
        return []

    def _row_scan(self) -> list:
        """
        Сканирует значения вниз по одному столбцу.

        :return: Список значений.
        """
        result = []
        row_offset = getattr(self.schema, "ROW_OFFSET", 1)

        for row in range(
                self.row_start + row_offset,
                self.sheet.max_row + 1
        ):
            value = self.sheet.cell(
                row=row,
                column=self.column_start
            ).value

            if self._validate_value(value):
                result.append(value)

        return result

    def _column_scan(self) -> list:
        """
        Сканирует значения вправо по одной строке.

        :return: Список значений.
        """
        result = []
        column_offset = getattr(self.schema, "COLUMN_OFFSET", 1)

        for column in range(
                self.column_start + column_offset,
                self.sheet.max_column + 1
        ):
            value = self.sheet.cell(
                row=self.row_start,
                column=column
            ).value

            if self._validate_value(value):
                result.append(value)

        return result

    def _mixed_scan(self) -> list:
        """
        Сканирует значения в табличной структуре (строки и столбцы).

        :return: Список значений.
        """
        result = []
        row_offset = getattr(self.schema, "ROW_OFFSET", 1)
        column_offset = getattr(self.schema, "COLUMN_OFFSET", 1)

        for row in range(
                self.row_start + row_offset,
                self.sheet.max_row + 1
        ):
            for column in range(
                    self.column_start + column_offset,
                    self.sheet.max_column + 1
            ):
                value = self.sheet.cell(
                    row=row,
                    column=column
                ).value

                if value is not None:
                    result.append(value)

        return result

    def _validate_value(self, value) -> bool:

        value_type = getattr(self.schema, "VALUE_TYPE")
        if value_type is None:
            return True
        return isinstance(value, value_type)

    def _validate_value_type(self):
        pass
