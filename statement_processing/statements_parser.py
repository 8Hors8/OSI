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
        self.row_offset = getattr(self.schema, "ROW_OFFSET", 1)
        self.column_offset = getattr(self.schema, "COLUMN_OFFSET", )

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

        for row in range(
                self.row_start + self.row_offset,
                self.sheet.max_row + 1
        ):
            value = self.sheet.cell(
                row=row,
                column=self.column_start + self.column_offset
            ).value

            if self._validate_value(value):
                result.append(value)
            else:
                logger.warning(f"Невалидное значение в ячейке "
                               f"{row}:{self.column_start + self.column_offset} — {value}")

        return result

    def _column_scan(self) -> list:
        """
        Сканирует значения вправо по одной строке.

        :return: Список значений.
        """
        result = []

        for column in range(
                self.column_start + self.column_offset,
                self.sheet.max_column + 1
        ):
            value = self.sheet.cell(
                row=self.row_start + self.row_offset,
                column=column
            ).value

            if self._validate_value(value):
                result.append(value)
            else:
                logger.warning(f"Невалидное значение в ячейке "
                               f"{self.row_start + self.row_offset}:{self.column_start + self.column_offset} — {value}")

        return result

    def _mixed_scan(self) -> list:
        """
        Сканирует значения в табличной структуре (строки и столбцы).

        :return: Список значений.
        """
        result = []

        for row in range(
                self.row_start + self.row_offset,
                self.sheet.max_row + 1
        ):
            for column in range(
                    self.column_start + self.column_offset,
                    self.sheet.max_column + 1
            ):
                value = self.sheet.cell(
                    row=row,
                    column=column
                ).value

                if value is not None:
                    result.append(value)
                else:
                    logger.warning(f"Невалидное значение в ячейке "
                                   f"{self.row_start + self.row_offset}:{self.column_start + self.column_offset} — {value}")

        return result

    def _validate_value(self, value) -> bool:
        """
        Проверяет одно значение согласно правилам схемы.

        Возвращает True — значение допустимо и может быть добавлено.
        Возвращает False — значение игнорируется.
        """

        # 1. IGNORE_VALUES
        ignore_values = getattr(self.schema, "IGNORE_VALUES", set())
        if value in ignore_values:
            logger.debug(f"Значение '{value}' проигнорировано (IGNORE_VALUES)")
            return False

        # 2. Пустые значения
        allow_empty = getattr(self.schema, "ALLOW_EMPTY", True)
        if value is None:
            if allow_empty:
                return True
            logger.debug("Пустое значение запрещено (ALLOW_EMPTY=False)")
            return False

        # 3. Тип значения
        value_type = getattr(self.schema, "VALUE_TYPE", None)
        if value_type is not None and not isinstance(value, value_type):
            logger.debug(
                f'Неверный тип значения: "{value}" - ({type(value)}), ожидался {value_type}'
            )
            return False

        # 4. Ограничения min / max
        min_value = getattr(self.schema, "MIN_VALUE", None)
        max_value = getattr(self.schema, "MAX_VALUE", None)

        if min_value is not None and value < min_value:
            logger.debug(f"Значение {value} меньше MIN_VALUE={min_value}")
            return False

        if max_value is not None and value > max_value:
            logger.debug(f"Значение {value} больше MAX_VALUE={max_value}")
            return False

        return True
