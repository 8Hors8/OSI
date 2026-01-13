"""
payment_distributor.py
"""
import logging
import re
from typing import Optional

from openpyxl.worksheet.worksheet import Worksheet

from statement_processing.statement_schema import ExpectedSheets
from statement_processing.distribution.distribution_utils import cell_values_sheet,writing_cell
from statement_processing.distribution.distribution_schema import DistributionSchema

logger = logging.getLogger(__name__)


class PaymentDistributor:
    """
    Отвечает за разнос банковских платежей в ведомость ОСИ
    согласно бизнес-правилам.
    """

    def __init__(self, book, payments_from_bank: Optional[dict[str, list[dict[str, str]]]],
                 apartments_numbers:dict[str,type[int,int]]):
        self.book = book
        self.apartments_numbers = apartments_numbers
        self.bank_payments = payments_from_bank
        self.month_name = None
        self.month_number = None
        self.expected_sheets = ExpectedSheets()


    def start_distribution(self, schema: type):
        allocation_apartments_sheet_name = getattr(schema, 'NAME_SHEET', None)
        start_apartments_row = getattr(schema,'START_APARTMENTS_ROW',1)
        allocation_apartments_sheet = self.book[allocation_apartments_sheet_name]
        max_row = allocation_apartments_sheet.max_row
        max_col = allocation_apartments_sheet.max_column
        dict_month_column = self._search_monthly_columns(max_col, allocation_apartments_sheet)
        logger.debug(f'Значение месяц и столбец {dict_month_column}')
        for key, cell in self.apartments_numbers.items():
            self._test_name(allocation_apartments_sheet, str(key), cell[0],dict_month_column)


    def _process_apartment_payments(self, sheet: Worksheet,apartment_number: str, row: int, dict_month_column: dict ):
        """
        Обрабатывает платежи одной квартиры и подготавливает их к разноске.

        Метод:
        - извлекает банковские платежи по квартире;
        - сопоставляет их с колонками месяцев в ведомости;
        - определяет, какие суммы и в какие ячейки должны быть разнесены;
        - выполняет логирование расхождений и проблемных ситуаций.

        Метод не выполняет запись в Excel напрямую,
        а отвечает за анализ и подготовку данных для разноски.
        """


        a = self.bank_payments.get(apartment_number, None)
        logger.debug(f'получение из лплатежной ведомости {apartment_number}-{a}')


    def _search_monthly_columns(self, max_col: int, sheet: Worksheet):
        """
            Сканирует первую строку листа и формирует соответствие
            между названием месяца и номером колонки.

            Для каждой ячейки первой строки:
            - считывает значение;
            - если значение строковое, удаляет цифры (например, год),
              приводит к нижнему регистру и обрезает пробелы;
            - сохраняет результат как ключ словаря, где значением
              является номер колонки.

            Пример:
                "Январь 2026" -> {"январь": 3}

            :param max_col: Максимальное количество колонок листа.
            :param sheet: Лист Excel, в котором выполняется поиск.
            :return: Словарь вида {название_месяца: номер_колонки}.
            """
        result = {}
        for col in range(1,max_col):
            values = cell_values_sheet(sheet,1, col)
            if values is not None:
                value = re.sub(r'\d+','',values).strip().lower() if isinstance(values, str) else values

                result[value] = col
        return result
    def _getting_month(self, month: int | str) -> Optional[str]:
        """
        Если передан int (1–12) → возвращает название месяца (str).
        Если передана строка с названием месяца → возвращает номер месяца (int).
        Если определить невозможно → None.
        """

        months = {
            1: "январь",
            2: "февраль",
            3: "март",
            4: "апрель",
            5: "май",
            6: "июнь",
            7: "июль",
            8: "август",
            9: "сентябрь",
            10: "октябрь",
            11: "ноябрь",
            12: "декабрь",
        }

        # обратный словарь
        processed_value = month
        if isinstance(month, str):
            cleaned = month.strip()
            if cleaned.isdigit():
                processed_value = int(cleaned)
            else:
                processed_value = cleaned.upper()

        # 2. Логика "Номер -> Название"
        if isinstance(processed_value, int):
            result = months.get(processed_value)
            self.month_name = result
            logger.debug(f"Поиск по номеру месяца {processed_value}: {result or 'не найден'}")
            return result

        # 3. Логика "Название -> Номер"
        if isinstance(processed_value, str):
            months_reverse = {v: k for k, v in months.items()}
            result = months_reverse.get(processed_value)
            self.month_number = result
            logger.debug(f"Поиск по названию месяца '{processed_value}': {result or 'не найден'}")
            return result

        return None

    def _search_match_sheet(self, name_sheet: str) -> str:

        match = self.expected_sheets.CORRESPONDENCE.get(name_sheet)

        result = str(match) if match is not None else ""

        status = "выбран" if result else "Не выбран"
        logger.debug(f'Соответствие листов "{name_sheet}" {status}: "{result}"')

        return result

    def run_test(self):
        self.start_distribution(DistributionSchema)
