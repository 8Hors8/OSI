"""
payment_distributor.py
"""
import logging
from typing import Optional

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
                 apartment_numbers:list[str]):
        self.book = book
        self.apartment_numbers = apartment_numbers
        self.bank_payments = payments_from_bank
        self.month_name = None
        self.month_number = None
        self.expected_sheets = ExpectedSheets()


    def start_distribution(self, schema: type):
        allocation_apartments_sheet_name = getattr(schema, 'NAME_SHEET', None)
        allocation_apartments_sheet = self.book[allocation_apartments_sheet_name]
        max_row = allocation_apartments_sheet.max_row
        max_col = allocation_apartments_sheet.max_column
        for row in range(1,max_row):
            for col in range(1,max_col):
                cell_values_sheet(allocation_apartments_sheet, row, col)



    def _getting_month(self, month: int | str) -> Optional[str]:
        """
        Если передан int (1–12) → возвращает название месяца (str).
        Если передана строка с названием месяца → возвращает номер месяца (int).
        Если определить невозможно → None.
        """

        months = {
            1: "ЯНВАРЬ",
            2: "ФЕВРАЛЬ",
            3: "МАРТ",
            4: "АПРЕЛЬ",
            5: "МАЙ",
            6: "ИЮНЬ",
            7: "ИЮЛЬ",
            8: "АВГУСТ",
            9: "СЕНТЯБРЬ",
            10: "ОКТЯБРЬ",
            11: "НОЯБРЬ",
            12: "ДЕКАБРЬ",
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
