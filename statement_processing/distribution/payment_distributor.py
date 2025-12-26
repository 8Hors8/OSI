"""
payment_distributor.py
"""
import logging
from typing import Optional

logger = logging.getLogger(__file__)


class PaymentDistributor:
    """
    Отвечает за разнос банковских платежей в ведомость ОСИ
    согласно бизнес-правилам.
    """

    def __init__(self, book, payments_from_bank: list[dict[str,dict[str,str]]]):
        self.book = book
        self.bank_payments = payments_from_bank
        self.month = None



    def _getting_month(self, number_month: int) -> Optional[str]:

        months_ru = {
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

        self.month = months_ru.get(number_month)


        logger.debug(f'месяц {'выбран' if self.month is not None else 'Не выбран'} {self.month}')
        return self.month