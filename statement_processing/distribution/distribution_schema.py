"""
distribution_schema.py
"""


import logging

from statement_processing.statement_schema import ExpectedSheets
from typing import Optional
logger = logging.getLogger(__name__)


class DistributionSchema(ExpectedSheets):
    NAME_SHEET = ExpectedSheets.DELIVERY_APARTMENTS
    CORRESPONDENCE = ExpectedSheets.CORRESPONDENCE
    START_APARTMENTS_ROW = 7
    START_APARTMENTS_COLUMN = 2
    STRING_SEARCHING_MONTH = 1
    SEARCH_STRING_FOR_SUBCOLUMNS = 4 # строка для поиска под колонок
    MONTHS = {
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

    # Название все видов счетов
    ALL_ACCOUNTS = ExpectedSheets.ALL_ACCOUNTS

    """=======================================
    Название столбцов
    """
    LABEL_CURRENT = "Текущие взносы"
    LABEL_CUMULATIVE = "Накопительные взносы"
    LABEL_PURPOSE = "Целевые взносы"

    ANCHOR_APT_NUMBER = '№ квартиры'

    """
    Маппинг для связи:
    - текущий счет -> Текущие взносы
    - накопительный счет -> Накопительные взносы
    - целевой счет-> Целевые взносы
    """
    PAYMENT_TYPE_MAPPING : dict[str, str] = {
        ExpectedSheets.CURRENT_ACCOUNT: LABEL_CURRENT,
        ExpectedSheets.SAVINGS_ACCOUNT: LABEL_CUMULATIVE,
        ExpectedSheets.TARGET_ACCOUNT: LABEL_PURPOSE,
    }