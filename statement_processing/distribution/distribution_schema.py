"""
distribution_schema.py
"""


import logging

from statement_processing.statement_schema import ExpectedSheets

logger = logging.getLogger(__name__)


class DistributionSchema(ExpectedSheets):
    NAME_SHEET = ExpectedSheets.DELIVERY_APARTMENTS
    START_APARTMENTS_ROW = 7
    START_APARTMENTS_COLUMN = 2