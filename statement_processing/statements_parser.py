"""
statements_parser.py
"""

import logging
from typing import Optional
from openpyxl import Workbook, worksheet

logger = logging.getLogger(__name__)



def universal_scan (sheet:Optional[worksheet], schema:Optional[type])->Optional[list[str]]:
    pass