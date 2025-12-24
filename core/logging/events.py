"""
events.py
"""

from dataclasses import dataclass
from enum import Enum
from typing import Optional


class LogLevel(Enum):
    INFO = "INFO"
    WARNING = "WARNING"
    ERROR = "ERROR"


@dataclass
class LogEvent:
    level: LogLevel
    code: str
    message: str

    row: Optional[int] = None
    column: Optional[int] = None

    raw_value: Optional[str] = None
    parsed_value: Optional[str] = None
