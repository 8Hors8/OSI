import logging
import datetime
import sys
from typing import List, Dict, Any


# ============================================================
# 🧠 Умный Форматтер
# ============================================================
class SmartFormatter(logging.Formatter):
    def __init__(self, fmt_info: str, fmt_debug: str, datefmt: str):
        super().__init__(datefmt=datefmt)
        self.fmt_info = fmt_info  # Краткий
        self.fmt_debug = fmt_debug  # Подробный
        self.datefmt = datefmt

    def format(self, record: logging.LogRecord) -> str:
        # Сохраняем оригинальное сообщение, чтобы не дублировать "Источник" при повторных вызовах
        original_msg = record.msg

        # 1. Если передано extra={'custom_caller': ...}
        if hasattr(record, 'custom_caller'):
            record.msg = f"{original_msg} | Источник: {record.custom_caller}"
            self._style._fmt = self.fmt_debug

        # 2. Если это DEBUG — используем подробный формат
        elif record.levelno == logging.DEBUG:
            self._style._fmt = self.fmt_debug

        # 3. Для INFO, WARNING, ERROR — краткий формат (текст сообщения)
        else:
            self._style._fmt = self.fmt_info

        result = super().format(record)

        # Возвращаем оригинальное сообщение записи, чтобы не портить объект record
        record.msg = original_msg
        return result


# ============================================================
# 🟣 Handler для сбора событий GUI
# ============================================================
class DomainLogListener(logging.Handler):
    def __init__(self, events_list: List[Dict[str, Any]]):
        super().__init__()
        self.events_list = events_list

    def emit(self, record: logging.LogRecord):
        # Если к хендлеру прикреплен форматтер, используем его для времени
        if self.formatter:
            time_str = self.formatter.formatTime(record, self.formatter.datefmt)
        else:
            # Если нет — берем текущее время
            time_str = datetime.datetime.now().strftime("%H:%M:%S")

        # Собираем данные в список
        self.events_list.append({
            "time": time_str,
            "level": record.levelname,
            "message": record.getMessage(),
            "module": record.module,
            "line": record.lineno,
            "func": record.funcName
        })


# ============================================================
# ⚙️ Настройка логирования
# ============================================
def setup_logging(log_events: list = None, console_level=logging.DEBUG):
    root = logging.getLogger()
    root.setLevel(logging.DEBUG)

    if root.hasHandlers():
        root.handlers.clear()

    datefmt = "%H:%M:%S"

    # Формат 1: Краткий (для INFO)
    fmt_info = "%(levelname)-7s - %(message)s"

    # Формат 2: Подробный (для DEBUG и спец-вызовов)
    # [%(asctime)s] %(module)s:%(lineno)d [%(funcName)s] %(levelname)s - %(message)s
    fmt_debug = "[%(asctime)s] %(module)s:%(lineno)d [%(funcName)s] %(levelname)s - %(message)s"

    formatter = SmartFormatter(fmt_info, fmt_debug, datefmt)

    # --- Консоль ---
    console = logging.StreamHandler()
    console.setLevel(console_level)
    console.setFormatter(formatter)
    root.addHandler(console)

    # --- Сборщик событий (log_events) ---
    if log_events is not None:
        listener = DomainLogListener(log_events)
        listener.setLevel(logging.WARNING)
        root.addHandler(listener)

    return root