import logging
import sys
from pathlib import Path
from typing import List, Dict, Any

# ============================================================
# 🧠 Умный Форматтер: адаптирует детализацию под ситуацию
# ============================================================
class SmartFormatter(logging.Formatter):
    def __init__(self, fmt_default: str, fmt_full: str, datefmt: str):
        super().__init__(datefmt=datefmt)
        self.fmt_default = fmt_default
        self.fmt_full = fmt_full

    def format(self, record: logging.LogRecord) -> str:
        # 1. Если это ошибка (ERROR) или предупреждение (WARNING) — всегда полный формат
        if record.levelno >= logging.WARNING:
            self._style._fmt = self.fmt_full
        
        # 2. Если в записи есть атрибут 'custom_caller' (наш маркер для утилит)
        elif hasattr(record, 'custom_caller'):
            # Вставляем информацию о вызывающей функции прямо в текст сообщения
            record.msg = f"{record.msg} | Источник: {record.custom_caller}"
            self._style._fmt = self.fmt_full
        
        # 3. В обычном случае для DEBUG/INFO используем краткий вид
        else:
            self._style._fmt = self.fmt_default

        return super().format(record)

# ============================================================
# 🟢 Обработчики (Handlers)
# ============================================================
class TextHandler(logging.Handler):
    """Вывод логов в Tkinter Text виджет."""
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        try:
            msg = self.format(record)
            # Безопасное обновление GUI из любого потока
            self.text_widget.after(0, self._append, msg)
        except Exception:
            self.handleError(record)

    def _append(self, msg: str):
        self.text_widget.configure(state='normal')
        self.text_widget.insert('end', msg + '\n')
        self.text_widget.configure(state='disabled')
        self.text_widget.yview('end')

# ============================================================
# ⚙️ Настройка (Setup)
# ============================================================
def setup_logging(text_widget=None, console_level=logging.DEBUG):
    root = logging.getLogger()
    root.setLevel(logging.DEBUG)

    if root.hasHandlers():
        root.handlers.clear()

    # Настройки форматов
    datefmt = "%H:%M:%S"
    # Консоль: только суть
    fmt_short = "%(levelname)-7s - %(message)s"
    # GUI/Ошибки: время, модуль, строка, функция
    fmt_full = "[%(asctime)s] %(module)s:%(lineno)d [%(funcName)s] %(levelname)s - %(message)s"

    formatter = SmartFormatter(fmt_short, fmt_full, datefmt)

    # Консоль
    console = logging.StreamHandler()
    console.setLevel(console_level)
    console.setFormatter(formatter)
    root.addHandler(console)

    # Tkinter GUI
    if text_widget is not None:
        gui_text = TextHandler(text_widget)
        gui_text.setLevel(logging.DEBUG)
        gui_text.setFormatter(formatter)
        root.addHandler(gui_text)

    return root