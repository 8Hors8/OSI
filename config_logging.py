"""
Модуль конфигурации логирования для приложения.

Содержит:
- Класс `TextHandler` для вывода логов в Tkinter Text/ScrolledText.
- Функцию `settings_logging()` для инициализации логирования с возможностью выбора уровня.

Особенности:
- Поддержка двух форматов логов: краткий (INFO) и подробный (DEBUG).
- Автоматический вывод логов в GUI и консоль.
"""
import logging

class TextHandler(logging.Handler):
    """
       Обработчик логов для вывода сообщений в Tkinter Text/ScrolledText.

       Позволяет безопасно добавлять строки в текстовый виджет из главного потока.
   """
    def __init__(self, text_widget):
        """
        Args:
            text_widget (tk.Text): Tkinter виджет для вывода логов.
        """
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        """
        Обрабатывает запись лога и добавляет её в текстовый виджет.

        Args:
            record (logging.LogRecord): Объект с данными лога.
        """
        try:
            msg = self.format(record)
            # Работа с виджетом из главного потока (Tkinter не потокобезопасен)
            self.text_widget.after(0, self._append, msg)
        except Exception:
            self.handleError(record)

    def _append(self, msg):
        """
        Добавляет сообщение в конец текстового виджета.

        Args:
            msg (str): Текст сообщения.
        """
        self.text_widget.configure(state='normal')
        self.text_widget.insert('end', msg + '\n')
        self.text_widget.configure(state='disabled')
        self.text_widget.yview('end')

def settings_logging(level=logging.INFO, text_widget=None):
    """
       Настраивает систему логирования для приложения.

       Поддерживаются два режима:
       - INFO: краткие сообщения
       - DEBUG: подробные сообщения с метками времени, модуля и строки

       Args:
           level (int): Уровень логирования (logging.INFO или logging.DEBUG).
           text_widget (tk.Text): Виджет для вывода логов в GUI.

       Returns:
           logging.Logger: Настроенный экземпляр логгера.

       Raises:
           ValueError: Если `text_widget` не передан.
   """

    if text_widget is None:
        raise ValueError("Нужно передать text_widget для вывода логов")

    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)  # ловим всё и фильтруем по обработчикам


    if logger.hasHandlers():
        logger.handlers.clear()

    # Создаем два форматтера — для DEBUG и для INFO
    datefmt = "%Y-%m-%d %H:%M:%S"
    fmt_debug = "[%(asctime)s.%(msecs)03d] %(module)s:%(lineno)d %(levelname)7s - %(message)s"
    fmt_info = "%(message)s"

    formatter_debug = logging.Formatter(fmt_debug, datefmt=datefmt)
    formatter_info = logging.Formatter(fmt_info)

    # Обработчики
    text_handler = TextHandler(text_widget)
    console_handler = logging.StreamHandler()

    # По умолчанию применяем один из форматтеров
    formatter = formatter_debug if level == logging.DEBUG else formatter_info
    text_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    text_handler.setLevel(level)
    console_handler.setLevel(level)

    logger.addHandler(text_handler)
    logger.addHandler(console_handler)

    # Сохраняем форматтеры в атрибуты для последующего переключения
    logger._formatter_debug = formatter_debug
    logger._formatter_info = formatter_info

    return logger
