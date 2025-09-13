"""
Главный модуль графического интерфейса помощника ОСИ.

Функционал:
- Запрос входных данных (файлы и параметры).
- Управление режимом логирования (INFO/DEBUG) с возможностью переключения "на лету".
- Запуск обработки данных через модуль `statement`.
- Отображение логов и результатов в GUI.

Особенности:
- В режиме запуска из IDE логирование по умолчанию DEBUG.
- В режиме exe логирование по умолчанию INFO.
"""

import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import logging
from time import time

from config_logging import settings_logging
import statement  # твой модуль с Assistant

VERSION = "1.2.0.1"


class OSIAssistantApp(tk.Tk):
    """
    Графическое приложение помощника ОСИ.

    Наследуется от `tk.Tk` и создает полноценное окно с формами для выбора файлов,
    управления параметрами и вывода логов.
    """

    def __init__(self, default_level=logging.DEBUG):
        """
        Инициализация интерфейса приложения.

        Args:
            default_level (int): Уровень логирования по умолчанию
                                 при запуске из исходников (обычно DEBUG).
        """
        super().__init__()
        self.title(f"Помощник ОСИ v{VERSION}")

        # Определяем, запущен ли exe или исходники
        self.is_frozen = getattr(sys, "frozen", False)

        # GUI-элементы
        self._build_interface()

        # Уровень логов: INFO для exe, DEBUG для кода
        initial_level = logging.INFO if self.is_frozen else default_level
        self.logger = settings_logging(initial_level, text_widget=self.output)

    def _build_interface(self) -> None:
        """Создаёт интерфейс приложения (виджеты и кнопки)."""

        # Количество квартир
        tk.Label(self, text="Укажите кол-во квартир").grid(row=0, column=0, sticky="w")
        self.kv_entry = tk.Entry(self)
        self.kv_entry.insert(0, "60")
        self.kv_entry.grid(row=0, column=1)

        # Файл оплаты
        tk.Label(self, text="Выберите файл с оплатой").grid(row=1, column=0, sticky="w")
        self.bank_path = tk.Entry(self, width=50)
        self.bank_path.grid(row=1, column=1)
        tk.Button(self, text="Выбрать", command=self.select_bank_file).grid(row=1, column=2)

        # Файл ведомости
        tk.Label(self, text="Выберите ведомость").grid(row=2, column=0, sticky="w")
        self.ved_path = tk.Entry(self, width=50)
        self.ved_path.grid(row=2, column=1)
        tk.Button(self, text="Выбрать", command=self.select_ved_file).grid(row=2, column=2)

        # Кнопки управления
        tk.Button(self, text="Запустить", command=self.run_assistant).grid(row=3, column=0)
        tk.Button(self, text="Очистить", command=self.clear_output).grid(row=3, column=1)
        tk.Button(self, text="Выход", command=self.quit).grid(row=3, column=2)

        # Чекбокс для включения режима отладки
        self.control_flag = tk.IntVar()
        self.flag_checkbox = tk.Checkbutton(
            self, text="Режим отладки", variable=self.control_flag,
            command=self._on_toggle_debug
        )
        self.flag_checkbox.grid(row=4, column=0, sticky="w")

        # Окно логов
        self.output = scrolledtext.ScrolledText(self, width=80, height=20, state="disabled")
        self.output.grid(row=5, column=0, columnspan=3, padx=5, pady=5)

    def _on_toggle_debug(self) -> None:
        """
        Срабатывает при переключении чекбокса "Режим отладки".
        Меняет уровень логирования и формат сообщений.
        """
        new_level = logging.DEBUG if self.control_flag.get() else logging.INFO
        self.set_log_level(new_level)
        self.logger.info(f"Уровень логирования переключен на {logging.getLevelName(new_level)}")

    def set_log_level(self, level: int) -> None:
        """
        Устанавливает уровень логирования и формат сообщений.

        Args:
            level (int): Новый уровень логирования (logging.INFO или logging.DEBUG).
        """
        self.logger.setLevel(level)
        for handler in self.logger.handlers:
            handler.setLevel(level)
            if level == logging.DEBUG:
                handler.setFormatter(self.logger._formatter_debug)
            else:
                handler.setFormatter(self.logger._formatter_info)

    def select_bank_file(self) -> None:
        """Открывает диалог выбора файла и устанавливает путь к файлу оплаты."""
        path = filedialog.askopenfilename()
        if path:
            self.bank_path.delete(0, tk.END)
            self.bank_path.insert(0, path)

    def select_ved_file(self) -> None:
        """Открывает диалог выбора файла и устанавливает путь к файлу ведомости."""
        path = filedialog.askopenfilename()
        if path:
            self.ved_path.delete(0, tk.END)
            self.ved_path.insert(0, path)

    def run_assistant(self) -> None:
        """
        Запускает обработку данных или тестовый режим в зависимости от чекбокса.

        - Если включен режим отладки, выводит тестовые сообщения.
        - Если выключен — запускает `statement.Assistant`.
        """
        if self.control_flag.get():
            # Тестовый запуск
            self.logger.debug("Тестовый запуск в режиме DEBUG")
            self.logger.info("Тестирование будет добавлено позже.")
        else:
            # Боевой запуск
            try:
                path_bank = self.bank_path.get()
                path_ved = self.ved_path.get()
                kv_count = int(self.kv_entry.get())

                self.logger.info("Запуск обработки...")
                start = time()
                statement.Assistant(path_ved, path_bank, kv_count).launch()
                elapsed = round(time() - start, 2)
                self.logger.info(f"Готово! Время выполнения: {elapsed} сек.")
            except Exception as exc:
                messagebox.showerror("Ошибка", str(exc))
                self.logger.exception("Ошибка при запуске обработки")

    def clear_output(self) -> None:
        """Очищает окно вывода логов."""
        self.output.configure(state="normal")
        self.output.delete(1.0, tk.END)
        self.output.configure(state="disabled")


if __name__ == "__main__":
    # При запуске из IDE → DEBUG, при запуске из exe → INFO
    app = OSIAssistantApp(default_level=logging.DEBUG)
    app.mainloop()
