"""
Главный модуль графического интерфейса помощника ОСИ.

Функционал:
- Запрос у пользователя входных данных (файлы и параметры).
- Управление режимом логирования (INFO/DEBUG).
- Запуск обработки данных через модуль `statement`.
- Отображение логов и результатов в GUI.

Особенности:
- В режиме сборки в exe логирование по умолчанию работает на уровне INFO.
- Возможность включить отладочный режим через Checkbutton.
"""
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import logging
from config_logging import settings_logging
import statement  # ваш модуль с Assistant
from time import time

version = '1,2,0,1'

class OSIAssistantApp(tk.Tk):
    """
    Класс графического приложения помощника ОСИ.

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
        self.title(f"Помощник ОСИ v{version}")

        # Проверяем где мы стартовали
        self.is_frozen = getattr(sys, 'frozen', False)

        self.default_level = default_level

        # Интерфейс
        tk.Label(self, text="Укажите кол-во квартир").grid(row=0, column=0, sticky="w")
        self.kv_entry = tk.Entry(self)
        self.kv_entry.insert(0, "60")
        self.kv_entry.grid(row=0, column=1)

        # Выбор файла оплаты
        tk.Label(self, text="Выберите файл с оплатой").grid(row=1, column=0, sticky="w")
        self.bank_path = tk.Entry(self, width=50)
        self.bank_path.grid(row=1, column=1)
        tk.Button(self, text="Выбрать", command=self.select_bank_file).grid(row=1, column=2)

        # Выбор ведомости
        tk.Label(self, text="Выберите ведомость").grid(row=2, column=0, sticky="w")
        self.ved_path = tk.Entry(self, width=50)
        self.ved_path.grid(row=2, column=1)
        tk.Button(self, text="Выбрать", command=self.select_ved_file).grid(row=2, column=2)

        # Кнопки
        tk.Button(self, text="Запустить", command=self.run_assistant).grid(row=3, column=0)
        tk.Button(self, text="Очистить", command=self.clear_output).grid(row=3, column=1)
        tk.Button(self, text="Выход", command=self.quit).grid(row=3, column=2)

        # отображение Checkbutton
        self.control_flag = tk.IntVar()
        self.flag_checkbox = tk.Checkbutton(self, text='Режим отладки', variable=self.control_flag)
        self.flag_checkbox.grid(row=4, column=0, sticky='w')
        # Поле вывода
        self.output = scrolledtext.ScrolledText(self, width=80, height=20, state='disabled')
        self.output.grid(row=5, column=0, columnspan=3, padx=5, pady=5)

        # Настройка логирования на старте с выбранным уровнем
        initial_level = logging.INFO if self.is_frozen else self.default_level
        self.logger = settings_logging(initial_level, text_widget=self.output)

    def set_log_level(self, level):
        """
        Устанавливает уровень логирования и соответствующий формат сообщений.

        Args:
            level (int): Новый уровень логирования (logging.INFO или logging.DEBUG).
        """
        self.logger.setLevel(level)
        for handler in self.logger.handlers:
            handler.setLevel(level)
            # Меняем форматтер под уровень
            if level == logging.DEBUG:
                handler.setFormatter(self.logger._formatter_debug)
            else:
                handler.setFormatter(self.logger._formatter_info)

    def select_bank_file(self):
        """
         Открывает диалог выбора файла и устанавливает путь к файлу оплаты.
        """
        path = filedialog.askopenfilename()
        if path:
            self.bank_path.delete(0, tk.END)
            self.bank_path.insert(0, path)

    def select_ved_file(self):
        """
        Открывает диалог выбора файла и устанавливает путь к файлу ведомости.
        """
        path = filedialog.askopenfilename()
        if path:
            self.ved_path.delete(0, tk.END)
            self.ved_path.insert(0, path)

    def run_assistant(self):
        """
        Запускает обработку данных или тестовый режим в зависимости от чекбокса.

        - Если включен режим отладки, активирует уровень DEBUG и выводит тестовые сообщения.
        - Если выключен, запускает `statement.Assistant` с введенными параметрами.
        """


        if self.control_flag.get():
            lvl = logging.DEBUG if self.control_flag.get() else logging.INFO
            self.set_log_level(lvl)
            self.logger.debug(f'Запуск в режиме DEBUG={bool(self.control_flag.get())}')
            self.logger.info('запуск тестов будет когда то')
        else:
            try:
                path_bank = self.bank_path.get()
                path_ved = self.ved_path.get()
                kv_count = int(self.kv_entry.get())

                self.logger.info("Запуск обработки...")
                start = time()
                statement.Assistant(path_ved, path_bank, kv_count).launch()
                elapsed = round(time() - start, 2)

                self.logger.info(f"Готово! Время выполнения: {elapsed} сек.")
            except Exception as e:
                messagebox.showerror("Ошибка", str(e))

    def clear_output(self):
        """
        Очищает окно вывода логов.
        """
        self.output.configure(state='normal')
        self.output.delete(1.0, tk.END)
        self.output.configure(state='disabled')


if __name__ == '__main__':
    # При запуске из IDE передаём DEBUG, при exe внутри уже будет forced INFO,
    # но чекбокс позволит переключать
    app = OSIAssistantApp(default_level=logging.DEBUG)
    app.mainloop()
