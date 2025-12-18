"""
statements_manager.py

Модуль управления жизненным циклом файлов ведомостей ОСИ.

Функциональность:
- Загрузка и валидация банковских ведомостей в формате Excel.
- Обеспечение безопасного сохранения измененных данных.
- Обработка исключений при работе с файловой системой (ошибки доступа, отсутствие файлов).
- Предоставление интерфейса для взаимодействия между GUI и объектами openpyxl.
"""

import logging
from typing import Optional
from openpyxl.workbook import Workbook
from pathlib import Path

from core.excel_loader import load_excel_file


logger = logging.getLogger(__name__)

class ManagerStatements:
    """
    Класс-контроллер для управления объектом книги Excel (ведомостью).

    Отвечает за стабильную работу с файлом ведомости, предотвращая аварийное
    завершение программы при ошибках чтения или записи.
    """

    def __init__(self, path: str):
        """
        Инициализирует менеджер ведомости.

        :param path: Полный путь к Excel-файлу ведомости.
        """
        self.path: str = path
        self.name_file = Path(self.path).name
        self.book: Optional[Workbook] = None
        self.list_sheets:list|None = None

    def load_statements(self) -> bool:
        """
        Выполняет загрузку ведомости из файловой системы.

        Использует вспомогательную функцию load_excel_file для открытия книги.
        В случае неудачи (например, файл поврежден или не найден) метод
        возвращает False, позволяя GUI корректно обработать ситуацию без падения.

        :return: True — если файл успешно загружен, False — если возникла ошибка.
        """
        self.book = load_excel_file(self.path)

        if self.book is None:
            logger.warning(
                f"Загрузка прервана: файл по пути '{self.path}' не может быть открыт."
            )
            return False

        self.list_sheets = self._get_list_sheets()

        if not self.list_sheets:
            logger.error(
                f"Файл '{self.name_file}' загружен, но не содержит листов. "
                "Ведомость считается некорректной."
            )
            self.book = None
            return False

        logger.info(f"Ведомость успешно загружена: {self.name_file}")
        logger.debug(f"Список листов: {self.list_sheets}")
        return True

    def _get_list_sheets(self) -> list[str]:
        list_sheets = self.book.sheetnames if self.book is not None else []
        return list_sheets

    def save_statement(self) -> bool:
        """
        Выполняет сохранение текущего состояния книги по исходному пути.

        Метод защищен от наиболее распространенной ошибки в бухгалтерии —
        PermissionError (когда файл, в который нужно записать данные, открыт в Excel).

        :return: True — сохранение успешно, False — если доступ к файлу заблокирован.
        """
        if not self.book:
            logger.error("Ошибка сохранения: отсутствует объект книги (self.book is None).")
            return False

        try:
            # Сохраняем по тому же пути, откуда загрузили (self.path)
            self.book.save(self.path)
            logger.info(f"Данные успешно записаны в файл: {self.path}")
            return True

        except PermissionError:
            logger.error(
                f"Отказано в доступе: файл '{self.path}' открыт в другой программе. "
                "Пожалуйста, закройте Excel и повторите попытку сохранения."
            )
            return False

        except Exception as e:
            logger.error(f"Критическая ошибка при сохранении ведомости: {e}")
            return False

if __name__ == "__main__":
    # Настройка базового логирования для отладки модуля
    if not logger.hasHandlers():
        logging.basicConfig(
            level=logging.DEBUG,
            format="[%(asctime)s.%(msecs)03d] %(module)s:%(lineno)d %(levelname)7s - %(message)s"
        )

    # Тестовый запуск
    statement_path1 = r'D:\googleDriver\ОСИ исходники\Ведомость на 2026v1год.xlsx'
    statement_path2 = r'D:\googleDriver\ОСИ исходники\тест ведомости.xlsx'
    manager = ManagerStatements(statement_path2)
    manager.load_statements()
    print(manager.book.active)
    print(manager.list_sheets)