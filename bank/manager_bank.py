"""
manager_bank.py

Менеджер банковского модуля ОСИ.

Модуль содержит класс ManagerBank, который отвечает за:
- загрузку банковского Excel-файла;
- инициализацию и хранение активного листа Excel;
- запуск парсинга банковских данных;
- формирование структурированных данных платежей.

ManagerBank не выполняет бизнес-логику разноски платежей
и не взаимодействует с ведомостью напрямую.
Он подготавливает данные для дальнейшей обработки
на уровне приложения (OSIApplication).
"""
import logging
from pathlib import Path
from typing import Optional

from core.excel_loader import load_excel_file
from .bank_parser import acquisition_data

logger = logging.getLogger(__name__)

class ManagerBank:
    """
        Управляет модулем банк
    """
    def __init__(self, path:str):
        self.path = path
        self.name_file = Path(self.path).name
        self.sheet = None
        self.data = None

    def load_sheet(self) -> bool:
        """
        Загружает Excel-файл банка.

        Пытается загрузить файл по пути self.path.
        В случае ошибки логирует причину и сообщает вызывающему коду
        о неудаче загрузки.

        :return: True — если файл успешно загружен, False — если произошла ошибка.
        """
        self.sheet = load_excel_file(self.path).active

        if self.sheet is None:
            logger.error(
                "Не удалось загрузить файл банка. "
                "Проверьте путь к файлу и повторите попытку."
            )
            return False
        logger.info(f'Банковский файл успешно загружен: "{self.name_file}"')
        return True

    def acquire_payments(self, apartment_number: list) -> Optional[dict[str, list[dict[str, str]]]]:
        """
        Выполняет парсинг банковского листа Excel и инициализирует основной
        словарь данных класса.

        Метод использует функцию acquisition_data() для преобразования
        сырых данных из self.sheet (объект Worksheet) в структурированный
        словарь, где ключ — номер квартиры, а значение — список всех платежей,
        сгруппированных по дате.

        По итогам работы:
            Записывает результат в self.data (Dict[str, List[Dict]]).

        :raises Exception: Пробрасывает исключения, возникшие в процессе парсинга
                           (например, ошибки чтения ячеек, неверный формат данных).
        """
        logger.info(f'Выполняется извлечение данных из банковского файла ...')
        apartment_number_reference = set(apartment_number)
        # acquisition_data — это функция, которая возвращает готовый словарь
        self.data = acquisition_data(self.sheet, apartment_number_reference)
        return self.data

if __name__ == '__main__':
    if not logger.hasHandlers():
        logging.basicConfig(
            level=logging.DEBUG,
            format="[%(asctime)s.%(msecs)03d] %(module)s:%(lineno)d %(levelname)7s - %(message)s"
        )
    form = ManagerBank(r'D:\googleDriver\ОСИ исходники\пробный вариант.xlsx')
    # print(form.sheet.cell(row=3, column=2).value)
    form.acquire_payments()
    print(form.data)
