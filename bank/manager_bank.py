"""
manager_bank.py
"""
import logging
from parser import load_bank_file, acquisition_data

logger = logging.getLogger(__name__)

class ManagerBank:
    """
        Управляет модулем банк
    """
    def __init__(self, path:str):
        self.path = path
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
        self.sheet = load_bank_file(self.path)

        if self.sheet is None:
            logger.error(
                "Не удалось загрузить файл банка. "
                "Проверьте путь к файлу и повторите попытку."
            )
            return False

        return True

    def acquire_payments(self):
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
        # acquisition_data — это функция, которая возвращает готовый словарь
        self.data = acquisition_data(self.sheet)


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
