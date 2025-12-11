import logging

import openpyxl as op
import re
from openpyxl.workbook import Workbook
from typing import Optional
from openpyxl.utils.exceptions import InvalidFileException


logger = logging.getLogger(__name__)

def load_bank_file(file_path: str)-> Optional[Workbook]:
    """
    Парсит Excel-файл с банковской ведомостью.
    :param file_path: путь к файлу банка
    :return: активный лист или None
    """
    try:
        # data_only=True — если нужен результат формул
        book = op.load_workbook(file_path, data_only=True)
        sheet = book.active
        logger.info(f'Фйл с банка успешно загружен!')
        return sheet
    except FileNotFoundError:
        logger.error(f'Ошибка: Файл не найден по пути: "{file_path}"')
        return None
    except PermissionError:
        logger.error(f'Ошибка: Нет доступа. Закройте файл "{file_path}"')
        return None
    except InvalidFileException:
        logger.error(f'Ошибка: Файл "{file_path}" не является корректным Excel-файлом.')
        return None
    except Exception as e:
        logger.error(f'Непредвиденная ошибка при загрузке файла: {e}')
        return None


def extract_apartment_number (apartment_data:str) -> Optional[int]:
    pass


def acquisition_data(sheet) -> Optional[dict[str, list[str]]]:
    """
    Получаем данные с листа эксель вид счета на который зачисляются средства, номера квартир дату платежа
     и сумму платежа
     :param sheet: Лист банковского фала
     :return: dict с данными вида платежа и списком квартир
    """
    result = {}
    max_row = sheet.max_row
    print(f'max_row: {max_row}')
    for row in range(2, max_row + 1):
        print(f'row: {row}, значение;{sheet.cell(row=row, column=2).value}')
        if sheet.cell(row=row, column=2).value is not None:
            payment_type = sheet.cell(row=row, column=1).value
            apartment_number = sheet.cell(row=row, column=2).value
            sum_amount = sheet.cell(row=row, column=4).value
            data_amount = sheet.cell(row=row, column=5).value
            result[f'{apartment_number}']=[{'type': payment_type, 'sum': sum_amount, 'data': data_amount}]
    return result



if __name__ == '__main__':
    if not logger.hasHandlers():
        logging.basicConfig(
            level=logging.DEBUG,
            format="[%(asctime)s.%(msecs)03d] %(module)s:%(lineno)d %(levelname)7s - %(message)s"
        )


