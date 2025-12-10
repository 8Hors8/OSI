import logging
import openpyxl as op
from openpyxl.workbook import Workbook
from typing import Optional

logger = logging.getLogger(__name__)

def load_bank_file(file_path: str)-> Optional[Workbook]:
    """
    Парсит Excel-файл с банковской ведомостью.
    :param file_path: путь к файлу банка
    :return: активный лист или None
    """
    try:
        # data_only=True — если нужен результат формул
        return op.load_workbook(file_path, data_only=True)
    except FileNotFoundError:
        logger.error(f'Ошибка: Файл не найден по пути: "{file_path}"')
        return None
    except PermissionError:
        logger.error(f'Ошибка: Нет доступа. Закройте файл "{file_path}"')
        return None
    except op.utils.exceptions.InvalidFileException:
        logger.error(f'Ошибка: Файл "{file_path}" не является корректным Excel-файлом.')
        return None
    except Exception as e:
        logger.error(f'Непредвиденная ошибка при загрузке файла: {e}')
        return None


if __name__ == '__main__':
    if not logger.hasHandlers():
        logging.basicConfig(
            level=logging.DEBUG,
            format="[%(asctime)s.%(msecs)03d] %(module)s:%(lineno)d %(levelname)7s - %(message)s"
        )

    sheet = load_bank_file(r"D:\для теста оси\Новые ведомости\03.25.xlsx")
    if sheet:
        print(f"Открыт лист: {sheet.max_row}")
