import openpyxl as op
import logging
import os

logger = logging.getLogger(__name__)

def open_file(path: str):
    """
    Открывает Excel-файл.
    :param path: путь к файлу
    :return: Workbook или None
    """

    name_file = os.path.basename(path)
    try:
        workbook = op.load_workbook(path)
        logger.info(f'Чтение файла {name_file} прошло успешно')
        return workbook
    except Exception as e:
        logger.error(f'Ошибка при открытии файла {name_file}: {e}')
        return None
