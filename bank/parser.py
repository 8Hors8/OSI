import logging
import open_excel_file

logger = logging.getLogger(__name__)

def bank_parser(path_bank_form: str):
    """
    Парсит Excel-файл с банковской ведомостью.
    :param path_bank_form: путь к файлу банка
    :return: активный лист или None
    """
    result = {}
    workbook = open_excel_file.open_file(path_bank_form)
    if workbook:

        return result
    else:
        logger.error(f'Не удалось открыть файл {path_bank_form}')
        pass


if __name__ == '__main__':
    if not logger.hasHandlers():
        logging.basicConfig(
            level=logging.DEBUG,
            format="[%(asctime)s.%(msecs)03d] %(module)s:%(lineno)d %(levelname)7s - %(message)s"
        )

    sheet = bank_parser(r"D:\для теста оси\Новые ведомости\03.25.xlsx")
    if sheet:
        print(f"Открыт лист: {sheet.max_row}")
