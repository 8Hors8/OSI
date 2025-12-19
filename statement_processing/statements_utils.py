"""
statements_utils.py

Вспомогательные функции для работы с ведомостями ОСИ.

Модуль содержит утилиты, не зависящие от состояния классов-менеджеров,
и предназначен для:
- проверки структуры ведомости;
- валидации наличия обязательных листов;
- поддержки модулей управления ведомостью и GUI.

Модуль использует описание ожидаемой структуры ведомости,
заданное в statement_schema.ExpectedSheets.
"""
import logging

from .statement_schema import ExpectedSheets

logger = logging.getLogger(__name__)

def checking_sheet_names(list_sheets: list[str]) -> bool:
    """
        Проверяет наличие всех обязательных листов ведомости.

        Функция сверяет список листов, полученных из Excel-книги,
        с набором обязательных листов, определённых в ExpectedSheets.ALL_SHEETS.

        Для каждого обязательного листа выполняется:
        - проверка наличия в list_sheets;
        - логирование результата проверки на уровне DEBUG.

        При отсутствии хотя бы одного обязательного листа
        функция немедленно возвращает False.

        :param list_sheets: Список имён листов, полученных из Excel-книги.
        :return: True — если все обязательные листы присутствуют,
                 False — если хотя бы один лист отсутствует.
        """


    for sheet in ExpectedSheets.ALL_SHEETS:
        check_result = sheet in list_sheets
        logger.debug(f'Лист "{sheet}"  результат проверки {check_result}')
        if not check_result:
            return False
    return True
