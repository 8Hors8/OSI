"""
excel_loader.py

Модуль общего назначения для работы с Excel-файлами.

Отвечает за:
- загрузку Excel-книг (.xlsx);
- обработку типовых ошибок чтения файлов;
- централизованное логирование ошибок доступа и формата.

Модуль не содержит бизнес-логики и используется
банковскими и ведомственными парсерами.
"""

import logging
import openpyxl as op
from openpyxl.workbook import Workbook
from openpyxl.utils.exceptions import InvalidFileException
from typing import Optional

logger = logging.getLogger(__name__)


def load_excel_file(file_path: str) -> Optional[Workbook]:
    """
    Загружает Excel-книгу по указанному пути.

    Открывает файл Excel с использованием openpyxl и возвращает
    объект Workbook. В случае ошибок чтения, доступа или формата
    файла логирует причину и возвращает None.

    :param file_path: Полный путь к Excel-файлу.
    :return: Объект Workbook или None при ошибке загрузки.
    """
    try:
        # data_only=True — если нужен результат формул
        book = op.load_workbook(file_path, data_only=True)
        return book
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
