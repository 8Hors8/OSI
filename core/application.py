"""
application.py
"""

import logging

from bank.manager_bank import ManagerBank
from statement_processing.statements_manager import ManagerStatements
from statement_processing.statement_schema import ApartmentsSchema
from core.logging.domain_logger import DomainLogListener
from core.logging.config_logging import setup_logging


logger = logging.getLogger(__name__)


class OSIApplication:
    """
    Центральный слой приложения.
    Связывает банковские данные и ведомость.
    """

    def __init__(self, bank_path: str, statement_path: str):
        self.logger = logging.getLogger("OSIApplication")

        self.bank_path = bank_path
        self.statement_path = statement_path
        self.bank = None
        self.statement = None

    def run(self):
        logging.info("Запуск помощника ОСИ...")
        self.statement = ManagerStatements(self.statement_path)
        self.statement.load_statements()
        apartment_numbers = list(self.statement.get_apartment_numbers(ApartmentsSchema))
        self.bank = ManagerBank(self.bank_path)
        self.bank.load_sheet()

        payments_from_bank = self.bank.acquire_payments(apartment_numbers)
        logger.debug(f'payments_from_bank - {payments_from_bank}')
        self.statement.distribute_payments(payments_from_bank)

        # self.statement.save_statement()


if __name__ == '__main__':
    log_events = []  # Список для GUI

    setup_logging(
        log_events=log_events,
        console_level=logging.DEBUG
    )

    # 🔹 запуск приложения
    bank_path = r'D:\googleDriver\ОСИ исходники\пробный вариант.xlsx'
    statement_path = r'D:\googleDriver\ОСИ исходники\тест ведомости v1.xlsx'

    app = OSIApplication(bank_path, statement_path)
    app.run()
    if log_events:
        print("\n--- СОБЫТИЯ ДЛЯ GUI (Warning+) ---")
        for event in log_events:
            print(event)