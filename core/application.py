"""
application.py
"""

import logging

from bank.manager_bank import ManagerBank

logger = logging.getLogger(__name__)


class OSIApplication:
    """
    Центральный слой приложения.
    Связывает банковские данные и ведомость.
    """

    def __init__(self, bank_path: str, statement_path: str):
        self.bank_path = bank_path
        self.statement_path = statement_path
        self.bank = None
        self.statement = None

    def run(self):
        self.bank = ManagerBank(self.bank_path)
        self.bank.load_sheet()
        self.bank.acquire_payments()



if __name__ == '__main__':
    if not logger.hasHandlers():
        logging.basicConfig(
            level=logging.DEBUG,
            format="[%(asctime)s.%(msecs)03d] %(module)s:%(lineno)d %(levelname)7s - %(message)s"
        )
    bank_path1 = r'D:\googleDriver\ОСИ исходники\пробный вариант.xlsx'
    statement_path1 = r'D:\googleDriver\ОСИ исходники\Ведомость на 2026год.xls'
    app = OSIApplication(bank_path1, statement_path1)
    app.run()
    print(app.bank.data)
