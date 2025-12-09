
import logging
from parser import bank_parser

log = logging.getLogger(__name__)

class ManagerBank:
    """
        Управляет модулем банк
    """
    def __init__(self, path:str):
        self.path = path
        self.bank = bank_parser(path)



if __name__ == '__main__':
    form = ManagerBank(r'D:\googleDriver\ОСИ исходники\пробный вариант.xlsx')
    print(form.bank)
