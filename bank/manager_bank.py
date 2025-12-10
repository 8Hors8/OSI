
import logging
from parser import load_bank_file

log = logging.getLogger(__name__)

class ManagerBank:
    """
        Управляет модулем банк
    """
    def __init__(self, path:str):
        self.path = path
        self.bank = load_bank_file(path)



if __name__ == '__main__':
    form = ManagerBank(r'D:\googleDriver\ОСИ исходники\пробный вариант.xlsx')
    print(form.bank)
