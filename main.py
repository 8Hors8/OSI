import PySimpleGUI as sg
import statement

from time import time


def interfes():
    sg.theme('DarkAmber')  # Add a touch of color
    # All the stuff inside your window.
    layout = [
        [sg.Text('Укажите коллво квартир'), sg.Input(60, key="in_kv", size=(5, 1))],
        [sg.Text('Выберете файл с оплатой'), sg.InputText(), sg.FileBrowse('выбрать файл ')],
        [sg.Text('Выберете ведомость       '), sg.InputText(), sg.FileBrowse('выбрать файл ')],
        [sg.Button('Ok'), sg.Button('Cancel')],
        [sg.Output(size=(90, 20), key='output')],
        [sg.Button('Очистить')]
    ]

    # Create the Window
    window = sg.Window('Помощник ОСИ v-1.1.01', layout)

    # Event Loop to process "events" and get the "values" of the inputs

    def clear_output():
        """Очистка поля вывода"""
        for key in values:
            window['output'].update('')
        return None

    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED or event == 'Cancel':  # if user closes window or clicks cancel
            break
        elif event == 'Ok':

            start = time()
            path_ved = values[1]  #
            path_bank = values[0]
            kv_quantity = int(values['in_kv'])

            statement.Assistant(path_ved, path_bank, kv_quantity).launch()
            end = time() - start
            print(f'\nвремя работы программы {round(end, 2)} (сек.)\n')
        if event == 'Очистить':
            clear_output()

    window.close()


if __name__ == '__main__':
    interfes()
