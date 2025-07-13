import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import statement
from time import time

class OSIAssistantApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Помощник ОСИ v1.2.0")

        # Ввод количества квартир
        tk.Label(self, text="Укажите кол-во квартир").grid(row=0, column=0, sticky="w")
        self.kv_entry = tk.Entry(self)
        self.kv_entry.insert(0, "60")
        self.kv_entry.grid(row=0, column=1)

        # Выбор файла оплаты
        tk.Label(self, text="Выберите файл с оплатой").grid(row=1, column=0, sticky="w")
        self.bank_path = tk.Entry(self, width=50)
        self.bank_path.grid(row=1, column=1)
        tk.Button(self, text="Выбрать", command=self.select_bank_file).grid(row=1, column=2)

        # Выбор ведомости
        tk.Label(self, text="Выберите ведомость").grid(row=2, column=0, sticky="w")
        self.ved_path = tk.Entry(self, width=50)
        self.ved_path.grid(row=2, column=1)
        tk.Button(self, text="Выбрать", command=self.select_ved_file).grid(row=2, column=2)

        # Кнопки
        tk.Button(self, text="Запустить", command=self.run_assistant).grid(row=3, column=0)
        tk.Button(self, text="Очистить", command=self.clear_output).grid(row=3, column=1)
        tk.Button(self, text="Выход", command=self.quit).grid(row=3, column=2)

        # Поле вывода
        self.output = scrolledtext.ScrolledText(self, width=80, height=20)
        self.output.grid(row=4, column=0, columnspan=3, padx=5, pady=5)

    def select_bank_file(self):
        path = filedialog.askopenfilename()
        if path:
            self.bank_path.delete(0, tk.END)
            self.bank_path.insert(0, path)

    def select_ved_file(self):
        path = filedialog.askopenfilename()
        if path:
            self.ved_path.delete(0, tk.END)
            self.ved_path.insert(0, path)

    def run_assistant(self):
        try:
            path_bank = self.bank_path.get()
            path_ved = self.ved_path.get()
            kv_count = int(self.kv_entry.get())

            self.output.insert(tk.END, "Запуск обработки...\n")
            self.output.update()

            start = time()
            statement.Assistant(path_ved, path_bank, kv_count).launch()
            elapsed = round(time() - start, 2)

            self.output.insert(tk.END, f"\nГотово! Время выполнения: {elapsed} сек.\n")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def clear_output(self):
        self.output.delete(1.0, tk.END)


if __name__ == '__main__':
    app = OSIAssistantApp()
    app.mainloop()
