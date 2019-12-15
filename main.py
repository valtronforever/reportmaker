from tkinter import filedialog
import tkinter as tk
import traceback
from fileutils import list_files
from report import make_reports


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        tk.Label(self, text="Код учасника:").grid(row=0, column=0)

        self.member_code_c = tk.StringVar()
        self.member_code_entry = tk.Entry(self, textvariable=self.member_code_c)
        self.member_code_entry.grid(row=0, column=1, columnspan=2, sticky=tk.W + tk.E)

        tk.Label(self, text="Папка з кодами звітів:").grid(row=1, column=0)

        self.path_reports_c = tk.StringVar()
        self.path_reports_entry = tk.Entry(self, width=30, textvariable=self.path_reports_c)
        self.path_reports_entry.grid(row=1, column=1)

        self.reports_browsebutton = tk.Button(self, text="...", command=self.browse_reports)
        self.reports_browsebutton.grid(row=1, column=2)

        tk.Label(self, text="Папка з накладними:").grid(row=2, column=0)

        self.path_invoices_c = tk.StringVar()
        self.path_invoices_entry = tk.Entry(self, width=30, textvariable=self.path_invoices_c)
        self.path_invoices_entry.grid(row=2, column=1)

        self.invoices_browsebutton = tk.Button(self, text="...", command=self.browse_invoices)
        self.invoices_browsebutton.grid(row=2, column=2)

        tk.Label(self, text="Папка для збереження звітів:").grid(row=3, column=0)

        self.path_result_c = tk.StringVar()
        self.path_result_entry = tk.Entry(self, width=30, textvariable=self.path_result_c)
        self.path_result_entry.grid(row=3, column=1)

        self.result_browsebutton = tk.Button(self, text="...", command=self.browse_result)
        self.result_browsebutton.grid(row=3, column=2)

        actions_frame = tk.Frame(self).grid(row=4, column=0, columnspan=3)

        self.calculate = tk.Button(actions_frame, text="Сформувати звіти", command=self.calculate_reports)
        self.calculate.pack(side=tk.LEFT)

        self.status_label = tk.Label(actions_frame, text="")
        self.status_label.pack(padx=20, side=tk.RIGHT)

    def set_widgets_disabled(self):
        self.member_code_entry['state'] = tk.DISABLED
        self.path_reports_entry['state'] = tk.DISABLED
        self.reports_browsebutton['state'] = tk.DISABLED
        self.path_invoices_entry['state'] = tk.DISABLED
        self.invoices_browsebutton['state'] = tk.DISABLED
        self.path_result_entry['state'] = tk.DISABLED
        self.result_browsebutton['state'] = tk.DISABLED
        self.calculate['state'] = tk.DISABLED

    def set_widgets_enabled(self):
        self.member_code_entry['state'] = 'normal'
        self.path_reports_entry['state'] = 'normal'
        self.reports_browsebutton['state'] = 'normal'
        self.path_invoices_entry['state'] = 'normal'
        self.invoices_browsebutton['state'] = 'normal'
        self.path_result_entry['state'] = 'normal'
        self.result_browsebutton['state'] = 'normal'
        self.calculate['state'] = 'normal'

    def browse_reports(self):
        filename = filedialog.askdirectory()
        self.path_reports_c.set(filename)
        list_files(filename)

    def browse_invoices(self):
        filename = filedialog.askdirectory()
        self.path_invoices_c.set(filename)

    def browse_result(self):
        filename = filedialog.askdirectory()
        self.path_result_c.set(filename)

    def calculate_reports(self):
        self.set_widgets_disabled()
        self.status_label['text'] = 'Формуються звіти...'
        try:
            make_reports(
                self.member_code_c.get(),
                self.path_reports_c.get(),
                self.path_invoices_c.get(),
                self.path_result_c.get(),
            )
            self.set_widgets_enabled()
            self.status_label['text'] = 'Звіти успішно сформовано'
        except Exception:
            with open("log.txt", "w") as log:
                traceback.print_exc(file=log)

            self.set_widgets_enabled()
            self.status_label['text'] = 'Виникла помилка'


root = tk.Tk()
app = Application(master=root)
app.mainloop()
