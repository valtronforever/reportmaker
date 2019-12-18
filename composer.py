from datetime import datetime, timedelta, date
import calendar
from tkinter import filedialog
import tkinter as tk
from tkinter import ttk
import traceback
from report import compose_report

BROWSE_BUTTON_WIDTH = 2
BROWSE_BUTTON_HEIGHT = 2


def monthdelta(date, delta):
    m, y = (date.month + delta) % 12, date.year + ((date.month) + delta - 1) // 12
    if not m:
        m = 12
    d = min(date.day, calendar.monthrange(y, m)[1])
    return date.replace(day=d, month=m, year=y)


class Application(ttk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        self.create_widgets()

    def create_widgets(self):
        ttk.Label(self, text="Початок звіту:").grid(row=0, column=0, sticky=tk.W)
        self.start = date.today().replace(day=1)
        self.start_c = tk.StringVar(value=self.start.strftime("%Y-%m-%d"))
        self.start_entry = ttk.Entry(self, textvariable=self.start_c)
        self.start_entry.grid(row=0, column=1, sticky=tk.W + tk.E, pady=(2, 4,))
        start_actions_frame = ttk.Frame(self)
        start_actions_frame.grid(row=0, column=2, padx=(3, 0,), pady=2)
        self.start_minus_button = ttk.Button(start_actions_frame, text="-", command=self.start_minus, width=1)
        self.start_minus_button.pack(side=tk.LEFT)
        self.start_plus_button = ttk.Button(start_actions_frame, text="+", command=self.start_plus, width=1)
        self.start_plus_button.pack(side=tk.LEFT)

        ttk.Label(self, text="Кінець звіту:").grid(row=1, column=0, sticky=tk.W)
        self.end = (date.today().replace(day=1) + timedelta(days=32)).replace(day=1)
        self.end_c = tk.StringVar(value=self.end.strftime("%Y-%m-%d"))
        self.end_entry = ttk.Entry(self, textvariable=self.end_c)
        self.end_entry.grid(row=1, column=1, sticky=tk.W + tk.E, pady=(2, 4,))
        end_actions_frame = ttk.Frame(self)
        end_actions_frame.grid(row=1, column=2, padx=(3, 0,), pady=2)
        self.end_minus_button = ttk.Button(end_actions_frame, text="-", command=self.end_minus, width=1)
        self.end_minus_button.pack(side=tk.LEFT)
        self.end_plus_button = ttk.Button(end_actions_frame, text="+", command=self.end_plus, width=1)
        self.end_plus_button.pack(side=tk.LEFT)

        ttk.Label(self, text="Звіт (з кодами):").grid(row=2, column=0, sticky=tk.W)
        self.path_report_c = tk.StringVar()
        self.path_report_entry = ttk.Entry(self, width=30, textvariable=self.path_report_c)
        self.path_report_entry.grid(row=2, column=1)
        self.report_browsebutton = ttk.Button(self, text="...", command=self.browse_report, width=BROWSE_BUTTON_WIDTH)
        self.report_browsebutton.grid(row=2, column=2, padx=(3, 0,), pady=2)

        ttk.Label(self, text="Папка зі звітами учасників:").grid(row=3, column=0, sticky=tk.W)
        self.path_members_c = tk.StringVar()
        self.path_members_entry = ttk.Entry(self, width=30, textvariable=self.path_members_c)
        self.path_members_entry.grid(row=3, column=1)
        self.members_browsebutton = ttk.Button(self, text="...", command=self.browse_members, width=BROWSE_BUTTON_WIDTH)
        self.members_browsebutton.grid(row=3, column=2, padx=(3, 0,), pady=2)

        ttk.Label(self, text="Папка для збереження:").grid(row=4, column=0, sticky=tk.W)
        self.path_result_c = tk.StringVar()
        self.path_result_entry = ttk.Entry(self, width=30, textvariable=self.path_result_c)
        self.path_result_entry.grid(row=4, column=1)
        self.result_browsebutton = ttk.Button(self, text="...", command=self.browse_result, width=BROWSE_BUTTON_WIDTH)
        self.result_browsebutton.grid(row=4, column=2, padx=(3, 0,), pady=2)

        actions_frame = ttk.Frame(self).grid(row=4, column=0, columnspan=3)

        self.calculate = ttk.Button(actions_frame, text="Скомпонувати звіт", command=self.calculate_reports)
        self.calculate.pack(side=tk.LEFT, padx=10, pady=10)

        self.status_label = ttk.Label(actions_frame, text="")
        self.status_label.pack(padx=20, side=tk.RIGHT)

    def start_minus(self):
        try:
            self.status_label['text'] = ''
            start_date = datetime.strptime(self.start_c.get(), "%Y-%m-%d").date()
            self.start_c.set(monthdelta(start_date, -1).strftime("%Y-%m-%d"))
        except Exception:
            self.status_label['text'] = 'Невірний формат: дата початку'

    def start_plus(self):
        try:
            self.status_label['text'] = ''
            start_date = datetime.strptime(self.start_c.get(), "%Y-%m-%d").date()
            self.start_c.set(monthdelta(start_date, 1).strftime("%Y-%m-%d"))
        except Exception:
            self.status_label['text'] = 'Невірний формат: дата початку'

    def end_minus(self):
        try:
            self.status_label['text'] = ''
            end_date = datetime.strptime(self.end_c.get(), "%Y-%m-%d").date()
            self.end_c.set(monthdelta(end_date, -1).strftime("%Y-%m-%d"))
        except Exception:
            self.status_label['text'] = 'Невірний формат: дата кінця'

    def end_plus(self):
        try:
            self.status_label['text'] = ''
            end_date = datetime.strptime(self.end_c.get(), "%Y-%m-%d").date()
            self.end_c.set(monthdelta(end_date, 1).strftime("%Y-%m-%d"))
        except Exception:
            self.status_label['text'] = 'Невірний формат: дата кінця'

    def set_widgets_disabled(self):
        self.start_entry['state'] = tk.DISABLED
        self.end_entry['state'] = tk.DISABLED
        self.path_report_entry['state'] = tk.DISABLED
        self.report_browsebutton['state'] = tk.DISABLED
        self.path_members_entry['state'] = tk.DISABLED
        self.members_browsebutton['state'] = tk.DISABLED
        self.path_result_entry['state'] = tk.DISABLED
        self.result_browsebutton['state'] = tk.DISABLED
        self.calculate['state'] = tk.DISABLED

    def set_widgets_enabled(self):
        self.start_entry['state'] = 'normal'
        self.end_entry['state'] = 'normal'
        self.path_report_entry['state'] = 'normal'
        self.report_browsebutton['state'] = 'normal'
        self.path_members_entry['state'] = 'normal'
        self.members_browsebutton['state'] = 'normal'
        self.path_result_entry['state'] = 'normal'
        self.result_browsebutton['state'] = 'normal'
        self.calculate['state'] = 'normal'

    def browse_report(self):
        filename = filedialog.askopenfilename(filetypes=(("Файли звітів", "*.xls"), ("Всі файли", "*.*"),))
        self.path_report_c.set(filename)

    def browse_members(self):
        filename = filedialog.askdirectory()
        self.path_members_c.set(filename)

    def browse_result(self):
        filename = filedialog.askdirectory()
        self.path_result_c.set(filename)

    def calculate_reports(self):
        if not self.path_report_c.get().lower().endswith('.xls'):
            self.status_label['text'] = 'Звіт не вибрано!'
            return

        self.set_widgets_disabled()
        self.status_label['text'] = 'Компонується звіт...'
        try:
            start_date = datetime.strptime(self.start_c.get(), "%Y-%m-%d").date()
        except Exception:
            self.status_label['text'] = 'Невірний формат: дата початку'
            return

        try:
            end_date = datetime.strptime(self.end_c.get(), "%Y-%m-%d").date()
        except Exception:
            self.status_label['text'] = 'Невірний формат: дата кінця'
            return

        try:
            self.set_widgets_disabled()
            compose_report(
                start_date,
                end_date,
                self.path_report_c.get(),
                self.path_members_c.get(),
                self.path_result_c.get(),
            )
            self.set_widgets_enabled()
            self.status_label['text'] = 'Звіт успішно скомпоновано'
        except Exception:
            with open("log.txt", "w") as log:
                traceback.print_exc(file=log)

            self.set_widgets_enabled()
            self.status_label['text'] = 'Виникла помилка'


root = tk.Tk()
root.title("Report composer")
app = Application(master=root)
app.mainloop()
