import csv
import tkinter as tk
from tkinter import filedialog
from collections import Counter
from datetime import date, datetime
from pathlib import Path
from xml.dom import minidom
from zipfile import ZipFile

base_dir = Path('/tmp/docx_dir')
date_from = date(2010, 1, 1)
date_to = date(2030, 1, 1)
csv_filename = base_dir / 'result.csv'


def run():
    files = get_files(base_dir, date_from, date_to)
    editors = get_editors(files)
    print('Dates and editors:')
    for dir_date, editor in editors:
        print(f'{dir_date} {editor}')
    counter = calc_count(editors)
    print('Editors and counts')
    for name, count in counter.items():
        print(f'{name}: {count}')
    print('Saving...')
    save_to_csv(counter, csv_filename)
    print('ok')


def get_files(base_dir, date_from, date_to):
    result = []
    for dir_path in base_dir.iterdir():
        if not dir_path.is_dir():
            continue
        dir_date = get_dir_date(dir_path.name)
        if not (date_from <= dir_date <= date_to):
            continue
        for path in dir_path.iterdir():
            # files on first level
            if path.is_file() and path.suffix == '.docx':
                result.append((dir_date, base_dir/path))
            elif path.is_dir():
                for inner_path in path.iterdir():
                    # files on second level
                    if inner_path.is_file() and inner_path.suffix == '.docx':
                        result.append((dir_date, base_dir/inner_path))
    return result


def get_dir_date(dir_name):
    yy = dir_name[:2]
    yyyy = int(f'20{yy}')
    mm = int(dir_name[2:4])
    dd = int(dir_name[4:6])
    return date(yyyy, mm, dd)


def get_editors(files):
    result = []
    for dir_date, filename in files:
        with ZipFile(filename) as zip_file:
            with zip_file.open('docProps/core.xml') as xml_file:
                last_modified_by = read_xml(xml_file)
                result.append((dir_date, last_modified_by))
    return result


def read_xml(xml_file):
    xmldoc = minidom.parse(xml_file)
    tags = xmldoc.getElementsByTagName('cp:lastModifiedBy')
    tag = tags[0]
    last_modified_by = tag.firstChild.nodeValue
    return last_modified_by


def calc_count(editors):
    names = [name for _, name in editors]
    # first two words from name-string
    names_clear = [' '.join(name.split(' ')[:2]) for name in names]
    counter = Counter(names_clear)
    return counter


def create_csv_filename():
    now_label = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    date_from_label = date_from.isoformat()
    date_to_label = date_to.isoformat()
    filename = f'{now_label}_from_{date_from_label}_to_{date_to_label}.csv'
    csv_filename = base_dir / filename
    return csv_filename


def save_to_csv(counter, csv_filename):
    csv_filename = create_csv_filename()
    with open(csv_filename, mode='w') as csv_file:
        writer = csv.writer(csv_file)
        writer.writerows(counter.items())


class TkApp(tk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.pack()
        self.str_docx_dir = None

        self.btn_ask_dir = tk.Button(text='Choose docx directory')
        self.btn_ask_dir.pack()
        self.btn_ask_dir.bind('<Button-1>', self.open_ask_dir)

        self.label_date_from = tk.Label(text='Date from:')
        self.label_date_from.pack()

        self.str_date_from = tk.StringVar(value='2010-01-15')
        self.entry_date_from = tk.Entry(textvariable=self.str_date_from)
        self.entry_date_from.pack()

        self.label_date_to = tk.Label(text='Date to:')
        self.label_date_to.pack()

        self.str_date_to = tk.StringVar(value='2040-12-31')
        self.entry_date_to = tk.Entry(textvariable=self.str_date_to)
        self.entry_date_to.pack()

        self.btn_count = tk.Button(text='Count and save')
        self.btn_count.pack()
        self.btn_count.bind('<Button-1>', self.count)

        self.str_status = tk.StringVar(value='')
        self.label_status = tk.Label(textvariable=self.str_status)
        self.label_status.pack()

    def open_ask_dir(self, event):
        self.str_docx_dir = filedialog.askdirectory()

    def count(self, event):
        str_date_from = self.str_date_from.get()
        str_date_to = self.str_date_to.get()
        str_docx_dir = self.str_docx_dir
        try:
            date_from = datetime.strptime(str_date_from, '%Y-%m-%d').date()
        except Exception:
            self.str_status.set('Incorrect "Date from"')
            return
        try:
            date_to = datetime.strptime(str_date_to, '%Y-%m-%d').date()
        except Exception:
            self.str_status.set('Incorrect "Date to"')
            return
        try:
            docx_dir = Path(str_docx_dir)
        except Exception:
            self.str_status.set('Incorrect "Docx directory"')
            return
        csv_file = filedialog.asksaveasfile(mode='w', initialfile='123.csv')
        if csv_file is None:
            self.str_status.set('Incorrect "Result filename"')
            return
        self.str_status.set('Working...')
        csv_file.close()
        self.str_status.set('Done')


def gui():
    root = tk.Tk()
    app = TkApp(root)
    app.mainloop()
