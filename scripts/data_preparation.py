import csv
import datetime
from openpyxl import Workbook
from pathlib import Path
from tkinter import filedialog

def select_csv_file(file_name):
    file_path = filedialog.askopenfilename(
        filetypes=[('CSV Files', '*.csv')],
        title=f"Выберите файл: {file_name}"
    )
    
    return file_path

def create_temp_folder():
    temp_folder = Path('temp')
    if not temp_folder.exists():
        temp_folder.mkdir()

def csv_to_excel(csv_path):
    excel_file = f'temp/{csv_path.stem}.xlsx'
    workbook = Workbook()
    sheet = workbook.active

    with open(csv_path, 'r', encoding='utf-8') as file:
        csv_reader = csv.reader(file)
        for row in csv_reader:
            sheet.append(row)

    workbook.save(excel_file)
    
    return excel_file

def get_name_with_date():
    now = datetime.datetime.now()
    formatted_date = now.strftime("%d.%m_%H.%M")
    file_name = Path(f'temp/SalesByWeek_{formatted_date}.xlsx')
    return file_name