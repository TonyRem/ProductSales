import os
from pathlib import Path
import tkinter as tk
from tkinter import messagebox

from scripts.sales_by_week import format_sales
from scripts.data_preparation import (
    csv_to_excel,
    get_name_with_date,
    create_temp_folder,
    select_csv_file
)

def process_button_click(root, sales_path, stock_path, merge_cells):
    sales_file_path = csv_to_excel(sales_path)
    stock_file_path = csv_to_excel(stock_path)
    result_file_path = get_name_with_date()
    format_sales(
        sales_file_path,
        stock_file_path,
        merge_cells,
        result_file_path
    )

    messagebox.showinfo("Готово", "Обработка данных завершена!")
    open_result_button = tk.Button(
        root,
        text=f"Открыть: {result_file_path.stem}",
        command=lambda: os.system(f'open "{result_file_path}"')
    )
    open_result_button.pack(pady=10)

def main():
    create_temp_folder()
    root = tk.Tk()
    root.title("Обработка данных")

    process_button = tk.Button(
        root,
        text="Обработать данные",
        command=lambda: process_button_click(
            root,
            Path(select_csv_file('Sales')),
            Path(select_csv_file('Stock')),
            merge_cells = False
        )
    )
    process_button.pack(pady=10)
    root.mainloop()

if __name__ == '__main__':
    main()
