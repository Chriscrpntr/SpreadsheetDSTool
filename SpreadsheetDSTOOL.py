import pandas as pd
import openpyxl
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

file_path = filedialog.askopenfilename()

old = openpyxl.load_workbook(file_path, data_only=True)
new = openpyxl.Workbook()
cell_range = input("What cell range do you want to pull from? (ex. A1:O100)\n")

all_data = []

for sheet in old.worksheets:
    data = sheet[cell_range]
    rows_list = []
    for row in data:
        rows_list.append([cell.value for cell in row])
    transposeq = input("Do you want to transpose the data? (y/n)\n")
    if transposeq == 'y':  
        df = pd.DataFrame(rows_list).transpose()
    all_data.append(df)
final_df = pd.concat(all_data)
final_df.to_excel('fixed.xlsx', index=False)