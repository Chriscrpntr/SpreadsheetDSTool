import pandas as pd
import openpyxl
import tkinter as tk
from tkinter import filedialog
import os


def hundredS_to_one():
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename()

    old = openpyxl.load_workbook(file_path, data_only=True)
    new = openpyxl.Workbook()
    cell_range = input("What cell range do you want to pull from? (ex. A1:O100)\n")

    all_data = []
    transposeq = input("Do you want to transpose the data? (y/n)\n")
    for sheet in old.worksheets:
        data = sheet[cell_range]
        rows_list = []
        for row in data:
            rows_list.append([cell.value for cell in row])

        if transposeq == 'y':  
            df = pd.DataFrame(rows_list).transpose()
        all_data.append(df)
    final_df = pd.concat(all_data)
    final_df.to_excel('fixed.xlsx', index=False)

def help(option):
    if option == '1':
        print("Option 1: This option does...")
    elif option == '0':
        print("Option 0: This option does...")
    else:
        print("Invalid option")


def main():
    print("""

   _____                          _     _               _     _______          _ 
  / ____|                        | |   | |             | |   |__   __|        | |
 | (___  _ __  _ __ ___  __ _  __| |___| |__   ___  ___| |_     | | ___   ___ | |
  \___ \| '_ \| '__/ _ \/ _` |/ _` / __| '_ \ / _ \/ _ \ __|    | |/ _ \ / _ \| |
  ____) | |_) | | |  __/ (_| | (_| \__ \ | | |  __/  __/ |_     | | (_) | (_) | |
 |_____/| .__/|_|  \___|\__,_|\__,_|___/_| |_|\___|\___|\__|    |_|\___/ \___/|_|
        | |                                                                      
        |_|                                                                      
   
        """)
    choice = input("What do you want to do?\nhelp. example help 1\n1. many to one\n0. exit\n")
    if choice.startswith('help'):
        _, option = choice.split()
        help(option)
    if choice == '1':
        hundredS_to_one()
    if choice == '0':
        os.system('cls' if os.name == 'nt' else 'clear')
        exit()
        
    else:
        os.system('cls' if os.name == 'nt' else 'clear')
        main()
if __name__ == "__main__":
    main()