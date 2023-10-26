import pdfplumber
import pandas as pd
import openpyxl
import os
import json
directory_path = r'C:\Users\TSECDC01\Desktop\SDxResearch\Line Class'
# Replace with the correct path to your Excel file
excel_file_path = 'file_plant_folder-match.xlsx'
# Load the Excel file
workbook = openpyxl.load_workbook(excel_file_path)
sheet = workbook.active
a = 1
while a>0:
    row_index = input("Which file do you want to read?: (letter restricted) ")
    for row in sheet.iter_rows():
        if row_index == row[1].value:
            print("Here is the chart of ",row[1].value)
            chart = json.loads(row[13].value)
            for roww in chart:
                print(roww)
            column_widths = [max(len(str(item)) for item in column) for column in zip(*chart)]
            # Print the data with aligned columns
            for row in chart:
                formatted_row = [str(item).ljust(width) for item, width in zip(row, column_widths)]
                print(" | ".join(formatted_row))
            break
        else:
            continue
    con = input("Do you still want to read next? Y/N ")
    if con == "Y" or con == "y" or con == "yes":
        a = 1
    else:
        a = 0