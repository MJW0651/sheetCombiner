# Gather and combine .xlsx and .csv files and combine them into a single worksheet/workbook.
# Michael Woolcott

import openpyxl
import csv
import os

def combine_files():
    current_directory = os.getcwd()
    combined_workbook = openpyxl.Workbook()
    combined_sheet = combined_workbook.active

    for filename in os.listdir(current_directory):
        if filename.endswith(".csv") or filename.endswith(".xlsx"):
            with open(filename, 'r') as file:
                if filename.endswith(".csv"):
                    reader = csv.reader(file)
                    data = list(reader)
                elif filename.endswith(".xlsx"):
                    workbook = openpyxl.load_workbook(filename)
                    sheet = workbook.active
                    data = [list(row) for row in sheet.iter_rows(values_only=True)]

            for row in data:
                combined_sheet.append(row)

    combined_workbook.save("combined_data.xlsx")


combine_files()
         
        
