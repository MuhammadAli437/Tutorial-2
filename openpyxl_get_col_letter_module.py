import pandas as pd
import time
import webbrowser
import os
import openpyxl

def get_colz_letter(rel_columns, sheet):
    header = [cell for cell in sheet['A1:XFD1'][0] if
              cell.value is not None and cell.value.strip() != '']  # you get all non-null columns
    cols_to_loop = [cell.column_letter for cell in header if cell.value in rel_columns]  # get column letter

    return cols_to_loop
