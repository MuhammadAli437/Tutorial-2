# -*- coding: utf-8 -*-
# @Time    : 9/10/2021 11:16 AM
# @Author  : Muhammad.Ali
# @Email   : apncflil@metropolitanwarehouse.com
# @File    : openpyxl_get_col_letter.py
# @Software: PyCharm

import pandas as pd
import all_file_links_module as all_files
import time
from save_default_loc import *
import webbrowser
import os
import openpyxl

# wb = pyxl.load_workbook(r'E:\AutoSave\AutoSave\9-8-2021 11-31-49 AM_Sheet1.xlsx')
#
# ws = wb.active


def get_colz_letter(rel_columns, sheet):
    header = [cell for cell in sheet['A1:XFD1'][0] if
              cell.value is not None and cell.value.strip() != '']  # you get all non-null columns
    # rel_columns = ['amount', 'cost_reconc', 'CPM', 'fuel_surch',
    #                  'calc_cost', 'py_recon', 'mnf_recon']
    cols_to_loop = [cell.column_letter for cell in header if cell.value in rel_columns]  # get column letter

    return cols_to_loop


# rel_columns = ['Size', 'Date Modified']
#
# zz = get_colz_letter(rel_columns, ws)
# print(zz)