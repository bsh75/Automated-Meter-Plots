import xlrd
import openpyxl
from openpyxl import Workbook
from datetime import datetime, timedelta
import os
import matplotlib.pyplot as plt
import struct
import calendar
from openpyxl.styles import Font, PatternFill
import pandas as pd

folder_path = 'Aug2023/Electric Plot Data'

folder_path = 'Aug2023/Electric Plot Data'
basement_f = f'{folder_path}/Basement.xlsx'
basement_df = pd.read_excel(basement_f)
common_f = f'{folder_path}/Common Areas.xlsx'
common_df = pd.read_excel(common_f)
low_levels_f = f'{folder_path}/Level 01-08.xlsx'
low_levels_df = pd.read_excel(low_levels_f)
high_levels_f = f'{folder_path}/Level 09-18.xlsx'
high_levels_df = pd.read_excel(high_levels_f)


meter_groups = {
    'Basement': ['D44', 'D43', 'D55'],
    'Common Areas': ['D1', 'D2', 'D3', 'D4', 'D5', 'D6', 'D27', 'D31', 'D11', 'D20', 'D15', 'D10', 'D14', 'D17', 'D18', 'D29', 'D41', 'D42', 'D30', 'D9', 'D13', 'D28', 'D22', 'D26', 'D16' 'D19', 'D12'],            
    'Level 01-08': ['D40', 'D59', 'D39', 'D38', 'D37', 'D36', 'D63', 'D35', 'D60', 'D57', 'D62', 'D68', 'D65', 'D34', 'D33', 'D61', 'D58', 'D66', 'D67', 'D64', 'D32', 'D7', 'D8'],
    'Level 09-18': [
        'D54','D53','D52','D51','D50','D23','D24','D49','D48','D47','D46','D45','D25']
}
meter_groups_data = {
    'Basement': basement_df,
    'Common Areas': common_df,
    'Level 01-08': low_levels_df,
    'Level 09-18': high_levels_df
}

for filename in os.listdir(folder_path):
    if filename.endswith('.xlsx'):
        file_path = os.path.join(folder_path, filename)
        df = pd.read_excel(file_path)
        for index, row in df.iterrows():
            print('nan')
            # print(f"index = {index}: row = {row}")


        # # print(subset)
        # for group, meters in meter_groups.items():
        #     for meter in meters:
        #         if meter in filename:
        #             # meter_groups_data[group] += df
        #             for row in subset:
        #                 print(row)
        #                 for cell in row:
        #                     print(cell)
        #                     # # print(type)
        #                     # print(f'Row: {index}, Column: {column}, Value: {type(value)}')
        #                     # # if value 
        #                     # meter_groups_data[group][index][column] += value


