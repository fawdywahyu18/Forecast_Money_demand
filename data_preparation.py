import openpyxl
import pandas as pd
from pandas import ExcelWriter
import numpy as np

def prep_file(filepath, sheetName, flow_type):
    df = pd.read_excel(filepath, sheet_name=sheetName)
    col = 2
    rapi_all = pd.DataFrame()
    col_rapi = ('uk_100k', 'uk_50k', 'uk_20k', 'uk_10k', 'uk_5k', 'uk_2k', 'uk_1k',
                'uk_0.5k', 'uk_0.1k', 'uk<0.1k', 'uang_kertas', 'ul_1k', 'ul_0.5k',
                'ul_0.2k', 'ul_0.1k', 'ul_50', 'ul_10', 'ul_5', 'ul<5', 'uang_logam',
                flow_type)
    test_loop = len(col_rapi) + col
    while col < test_loop:
        i = 3
        result_all = pd.DataFrame()
        target_col = df[f'Unnamed: {col}']
        end_point = len(target_col)
        while i <= end_point:
            result = target_col[i:(i+12)]
            result_all = pd.concat([result_all, result])
            i += 13
        rapi_all[col_rapi[(col-2)]] = result_all
        col += 1

    rapi_all['date_index'] = pd.date_range(start='1994-01-01', periods=len(rapi_all), freq='MS')
    rapi_all.set_index('date_index', inplace=True)
    rapi_all = rapi_all.loc[(rapi_all.index>='2010-01-01')&(rapi_all.index<='2021-03-01')]
    return rapi_all

source_file = './DPU2021/INFLOW 1994 - 2021.xlsx' # Setting your filename here
wb = openpyxl.load_workbook(source_file)
sheet_names = wb.sheetnames

export_name = './DPU2021/inflow_indonesia.xlsx' # Setting your filename to export here
writer = ExcelWriter(export_name)
for sheet_n in sheet_names:
    prepared = prep_file(filepath=source_file, 
                     sheetName=sheet_n, 
                     flow_type='Inflow') # Specify your type of Outflow/Inflow
    prepared.to_excel(writer, sheet_name=sheet_n)
    progress = sheet_names.index(sheet_n)/len(sheet_names)
    print(np.round(progress, 2)*100)
writer.save()