# sql_sqlite.py

import os
import pandas as pd

def sql_sqlite_import(request):

    file_path = request.get("data", {}).get("file_path", "")

    sheet_file = pd.ExcelFile(file_path)                                   # 使用pandas讀取Excel文件
    sheetnames = sheet_file.sheet_names                                    # 獲取所有工作表名稱

    return ['sql_sqlite_import', sheetnames]


def sql_sqlite_index(request):

    file_path = request.get("data", {}).get("file_path", "")
    sheet_name = request.get("data", {}).get("sheet_name", "")

    file_extension = os.path.splitext(file_path)[1].lower()

    if file_extension == '.xlsx':
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
    elif file_extension == '.xls':
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='xlrd')

    columns = df.columns.tolist()                           # 获取工作表的列名

    return ['sql_sqlite_index', columns]


def sql_sqlite_export(request):

    file_path = request.get("data", {}).get("file_path", "")
    sheet_name = request.get("data", {}).get("sheet_name", "")
    value_column = request.get("data", {}).get("value_column", "")

    file_extension = os.path.splitext(file_path)[1].lower()

    if file_extension == '.xlsx':
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
    elif file_extension == '.xls':
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='xlrd')

    # 数据清洗
    # 
    df.loc[:, sheet_name] = df[sheet_name].fillna('<空白>').str.strip().replace('', '<空白>')

    # 將指定列的空白行轉換為0
    df[value_column] = df[value_column].replace('', '0')

    # 去除千分位符
    df[value_column] = df[value_column].astype(str).str.replace(',', '')

    # 將字符格式的數字轉換為數值
    df.loc[:, value_column] = pd.to_numeric(df.loc[0:, value_column], errors='coerce').fillna(0)



    return ['sql_sqlite_export']