# data_cleaning.py

# 数据清洗：对获取的数据进行清洗，确保其格式统一、无重复项。

import os
import pandas as pd

def data_cleaning_import(request):

    file_path = request.get("data", {}).get("file_path", "")

    sheet_file = pd.ExcelFile(file_path)                    # 使用pandas讀取Excel文件
    sheetnames = sheet_file.sheet_names                     # 獲取所有工作表名稱

    return ['data_cleaning_import', sheetnames]


def data_cleaning_index(request):

    file_path = request.get("data", {}).get("file_path", "")
    sheet_name = request.get("data", {}).get("sheet_name", "")

    file_extension = os.path.splitext(file_path)[1].lower()

    if file_extension == '.xlsx':
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
    elif file_extension == '.xls':
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='xlrd')

    columns = df.columns.tolist()                           # 获取工作表的列名

    return ['data_cleaning_index', columns]


def data_cleaning_export(request):

    file_path = request.get("data", {}).get("file_path", "")
    save_path = request.get("data", {}).get("save_path", "")
    sheet_name = request.get("data", {}).get("sheet_name", "")
    column_name = request.get("data", {}).get("column_name", "")
    cleaning_mode = request.get("data", {}).get("cleaning_mode", "")

    file_extension = os.path.splitext(file_path)[1].lower()

    if file_extension == '.xlsx':
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
    elif file_extension == '.xls':
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='xlrd')

    if cleaning_mode == 'remove_duplicates':
        df = df.drop_duplicates(subset=[column_name])               # 删除指定列的重复行

    elif cleaning_mode == 'fill_missing_zero':
        df[column_name] = df[column_name].fillna(0)                 # 填充缺失值为 0

    elif cleaning_mode == 'fill_missing_blank':
        df[column_name] = df[column_name].fillna('<空白>')          # 填充缺失值为 '<空白>'

    elif cleaning_mode == 'standardize_text':
        df[column_name] = df[column_name].str.strip().str.lower()

    elif cleaning_mode == 'convert_data_str':
        df[column_name] = df[column_name].astype(str)

    elif cleaning_mode == 'convert_data_int':
        df[column_name] = df[column_name].astype(str).str.replace(',', '')
        df[column_name] = pd.to_numeric(df[column_name], errors='coerce').fillna(0).astype(int) # 转换为整数，支持缺失值

    elif cleaning_mode == 'convert_data_float':
        df[column_name] = df[column_name].astype(str).str.replace(',', '')
        df[column_name] = pd.to_numeric(df[column_name], errors='coerce').astype(float)         # 转换为浮点数，支持缺失值

    elif cleaning_mode == 'convert_data_date':
        df[column_name] = pd.to_datetime(df[column_name], errors='coerce')

    elif cleaning_mode == 'drop_columns':
        df = df.drop(columns=[column_name])

    df.to_excel(save_path, index=False)

    return ['data_cleaning_export', ['导出成功！']]