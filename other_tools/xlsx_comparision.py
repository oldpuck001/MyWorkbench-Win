# xlsx_comparision.py

import os
import pandas as pd


def xlsx_comparision_sheetnames(request):

    file = request.get("data", {}).get("file", "")
    file_path = request.get("data", {}).get("file_path", "")

    sheet_file = pd.ExcelFile(file_path)
    sheetnames = sheet_file.sheet_names

    if file == 1:
        return ['xlsx_comparision_sheetnames_1', sheetnames]
    elif file == 2:
        return ['xlsx_comparision_sheetnames_2', sheetnames]


def compare_excels(request):

    xlsx_path_1 = request.get("data", {}).get("xlsx_path_1", "")
    sheet_name_1 = request.get("data", {}).get("sheet_name_1", "")
    file_extension_1 = os.path.splitext(xlsx_path_1)[1].lower()
    if file_extension_1 == '.xlsx':
        df1 = pd.read_excel(xlsx_path_1, sheet_name=sheet_name_1, engine='openpyxl', header=None)
    elif file_extension_1 == '.xls':
        df1 = pd.read_excel(xlsx_path_1, sheet_name=sheet_name_1, engine='xlrd', header=None)

    xlsx_path_2 = request.get("data", {}).get("xlsx_path_2", "")
    sheet_name_2 = request.get("data", {}).get("sheet_name_2", "")
    file_extension_2 = os.path.splitext(xlsx_path_2)[1].lower()
    if file_extension_2 == '.xlsx':
        df2 = pd.read_excel(xlsx_path_2, sheet_name=sheet_name_2, engine='openpyxl', header=None)
    elif file_extension_2 == '.xls':
        df2 = pd.read_excel(xlsx_path_2, sheet_name=sheet_name_2, engine='xlrd', header=None)

    # 對齊資料行列數
    max_rows = max(df1.shape[0], df2.shape[0])
    max_cols = max(df1.shape[1], df2.shape[1])

    diffs = []
    for row in range(max_rows):
        for col in range(max_cols):
            val1 = df1.iat[row, col] if row < df1.shape[0] and col < df1.shape[1] else None
            val2 = df2.iat[row, col] if row < df2.shape[0] and col < df2.shape[1] else None
            if val1 != val2:
                diffs.append({
                    'Row': row + 1,
                    'Column': col + 1,
                    'File1': val1,
                    'File2': val2
                })

    # 顯示差異
    if diffs:

        text = f'发现 {len(diffs)} 个不同之处：\n\n'

        for diff in diffs:

            text += f"第 {diff['Row']} 行, 第 {diff['Column']} 列：File1 = '{diff['File1']}', File2 = '{diff['File2']}'\n"
            text += '\n'

        result_text = {'result_message': text}

    else:

        result_text = {'result_message': '两个表格完全一致。'}

    return ['xlsx_comparison', result_text]