# generate_chronological_account.py

import os
import pandas as pd

def generate_chronological_account_import(request):

    file_path = request.get("data", {}).get("file_path", "")

    sheet_file = pd.ExcelFile(file_path)                                   # 使用pandas讀取Excel文件
    sheetnames = sheet_file.sheet_names                                    # 獲取所有工作表名稱

    return ['generate_chronological_account_import', sheetnames]

def generate_chronological_account_index(request):

    file_path = request.get("data", {}).get("file_path", "")
    sheet_name = request.get("data", {}).get("sheet_name", "")

    file_extension = os.path.splitext(file_path)[1].lower()

    if file_extension == '.xlsx':
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
    elif file_extension == '.xls':
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='xlrd')

    columns = df.columns.tolist()                           # 获取工作表的列名

    return ['generate_chronological_account_index', columns]

def generate_chronological_account_debit_or_credit(request):

    file_path = request.get("data", {}).get("file_path", "")
    sheet_name = request.get("data", {}).get("sheet_name", "")
    column_name = request.get("data", {}).get("column_name", "")

    debit_or_credit_list_cleaned = []

    file_extension = os.path.splitext(file_path)[1].lower()

    if file_extension == '.xlsx':
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
    elif file_extension == '.xls':
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='xlrd')

    debit_or_credit_np = df[column_name].unique()
    debit_or_credit_list = debit_or_credit_np.tolist()

    for item in debit_or_credit_list:
        if not isinstance(item, (int, float)):
            debit_or_credit_list_cleaned.append(item.replace('\t', ''))
        else:
            debit_or_credit_list_cleaned.append(item)

    return ['generate_chronological_account_debit_or_credit', debit_or_credit_list_cleaned]

def generate_chronological_account_export(request):

    file_path = request.get("data", {}).get("file_path", "")
    save_path = request.get("data", {}).get("save_path", "")
    sheet_name = request.get("data", {}).get("sheet_name", "")
    number_name = request.get("data", {}).get("number_column", "")
    debit_or_credit_column = request.get("data", {}).get("debit_or_credit_column", "")
    credit_column = request.get("data", {}).get("credit_column", "")
    debit_column = request.get("data", {}).get("debit_column", "")
    date_column = request.get("data", {}).get("date_column", "")
    name_column = request.get("data", {}).get("name_column", "")
    summary_column = request.get("data", {}).get("summary_column", "")
    currency_column = request.get("data", {}).get("currency_column", "")
    value_column = request.get("data", {}).get("value_column", "")
    yszk_name = request.get("data", {}).get("yszk_name", "")
    yszk_name_list = [name.strip() for name in yszk_name.split(',')]
    yfzk_name = request.get("data", {}).get("yfzk_name", "")
    yfzk_name_list = [name.strip() for name in yfzk_name.split(',')]

    # 讀取 Excel 文件
    file_extension = os.path.splitext(file_path)[1].lower()

    if file_extension == '.xlsx':
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
    elif file_extension == '.xls':
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='xlrd')

    # 創建一個空列表來存儲最終的 DataFrame
    new_rows = []

    for index, row in df.iterrows():

        if str(row[debit_or_credit_column]) == str(credit_column):

            if row[name_column] in yszk_name_list:
                level_1_subject = '应收帐款'
            elif row[name_column] in yfzk_name_list:
                level_1_subject = '应付帐款'
            else:
                level_1_subject = '其他应收款'

            new_row = [row[date_column], None, '银行存款', row[number_name], row[summary_column], row[currency_column], row[value_column], None]
            new_row += row.tolist()

            new_rows.append(new_row)

            new_row = []

            new_row = [row[date_column], None, level_1_subject, row[name_column], row[summary_column], row[currency_column], None, row[value_column]]

            new_rows.append(new_row)

            new_row = []
            
        elif str(row[debit_or_credit_column]) == str(debit_column):

            if row[name_column] in yszk_name_list:
                level_1_subject = '应收帐款'
            elif row[name_column] in yfzk_name_list:
                level_1_subject = '应付帐款'
            else:
                level_1_subject = '其他应收款'

            new_row = [row[date_column], None, level_1_subject, row[name_column], row[summary_column], row[currency_column], row[value_column], None]

            new_rows.append(new_row)

            new_row = []

            new_row = [row[date_column], None, '银行存款', row[number_name], row[summary_column], row[currency_column], None, row[value_column]]
            new_row += row.tolist()

            new_rows.append(new_row)

            new_row = []

    # 將 new_rows 轉換為新的 DataFrame
    columns = ['记账日期', '凭证字号', '一级科目', '二级科目', '摘要', '币种', '借方金额', '贷方金额'] + df.columns.tolist()
    new_df = pd.DataFrame(new_rows, columns=columns)

    # 將結果保存為新的 Excel 文件
    new_df.to_excel(save_path, index=False)

    return ['generate_chronological_account_export']