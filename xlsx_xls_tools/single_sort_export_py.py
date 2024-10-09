# sort_export.py

import os
import pandas as pd

def single_sort_export_import(request):

    file_path = request.get("data", {}).get("filePath", "")

    sheet_file = pd.ExcelFile(file_path)                                   # 使用pandas讀取Excel文件
    sheetnames = sheet_file.sheet_names                                    # 獲取所有工作表名稱

    return ['single_sort_export_import', sheetnames]


def single_sort_export_index(request):

    file_path = request.get("data", {}).get("filePath", "")
    sheet_name = request.get("data", {}).get("sheetName", "")

    df = pd.read_excel(file_path, sheet_name=sheet_name)    # 读取指定的工作表
    columns = df.columns.tolist()                           # 获取工作表的列名

    return ['single_sort_export_index', columns]


def single_sort_export_export(request):

    source_file_path = request.get("data", {}).get("filePath", "")
    sheet_name = request.get("data", {}).get("sheet_name", "")
    sort_column = request.get("data", {}).get("sort_column", "")
    secondary_column = request.get("data", {}).get("secondary_column", "")
    value_column = request.get("data", {}).get("value_column", "")

    secondary_column_list = ['名称/项目', '数据行数']

    if sort_column == value_column:
        return ['single_sort_export_no']
    else:
        file_extension = os.path.splitext(source_file_path)[1].lower()
        if file_extension == '.xlsx':
            df = pd.read_excel(source_file_path, sheet_name=sheet_name, engine='openpyxl')
        elif file_extension == '.xls':
            df = pd.read_excel(source_file_path, sheet_name=sheet_name, engine='xlrd')

        # 將指定列的空白行轉換為0
        df[value_column] = df[value_column].replace('', '0')

        # 去除千分位符
        df[value_column] = df[value_column].astype(str).str.replace(',', '')

        # 將字符格式的數字轉換為數值
        df.loc[0:, value_column] = pd.to_numeric(df.loc[0:, value_column], errors='coerce').fillna(0)

        # 进行分类
        sorts = df[sort_column].unique()
        dfs = {sort: df.loc[df[sort_column] == sort] for sort in sorts}
        primary_sorts = df[secondary_column].unique()
        secondary_column_list = secondary_column_list + primary_sorts.tolist()
        secondary_num = len(secondary_column_list)

        # 建立输出文件夹
        export_path = os.path.splitext(source_file_path)[0]
        os.makedirs(export_path, exist_ok=True)

        # 用于存储每个分类的汇总信息
        summary_data = []

        # 输出单个分类项为xlsx文件和进行数据汇总
        for sort, group_df in dfs.items():
            # 输出单个分类项为xlsx文件
            df_export_path = os.path.join(export_path, f'{sort}.xlsx')
            group_df.to_excel(df_export_path, index=False)

            # 进行数据汇总
            secondary_list = [None] * secondary_num
            secondary_list[0] = sort
            secondary_list[1] = len(group_df)

            secondary_sorts = group_df[secondary_column].unique()
            secondary_dfs = {secondary_sort: group_df.loc[group_df[secondary_column] == secondary_sort] for secondary_sort in secondary_sorts}

            for secondary_sort, secondary_group_df in secondary_dfs.items():
                total = secondary_group_df[value_column].sum()
                secondary_index = secondary_column_list.index(secondary_sort)
                secondary_list[secondary_index] = total
            summary_data.append(secondary_list)

        # 输出汇总文件
        summary_df = pd.DataFrame(summary_data, columns=secondary_column_list)
        sorts_export_path = os.path.join(export_path, 'sorts.xlsx')
        summary_df.to_excel(sorts_export_path, index=False)

        return ['single_sort_export_export']