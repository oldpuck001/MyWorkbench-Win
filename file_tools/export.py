# export.py

import os
import re
import pandas as pd
import openpyxl

def export(request):

    folderPath = request.get('data', {}).get('folderPath', '')
    savePath = request.get('data', {}).get('savePath', '')
    search_subfolders = request.get('data', {}).get('yes_or_no', 'yes') == 'yes'

    # 定义递归函数，用于遍历文件夹及其子文件夹
    def get_files_recursive(folder_path):
        file_list = []
        folder_list = []
        for root, dirs, files in os.walk(folder_path):
            # 添加当前文件夹的信息
            if root != folder_path:  # 排除根文件夹
                folder_list.append((os.path.basename(root), root, '文件夹'))
            # 添加当前文件夹中的文件信息
            for file in files:
                file_path = os.path.join(root, file)
                file_list.append((file, file_path, '文件'))
            if not search_subfolders:
                break  # 如果不检索子文件夹，则只遍历当前文件夹
        return folder_list + file_list  # 返回文件夹和文件的组合列表

    # 获取所有文件及其路径
    items = get_files_recursive(folderPath)

    # 将信息存储到列表中
    nameList = [item[0] for item in items]  # 名称（文件夹名或文件名）
    pathList = [item[1] for item in items]  # 路径（文件夹路径或文件路径）
    typeList = [item[2] for item in items]  # 类型（文件夹或文件）

    # 使用pandas创建一个DataFrame
    df = pd.DataFrame({
        '文件名': nameList,
        '文件路径': pathList,
        '类型': typeList
    })

    # 创建一个Excel工作簿
    wb = openpyxl.Workbook()
    ws = wb.active

    # 将DataFrame的数据添加到Excel工作表
    for r in openpyxl.utils.dataframe.dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    # 调整列宽
    for column_cells in ws.columns:
        length = max(len(str(cell.value))+10 for cell in column_cells)
        ws.column_dimensions[column_cells[0].column_letter].width = length

    # 设置链接
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                if cell.value.startswith('http://') or cell.value.startswith('https://') or re.match(r'^[C-Z]:', cell.value, re.IGNORECASE):
                    cell.hyperlink = cell.value
                    cell.style = 'Hyperlink'

    # 保存Excel文件
    wb.save(savePath)

    return ['filename_export']