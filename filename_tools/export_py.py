# export_py.py

import os
import pandas as pd
import openpyxl

def export(request):

    folderPath = request.get("data", {}).get("folderPath", "")
    savePath = request.get("data", {}).get("savePath", "")

    filenameList = []
    filepathList = []

    for filename in os.listdir(folderPath):
        filenamepath = os.path.join(folderPath, filename)
        filenameList.append(filename)
        filepathList.append('file://' + filenamepath)

    # 使用pandas创建一个DataFrame
    df = pd.DataFrame({
        '文件名': filenameList,
        '文件路径': filepathList
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

    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                # Simple check for common URL patterns
                if cell.value.startswith("http://") or cell.value.startswith("https://") or cell.value.startswith("file://"):
                    cell.hyperlink = cell.value  # Convert to hyperlink
                    cell.style = "Hyperlink"  # Apply hyperlink style

    # 保存Excel文件
    wb.save(savePath)

    return ['filename_export']