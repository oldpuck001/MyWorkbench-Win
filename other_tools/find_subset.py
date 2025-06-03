#find_subset.py

import os
import pandas as pd

def find_subset_sheetnames_import(request):

    file_path = request.get("data", {}).get("file_path", "")

    sheet_file = pd.ExcelFile(file_path)
    sheetnames = sheet_file.sheet_names

    return ['find_subset_sheetnames_import', sheetnames]


def find_subset_columns_index(request):

    file_path = request.get("data", {}).get("file_path", "")
    sheet_name = request.get("data", {}).get("sheet_name", "")

    file_extension = os.path.splitext(file_path)[1].lower()

    if file_extension == '.xlsx':
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
    elif file_extension == '.xls':
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='xlrd')

    columns = df.columns.tolist()

    return ['find_subset_columns_index', columns]


def find_subset_import(request):

    file_path = request.get("data", {}).get("file_path", "")
    sheet_name = request.get("data", {}).get("sheet_name", "")
    target_value = request.get("data", {}).get("target_value", "")
    value_name = request.get("data", {}).get("value_name", "")
    value_num = request.get("data", {}).get("value_num", "")

    file_extension = os.path.splitext(file_path)[1].lower()

    if file_extension == '.xlsx':
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
    elif file_extension == '.xls':
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='xlrd')

    df = df.fillna('N/A')
    df_rename = df.rename(columns={target_value: 'target_value', value_name: 'name_value', value_num: 'num_value'})
    data_import = df_rename.to_dict(orient='dict')

    return ['find_subset_import', data_import]


def find_subset_export(request):

    export_file_path = request.get("data", {}).get("savePath", "")
    target_value = request.get("data", {}).get("target_value", 0.0)
    value_name = request.get("data", {}).get("value_name", [])
    value_num = request.get("data", {}).get("value_num", [])

    # 转换目标值为 float
    target_value = float(str(target_value).replace(',', '').strip() or 0)

    # 清洗 value_num，把千分位字符串转为 float，空白或非法输入设为 0.0
    cleaned_value_num = []
    for val in value_num:
        try:
            num = float(str(val).replace(',', '').strip())
        except (ValueError, TypeError):
            num = 0.0
        cleaned_value_num.append(num)

    # 仅剔除数值为 0 的项，保留名称为空的项
    valid_pairs = [
        (name, num)
        for name, num in zip(value_name, cleaned_value_num)
        if num != 0.0
    ]

    # 若没有有效数据，直接导出提示
    if not valid_pairs:
        df_empty = pd.DataFrame([["无有效数据（所有数值均为 0）"]], columns=["提示"])
        df_empty.to_excel(export_file_path, index=False)
        result_text = {'result_message': f'无有效数据（所有数值均为 0），已保存提示信息到：{export_file_path}'}
        return ['find_subset_export', result_text]

    tolerance = 0.01

    dp = {0.0: [[]]}
    results = []

    for name, num in valid_pairs:
        new_dp = {}
        for current_sum in dp:
            new_sum = current_sum + num
            subset_list = dp[current_sum]
            for subset in subset_list:
                new_subset = subset + [(name, num)]
                if abs(new_sum - target_value) <= tolerance:
                    results.append(new_subset)
                if new_sum not in new_dp:
                    new_dp[new_sum] = []
                new_dp[new_sum].append(new_subset)
        for k, v in new_dp.items():
            if k not in dp:
                dp[k] = []
            dp[k].extend(v)

    if not results:
        df_empty = pd.DataFrame([["未找到满足条件的组合"]], columns=["提示"])
        df_empty.to_excel(export_file_path, index=False)
        result_text = {'result_message': f'未找到满足条件的组合，已保存提示信息到：{export_file_path}'}
        return ['find_subset_export', result_text]

    # 输出结果，每组加小计和空行
    output_data = []
    for idx, subset in enumerate(results, 1):
        total = 0.0
        for name, num in subset:
            output_data.append({
                "组合编号": f"组合 {idx}",
                "名称": name,
                "数值": round(num, 2)
            })
            total += num
        output_data.append({
            "组合编号": f"组合 {idx}",
            "名称": "小计",
            "数值": round(total, 2)
        })
        output_data.append({})  # 空行分隔

    df_result = pd.DataFrame(output_data)

    df_result.to_excel(export_file_path, index=False)

    result_text = {'result_message': f'已导出 {len(results)} 个满足条件的组合到：{export_file_path}'}
    return ['find_subset_export', result_text]