# preview_balance_sheet.py

import pandas as pd

def get_balance_sheet(balance_path):
    # 读取 CSV 文件
    df = pd.read_csv(balance_path)

    # 格式化数值列，添加千分位分隔符
    df = df.applymap(lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else x)

    # 转置 DataFrame，使其纵向显示（即行变成列，列变成行）
    df_transposed = df.T

    # 增加表头：给第一列添加表头“项目”
    df_transposed.columns = ['上年年末未审数', '上年年末调整数', '上年年末重分类', '上年年末审定数', '上年年末明细数',\
                             '期末未审数', '期末调整数', '期末重分类', '期末审定数', '期末数明细数']
    df_transposed.insert(0, '项目', df_transposed.index)

    # 将 DataFrame 转换为 HTML 表格字符串，并保留表格边框
    content = df_transposed.to_html(header=True, index=False, border=1, justify='left', classes='table', table_id="balance_table")

    # 返回 HTML 内容
    return content