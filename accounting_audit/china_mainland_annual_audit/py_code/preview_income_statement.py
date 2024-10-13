# preview_income_statement.py

import pandas as pd

def get_income_statement(income_path):
    # 读取 CSV 文件
    df = pd.read_csv(income_path)

    # 格式化数值列，添加千分位分隔符
    df = df.applymap(lambda x: f"{x:,.2f}" if isinstance(x, (int, float)) else x)

    # 转置 DataFrame，使其纵向显示（即行变成列，列变成行）
    df_transposed = df.T

    # 增加表头：给第一列添加表头“项目”
    df_transposed.columns = ['上期金额', '本期未审数', '本期调整数', '本期重分类', '本期审定数', '本期明细数']
    df_transposed.insert(0, '项目', df_transposed.index)

    # 将 DataFrame 转换为 HTML 表格字符串，并保留表格边框
    content = df_transposed.to_html(header=True, index=False, border=1, justify='left', classes='table', table_id="income_table")

    # 返回 HTML 内容
    return content