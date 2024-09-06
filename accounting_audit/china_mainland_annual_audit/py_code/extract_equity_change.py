# extract_equity_change.py

import os
import re
import pandas as pd

def extract_equity_change_xlsx(filepath, data_folder):
    equity_change_list = ['上年年末余额', '会计政策变更', '前期差错更正', '其他',\
                            '本年年初余额',\
                            '本年增减变动金额',\
                            '综合收益总额',\
                            '所有者投入和减少资本', '所有者投入的普通股', '其他权益工具持有者投入资本', '股份支付计入所有者权益的金额',\
                            '所有者投入和减少资本：其他',\
                            '利润分配', '提取盈余公积', '提取一般风险准备*', '对所有者（或股东）的分配', '利润分配：其他',\
                            '所有者权益内部结转', '资本公积转增资本（或股本）', '盈余公积转增资本（或股本）', '盈余公积弥补亏损',\
                            '设定受益计划变动额结转留存收益', '其他综合收益结转留存收益', '所有者权益内部结转：其他',\
                            '本年年末余额']

    equity_change_dict = {key: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0] for key in equity_change_list}

    equity_change_regex_dict = {'上年年末余额': re.compile(r'\S*上年年末余额\S*'),\
                                '会计政策变更': re.compile(r'\S*会计政策变更\S*'),\
                                '前期差错更正': re.compile(r'\S*前期差错更正\S*'),\
                                '本年年初余额': re.compile(r'\S*本年年初余额\S*'),\
                                '本年增减变动金额': re.compile(r'\S*本年增减变动金额\S*'),\
                                '综合收益总额': re.compile(r'\S*综合收益总额\S*'),\
                                '所有者投入和减少资本': re.compile(r'\S*所有者投入和减少资本\S*'),\
                                '所有者投入的普通股': re.compile(r'\S*所有者投入的普通股\S*|S*所有者投入的资本\S*'),\
                                '其他权益工具持有者投入资本': re.compile(r'\S*其他权益工具持有者投入资本\S*'),\
                                '股份支付计入所有者权益的金额': re.compile(r'\S*股份支付计入所有者权益的金额\S*'),\
                                '利润分配': re.compile(r'\S*利润分配\S*'),\
                                '提取盈余公积': re.compile(r'\S*提取盈余公积\S*'),\
                                '提取一般风险准备*': re.compile(r'\S*提取一般风险准备\S*'),\
                                '对所有者（或股东）的分配': re.compile(r'\S*对所有者\S{1,2}股东\S{1}的分配\S*|\S*对所有者的分配\S*|\S*对股东的分配\S*'),\
                                '所有者权益内部结转': re.compile(r'\S*所有者权益内部结转\S*'),\
                                '资本公积转增资本（或股本）': re.compile(r'\S*资本公积转增资本\S{1,2}股本\S*|\S*资本公积转增资本\S*|\S*资本公积转增股本\S*'),\
                                '盈余公积转增资本（或股本）': re.compile(r'\S*盈余公积转增资本\S{1,2}股本\S*|\S*盈余公积转增资本\S*|\S*盈余公积转增股本\S*'),\
                                '盈余公积弥补亏损': re.compile(r'\S*盈余公积弥补亏损\S*'),\
                                '设定受益计划变动额结转留存收益': re.compile(r'\S*设定受益计划变动额结转留存收益\S*'),\
                                '其他综合收益结转留存收益': re.compile(r'\S*其他综合收益结转留存收益\S*'),\
                                '本年年末余额': re.compile(r'\S*本年年末余额\S*')}

    display_equity_dict = {}

    num = 0
    
    file_extension = os.path.splitext(filepath)[1].lower()
    if file_extension == '.xlsx':
        equity_change_df = pd.read_excel(filepath, engine='openpyxl')
    elif file_extension == '.xls':
        equity_change_df = pd.read_excel(filepath, engine='xlrd')

    # 检查文件有多少列及表头所在的行次
    num_columns = equity_change_df.shape[1]

    # 使用正则表达式来匹配前后和中间可能存在空格的“项目”关键字
    pattern_project_equity = re.compile(r'\s*项\s*目\s*')
    # 查找含有“项目”的行
    project_row_index = equity_change_df[equity_change_df.astype(str).apply(lambda row: row.str.contains(pattern_project_equity).any(), axis=1)].index[0]

    # 正则表达式匹配“实收资本(或股本)”关键字
    pattern_capital_equity = re.compile(r'\s*实\s*收\s*资\s*本\s*股\s*本\s*|\s*实\s*收\s*资\s*本\s*|\s*股\s*本\s*')
    # 查找含有“实收资本(或股本)”的行
    capital_row_index = equity_change_df[equity_change_df.astype(str).apply(lambda row: row.str.contains(pattern_capital_equity).any(), axis=1)].index[0]

    # 使用正则表达式来匹配前后和中间可能存在空格的“优先股”关键字
    pattern_preferred_equity = re.compile(r'\s*优\s*先\s*股\s*')
    # 查找含有“优先股”的行
    preferred_row_index = equity_change_df[equity_change_df.astype(str).apply(lambda row: row.str.contains(pattern_preferred_equity).any(), axis=1)].index[0]

    # 使用正则表达式来匹配前后和中间可能存在空格的“永续债”关键字
    pattern_perpetual_equity = re.compile(r'\s*永\s*续\s*债\s*')
    # 查找含有“永续债”的行
    perpetual_row_index = equity_change_df[equity_change_df.astype(str).apply(lambda row: row.str.contains(pattern_perpetual_equity).any(), axis=1)].index[0]

    # 使用正则表达式来匹配前后和中间可能存在空格的“其他”关键字
    pattern_other_equity = re.compile(r'^其\s*他$|^其\s*它$')
    # 查找含有“其他”的行
    other_row_index = equity_change_df[equity_change_df.astype(str).apply(lambda row: row.str.contains(pattern_other_equity).any(), axis=1)].index[0]

    # 使用正则表达式来匹配前后和中间可能存在空格的“资本公积”关键字
    pattern_reserve_equity = re.compile(r'\s*资\s*本\s*公\s*积\s*')
    # 查找含有“资本公积”的行
    reserve_row_index = equity_change_df[equity_change_df.astype(str).apply(lambda row: row.str.contains(pattern_reserve_equity).any(), axis=1)].index[0]

    # 使用正则表达式来匹配前后和中间可能存在空格的“减：库存股”关键字
    pattern_stock_equity = re.compile(r'\s*\S*库\s*存\s*股\s*')
    # 查找含有“减：库存股”的行
    stock_row_index = equity_change_df[equity_change_df.astype(str).apply(lambda row: row.str.contains(pattern_stock_equity).any(), axis=1)].index[0]

    # 使用正则表达式来匹配前后和中间可能存在空格的“其他综合收益”关键字
    pattern_income_equity = re.compile(r'\s*其\s*他\s*综\s*合\s*收\s*益\s*')
    # 查找含有“其他综合收益”的行
    income_row_index = equity_change_df[equity_change_df.astype(str).apply(lambda row: row.str.contains(pattern_income_equity).any(), axis=1)].index[0]

    # 使用正则表达式来匹配前后和中间可能存在空格的“专项储备”关键字
    pattern_special_equity = re.compile(r'\s*专\s*项\s*储\s*备\s*')
    # 查找含有“专项储备”的行
    special_row_index = equity_change_df[equity_change_df.astype(str).apply(lambda row: row.str.contains(pattern_special_equity).any(), axis=1)].index[0]

    # 使用正则表达式来匹配前后和中间可能存在空格的“盈余公积”关键字
    pattern_surplus_equity = re.compile(r'\s*盈\s*余\s*公\s*积\s*')
    # 查找含有“盈余公积”的行
    surplus_row_index = equity_change_df[equity_change_df.astype(str).apply(lambda row: row.str.contains(pattern_surplus_equity).any(), axis=1)].index[0]

    # 使用正则表达式来匹配前后和中间可能存在空格的“一般风险准备”关键字
    pattern_risk_equity = re.compile(r'\s*一\s*般\s*风\s*险\s*准\s*备\s*')
    # 查找含有“一般风险准备”的行
    risk_row_index = equity_change_df[equity_change_df.astype(str).apply(lambda row: row.str.contains(pattern_risk_equity).any(), axis=1)].index[0]

    # 使用正则表达式来匹配前后和中间可能存在空格的“未分配利润”关键字
    pattern_profit_equity = re.compile(r'\s*未\s*分\s*配\s*利\s*润\s*')
    # 查找含有“未分配利润”的行
    profit_row_index = equity_change_df[equity_change_df.astype(str).apply(lambda row: row.str.contains(pattern_profit_equity).any(), axis=1)].index[0]

    # 使用正则表达式来匹配前后和中间可能存在空格的“小计”关键字
    pattern_subtotal_equity = re.compile(r'\s*小\s*计\s*')
    # 查找含有“小计”的行
    subtotal_row_index = equity_change_df[equity_change_df.astype(str).apply(lambda row: row.str.contains(pattern_subtotal_equity).any(), axis=1)].index[0]

    # 使用正则表达式来匹配前后和中间可能存在空格的“少数股东权益”关键字
    pattern_minority_equity = re.compile(r'\s*少\s*数\s*股\s*东\s*权\s*益\s*')
    # 查找含有“少数股东权益”的行
    minority_row_index = equity_change_df[equity_change_df.astype(str).apply(lambda row: row.str.contains(pattern_minority_equity).any(), axis=1)].index[0]

    # 使用正则表达式来匹配前后和中间可能存在空格的“所有者权益合计”关键字
    pattern_total_equity = re.compile(r'\s*所\s*有\s*者\s*权\s*益\s*合\s*计\s*')
    # 查找含有“所有者权益合计”的行
    total_row_index = equity_change_df[equity_change_df.astype(str).apply(lambda row: row.str.contains(pattern_total_equity).any(), axis=1)].index[0]

    min_row_index = min(project_row_index,\
                        capital_row_index, preferred_row_index, perpetual_row_index, other_row_index,\
                        reserve_row_index, stock_row_index, income_row_index, special_row_index,\
                        surplus_row_index, risk_row_index, profit_row_index, subtotal_row_index,\
                        minority_row_index, total_row_index)

    start_row_index = max(project_row_index,\
                            capital_row_index, preferred_row_index, perpetual_row_index, other_row_index,\
                            reserve_row_index, stock_row_index, income_row_index, special_row_index,\
                            surplus_row_index, risk_row_index, profit_row_index, subtotal_row_index,\
                            minority_row_index, total_row_index)

    # 列表来保存找到的列索引
    project_equity_columns = []
    capital_equity_columns = []
    preferred_equity_columns = []
    perpetual_equity_columns = []
    other_equity_columns = []
    reserve_equity_columns = []
    stock_equity_columns = []
    income_equity_columns = []
    special_equity_columns = []
    surplus_equity_columns = []
    risk_equity_columns = []
    profit_equity_columns = []
    subtotal_equity_columns = []
    minority_equity_columns = []
    total_equity_columns = []

    # 遍历列，寻找匹配的列索引
    for m in range(num_columns):
        for n in range(min_row_index, start_row_index+1):
            header = str(equity_change_df.iloc[n, m])
            if re.search(pattern_project_equity, header):
                project_equity_columns.append(m)
            elif re.search(pattern_capital_equity, header):
                capital_equity_columns.append(m)
            elif re.search(pattern_preferred_equity, header):
                preferred_equity_columns.append(m)
            elif re.search(pattern_perpetual_equity, header):
                perpetual_equity_columns.append(m)
            elif re.search(pattern_other_equity, header):
                other_equity_columns.append(m)
            elif re.search(pattern_reserve_equity, header):
                reserve_equity_columns.append(m)
            elif re.search(pattern_stock_equity, header):
                stock_equity_columns.append(m)
            elif re.search(pattern_income_equity, header):
                income_equity_columns.append(m)
            elif re.search(pattern_special_equity, header):
                special_equity_columns.append(m)
            elif re.search(pattern_surplus_equity, header):
                surplus_equity_columns.append(m)
            elif re.search(pattern_risk_equity, header):
                risk_equity_columns.append(m)
            elif re.search(pattern_profit_equity, header):
                profit_equity_columns.append(m)
            elif re.search(pattern_subtotal_equity, header):
                subtotal_equity_columns.append(m)
            elif re.search(pattern_minority_equity, header):
                minority_equity_columns.append(m)
            elif re.search(pattern_total_equity, header):
                total_equity_columns.append(m)

    for n in project_equity_columns:
        index = project_equity_columns.index(n)
        for i in range(start_row_index + 1, len(equity_change_df)):
            capital = 0
            capital_0 = 0
            preferred = 0
            preferred_0 = 0
            perpetual = 0
            perpetual_0 = 0
            other = 0
            other_0 = 0
            reserve = 0
            reserve_0 = 0
            stock = 0
            stock_0 = 0
            income = 0
            income_0 = 0
            special = 0
            special_0 = 0
            surplus = 0
            surplus_0 = 0
            risk = 0
            risk_0 = 0
            profit = 0
            profit_0 = 0
            subtotal = 0
            subtotal_0 = 0
            minority = 0
            minority_0 = 0
            total = 0
            total_0 = 0
            keywords = equity_change_df.iloc[i, n]
            if isinstance(keywords, str):
                keywords = keywords.replace(' ', '')
            if isinstance(keywords, float):
                continue
            for key in equity_change_list:
                if key in ['其他', '所有者投入和减少资本：其他', '利润分配：其他', '所有者权益内部结转：其他']:
                    if re.search(re.compile(r'\S*其他\S*|\S*其它\S*'), keywords):
                        if equity_change_df.iloc[i-3, n].replace(' ', '') == '上年年末余额':
                            key = '其他'
                        elif equity_change_df.iloc[i-4, n].replace(' ', '') != '所有者投入和减少资本':
                            key = '所有者投入和减少资本：其他'
                        elif equity_change_df.iloc[i-4, n].replace(' ', '') != '利润分配':
                            key = '利润分配：其他'
                        elif equity_change_df.iloc[i-6, n].replace(' ', '') != '所有者权益内部结转':
                            key = '所有者权益内部结转：其他'
                        capital = equity_change_df.iloc[i, capital_equity_columns[index]]
                        capital = 0 if pd.isna(capital) else capital
                        equity_change_dict[key][0] = capital
                        if len(project_equity_columns) == 1 and len(capital_equity_columns) == 2:
                            capital_0 = equity_change_df.iloc[i, capital_equity_columns[1]]
                            capital_0 = 0 if pd.isna(capital_0) else capital_0
                            equity_change_dict[key][14] = capital_0
                        preferred = equity_change_df.iloc[i, preferred_equity_columns[index]]
                        preferred = 0 if pd.isna(preferred) else preferred
                        equity_change_dict[key][1] = preferred
                        if len(project_equity_columns) == 1 and len(preferred_equity_columns) == 2:
                            preferred_0 = equity_change_df.iloc[i, preferred_equity_columns[1]]
                            preferred_0 = 0 if pd.isna(preferred_0) else preferred_0
                            equity_change_dict[key][15] = preferred_0
                        perpetual = equity_change_df.iloc[i, perpetual_equity_columns[index]]
                        perpetual = 0 if pd.isna(perpetual) else perpetual
                        equity_change_dict[key][2] = perpetual
                        if len(project_equity_columns) == 1 and len(perpetual_equity_columns) == 2:
                            perpetual_0 = equity_change_df.iloc[i, perpetual_equity_columns[1]]
                            perpetual_0 = 0 if pd.isna(perpetual_0) else perpetual_0
                            equity_change_dict[key][16] = perpetual_0
                        other = equity_change_df.iloc[i, other_equity_columns[index]]
                        other = 0 if pd.isna(other) else other
                        equity_change_dict[key][3] = other
                        if len(project_equity_columns) == 1 and len(other_equity_columns) == 2:
                            other_0 = equity_change_df.iloc[i, other_equity_columns[1]]
                            other_0 = 0 if pd.isna(other_0) else other_0
                            equity_change_dict[key][17] = other_0
                        reserve = equity_change_df.iloc[i, reserve_equity_columns[index]]
                        reserve = 0 if pd.isna(reserve) else reserve
                        equity_change_dict[key][4] = reserve
                        if len(project_equity_columns) == 1 and len(reserve_equity_columns) == 2:
                            reserve_0 = equity_change_df.iloc[i, reserve_equity_columns[1]]
                            reserve_0 = 0 if pd.isna(reserve_0) else reserve_0
                            equity_change_dict[key][18] = reserve_0
                        stock = equity_change_df.iloc[i, stock_equity_columns[index]]
                        stock = 0 if pd.isna(stock) else stock
                        equity_change_dict[key][5] = stock
                        if len(project_equity_columns) == 1 and len(stock_equity_columns) == 2:
                            stock_0 = equity_change_df.iloc[i, stock_equity_columns[1]]
                            stock_0 = 0 if pd.isna(stock_0) else stock_0
                            equity_change_dict[key][19] = stock_0
                        income = equity_change_df.iloc[i, income_equity_columns[index]]
                        income = 0 if pd.isna(income) else income
                        equity_change_dict[key][6] = income
                        if len(project_equity_columns) == 1 and len(income_equity_columns) == 2:
                            income_0 = equity_change_df.iloc[i, income_equity_columns[1]]
                            income_0 = 0 if pd.isna(income_0) else income_0
                            equity_change_dict[key][20] = income_0
                        special = equity_change_df.iloc[i, special_equity_columns[index]]
                        special = 0 if pd.isna(special) else special
                        equity_change_dict[key][7] = special
                        if len(project_equity_columns) == 1 and len(special_equity_columns) == 2:
                            special_0 = equity_change_df.iloc[i, special_equity_columns[1]]
                            special_0 = 0 if pd.isna(special_0) else special_0
                            equity_change_dict[key][21] = special_0
                        surplus = equity_change_df.iloc[i, surplus_equity_columns[index]]
                        surplus = 0 if pd.isna(surplus) else surplus
                        equity_change_dict[key][8] = surplus
                        if len(project_equity_columns) == 1 and len(surplus_equity_columns) == 2:
                            surplus_0 = equity_change_df.iloc[i, surplus_equity_columns[1]]
                            surplus_0 = 0 if pd.isna(surplus_0) else surplus_0
                            equity_change_dict[key][22] = surplus_0
                        risk = equity_change_df.iloc[i, risk_equity_columns[index]]
                        risk = 0 if pd.isna(risk) else risk
                        equity_change_dict[key][9] = risk
                        if len(project_equity_columns) == 1 and len(risk_equity_columns) == 2:
                            risk_0 = equity_change_df.iloc[i, risk_equity_columns[1]]
                            risk_0 = 0 if pd.isna(risk_0) else risk_0
                            equity_change_dict[key][23] = risk_0
                        profit = equity_change_df.iloc[i, profit_equity_columns[index]]
                        profit = 0 if pd.isna(profit) else profit
                        equity_change_dict[key][10] = profit
                        if len(project_equity_columns) == 1 and len(profit_equity_columns) == 2:
                            profit_0 = equity_change_df.iloc[i, profit_equity_columns[1]]
                            profit_0 = 0 if pd.isna(profit_0) else profit_0
                            equity_change_dict[key][24] = profit_0
                        subtotal = equity_change_df.iloc[i, subtotal_equity_columns[index]]
                        subtotal = 0 if pd.isna(subtotal) else subtotal
                        equity_change_dict[key][11] = subtotal
                        if len(project_equity_columns) == 1 and len(subtotal_equity_columns) == 2:
                            subtotal_0 = equity_change_df.iloc[i, subtotal_equity_columns[1]]
                            subtotal_0 = 0 if pd.isna(subtotal_0) else subtotal_0
                            equity_change_dict[key][25] = subtotal_0
                        minority = equity_change_df.iloc[i, minority_equity_columns[index]]
                        minority = 0 if pd.isna(minority) else minority
                        equity_change_dict[key][12] = minority
                        if len(project_equity_columns) == 1 and len(minority_equity_columns) == 2:
                            minority_0 = equity_change_df.iloc[i, minority_equity_columns[1]]
                            minority_0 = 0 if pd.isna(minority_0) else minority_0
                            equity_change_dict[key][26] = minority_0
                        total = equity_change_df.iloc[i, total_equity_columns[index]]
                        total = 0 if pd.isna(total) else total
                        equity_change_dict[key][13] = total
                        if len(project_equity_columns) == 1 and len(total_equity_columns) == 2:
                            total_0 = equity_change_df.iloc[i, total_equity_columns[1]]
                            total_0 = 0 if pd.isna(total_0) else total_0
                            equity_change_dict[key][27] = total_0
                        if capital != 0 or capital_0 != 0 or preferred != 0 or preferred_0 != 0 \
                                or perpetual != 0 or perpetual_0 != 0 or other != 0 or other_0 != 0 \
                                or reserve != 0 or reserve_0 != 0 or stock != 0 or stock_0 != 0\
                                or income != 0 or income_0 != 0 or special != 0 or special_0 != 0\
                                or surplus != 0 or surplus_0 != 0 or risk != 0 or risk_0 != 0\
                                or profit != 0 or profit_0 != 0 or subtotal != 0 or subtotal_0 != 0\
                                or minority != 0 or minority_0 != 0 or total != 0 or total_0 != 0:
                            if pd.notna(capital) or pd.notna(capital_0) or pd.notna(preferred) or pd.notna(preferred_0)\
                                or pd.notna(perpetual) or pd.notna(perpetual_0) or pd.notna(other) or pd.notna(other_0)\
                                or pd.notna(reserve) or pd.notna(reserve_0) or pd.notna(stock) or pd.notna(stock_0)\
                                or pd.notna(income) or pd.notna(income_0) or pd.notna(special) or pd.notna(special_0)\
                                or pd.notna(surplus) or pd.notna(surplus_0) or pd.notna(risk) or pd.notna(risk_0)\
                                or pd.notna(profit) or pd.notna(profit_0) or pd.notna(subtotal) or pd.notna(subtotal_0)\
                                or pd.notna(minority) or pd.notna(minority_0) or pd.notna(total) or pd.notna(total_0):
                                num += 1
                                display_equity_dict[num] = [key, capital, preferred, perpetual, other,\
                                                            reserve, stock, income, special,\
                                                            surplus, risk, profit, subtotal,\
                                                            minority, total,\
                                                            capital_0, preferred_0, perpetual_0, other_0,\
                                                            reserve_0, stock_0, income_0, special_0,\
                                                            surplus_0, risk_0, profit_0, subtotal_0,\
                                                            minority_0, total_0]
                                break
                elif re.search(equity_change_regex_dict[key], keywords):
                    capital = equity_change_df.iloc[i, capital_equity_columns[index]]
                    capital = 0 if pd.isna(capital) else capital
                    equity_change_dict[key][0] = capital
                    if len(project_equity_columns) == 1 and len(capital_equity_columns) == 2:
                        capital_0 = equity_change_df.iloc[i, capital_equity_columns[1]]
                        capital_0 = 0 if pd.isna(capital_0) else capital_0
                        equity_change_dict[key][14] = capital_0
                    preferred = equity_change_df.iloc[i, preferred_equity_columns[index]]
                    preferred = 0 if pd.isna(preferred) else preferred
                    equity_change_dict[key][1] = preferred
                    if len(project_equity_columns) == 1 and len(preferred_equity_columns) == 2:
                        preferred_0 = equity_change_df.iloc[i, preferred_equity_columns[1]]
                        preferred_0 = 0 if pd.isna(preferred_0) else preferred_0
                        equity_change_dict[key][15] = preferred_0
                    perpetual = equity_change_df.iloc[i, perpetual_equity_columns[index]]
                    perpetual = 0 if pd.isna(perpetual) else perpetual
                    equity_change_dict[key][2] = perpetual
                    if len(project_equity_columns) == 1 and len(perpetual_equity_columns) == 2:
                        perpetual_0 = equity_change_df.iloc[i, perpetual_equity_columns[1]]
                        perpetual_0 = 0 if pd.isna(perpetual_0) else perpetual_0
                        equity_change_dict[key][16] = perpetual_0
                    other = equity_change_df.iloc[i, other_equity_columns[index]]
                    other = 0 if pd.isna(other) else other
                    equity_change_dict[key][3] = other
                    if len(project_equity_columns) == 1 and len(other_equity_columns) == 2:
                        other_0 = equity_change_df.iloc[i, other_equity_columns[1]]
                        other_0 = 0 if pd.isna(other_0) else other_0
                        equity_change_dict[key][17] = other_0
                    reserve = equity_change_df.iloc[i, reserve_equity_columns[index]]
                    reserve = 0 if pd.isna(reserve) else reserve
                    equity_change_dict[key][4] = reserve
                    if len(project_equity_columns) == 1 and len(reserve_equity_columns) == 2:
                        reserve_0 = equity_change_df.iloc[i, reserve_equity_columns[1]]
                        reserve_0 = 0 if pd.isna(reserve_0) else reserve_0
                        equity_change_dict[key][18] = reserve_0
                    stock = equity_change_df.iloc[i, stock_equity_columns[index]]
                    stock = 0 if pd.isna(stock) else stock
                    equity_change_dict[key][5] = stock
                    if len(project_equity_columns) == 1 and len(stock_equity_columns) == 2:
                        stock_0 = equity_change_df.iloc[i, stock_equity_columns[1]]
                        stock_0 = 0 if pd.isna(stock_0) else stock_0
                        equity_change_dict[key][19] = stock_0
                    income = equity_change_df.iloc[i, income_equity_columns[index]]
                    income = 0 if pd.isna(income) else income
                    equity_change_dict[key][6] = income
                    if len(project_equity_columns) == 1 and len(income_equity_columns) == 2:
                        income_0 = equity_change_df.iloc[i, income_equity_columns[1]]
                        income_0 = 0 if pd.isna(income_0) else income_0
                        equity_change_dict[key][20] = income_0
                    special = equity_change_df.iloc[i, special_equity_columns[index]]
                    special = 0 if pd.isna(special) else special
                    equity_change_dict[key][7] = special
                    if len(project_equity_columns) == 1 and len(special_equity_columns) == 2:
                        special_0 = equity_change_df.iloc[i, special_equity_columns[1]]
                        special_0 = 0 if pd.isna(special_0) else special_0
                        equity_change_dict[key][21] = special_0
                    surplus = equity_change_df.iloc[i, surplus_equity_columns[index]]
                    surplus = 0 if pd.isna(surplus) else surplus
                    equity_change_dict[key][8] = surplus
                    if len(project_equity_columns) == 1 and len(surplus_equity_columns) == 2:
                        surplus_0 = equity_change_df.iloc[i, surplus_equity_columns[1]]
                        surplus_0 = 0 if pd.isna(surplus_0) else surplus_0
                        equity_change_dict[key][22] = surplus_0
                    risk = equity_change_df.iloc[i, risk_equity_columns[index]]
                    risk = 0 if pd.isna(risk) else risk
                    equity_change_dict[key][9] = risk
                    if len(project_equity_columns) == 1 and len(risk_equity_columns) == 2:
                        risk_0 = equity_change_df.iloc[i, risk_equity_columns[1]]
                        risk_0 = 0 if pd.isna(risk_0) else risk_0
                        equity_change_dict[key][23] = risk_0
                    profit = equity_change_df.iloc[i, profit_equity_columns[index]]
                    profit = 0 if pd.isna(profit) else profit
                    equity_change_dict[key][10] = profit
                    if len(project_equity_columns) == 1 and len(profit_equity_columns) == 2:
                        profit_0 = equity_change_df.iloc[i, profit_equity_columns[1]]
                        profit_0 = 0 if pd.isna(profit_0) else profit_0
                        equity_change_dict[key][24] = profit_0
                    subtotal = equity_change_df.iloc[i, subtotal_equity_columns[index]]
                    subtotal = 0 if pd.isna(subtotal) else subtotal
                    equity_change_dict[key][11] = subtotal
                    if len(project_equity_columns) == 1 and len(subtotal_equity_columns) == 2:
                        subtotal_0 = equity_change_df.iloc[i, subtotal_equity_columns[1]]
                        subtotal_0 = 0 if pd.isna(subtotal_0) else subtotal_0
                        equity_change_dict[key][25] = subtotal_0
                    minority = equity_change_df.iloc[i, minority_equity_columns[index]]
                    minority = 0 if pd.isna(minority) else minority
                    equity_change_dict[key][12] = minority
                    if len(project_equity_columns) == 1 and len(minority_equity_columns) == 2:
                        minority_0 = equity_change_df.iloc[i, minority_equity_columns[1]]
                        minority_0 = 0 if pd.isna(minority_0) else minority_0
                        equity_change_dict[key][26] = minority_0
                    total = equity_change_df.iloc[i, total_equity_columns[index]]
                    total = 0 if pd.isna(total) else total
                    equity_change_dict[key][13] = total
                    if len(project_equity_columns) == 1 and len(total_equity_columns) == 2:
                        total_0 = equity_change_df.iloc[i, total_equity_columns[1]]
                        total_0 = 0 if pd.isna(total_0) else total_0
                        equity_change_dict[key][27] = total_0
                    if capital != 0 or capital_0 != 0 or preferred != 0 or preferred_0 != 0 \
                            or perpetual != 0 or perpetual_0 != 0 or other != 0 or other_0 != 0 \
                            or reserve != 0 or reserve_0 != 0 or stock != 0 or stock_0 != 0\
                            or income != 0 or income_0 != 0 or special != 0 or special_0 != 0\
                            or surplus != 0 or surplus_0 != 0 or risk != 0 or risk_0 != 0\
                            or profit != 0 or profit_0 != 0 or subtotal != 0 or subtotal_0 != 0\
                            or minority != 0 or minority_0 != 0 or total != 0 or total_0 != 0:
                        if pd.notna(capital) or pd.notna(capital_0) or pd.notna(preferred) or pd.notna(preferred_0)\
                            or pd.notna(perpetual) or pd.notna(perpetual_0) or pd.notna(other) or pd.notna(other_0)\
                            or pd.notna(reserve) or pd.notna(reserve_0) or pd.notna(stock) or pd.notna(stock_0)\
                            or pd.notna(income) or pd.notna(income_0) or pd.notna(special) or pd.notna(special_0)\
                            or pd.notna(surplus) or pd.notna(surplus_0) or pd.notna(risk) or pd.notna(risk_0)\
                            or pd.notna(profit) or pd.notna(profit_0) or pd.notna(subtotal) or pd.notna(subtotal_0)\
                            or pd.notna(minority) or pd.notna(minority_0) or pd.notna(total) or pd.notna(total_0):
                            num += 1
                            display_equity_dict[num] = [key, capital, preferred, perpetual, other,\
                                                        reserve, stock, income, special,\
                                                        surplus, risk, profit, subtotal,\
                                                        minority, total,\
                                                        capital_0, preferred_0, perpetual_0, other_0,\
                                                        reserve_0, stock_0, income_0, special_0,\
                                                        surplus_0, risk_0, profit_0, subtotal_0,\
                                                        minority_0, total_0]
                            break
            else:
                key = '@未识别报表项目：' + keywords
                capital = equity_change_df.iloc[i, capital_equity_columns[index]]
                capital = 0 if pd.isna(capital) else capital
                if len(project_equity_columns) == 1 and len(capital_equity_columns) == 2:
                    capital_0 = equity_change_df.iloc[i, capital_equity_columns[1]]
                    capital_0 = 0 if pd.isna(capital_0) else capital_0
                preferred = equity_change_df.iloc[i, preferred_equity_columns[index]]
                preferred = 0 if pd.isna(preferred) else preferred
                if len(project_equity_columns) == 1 and len(preferred_equity_columns) == 2:
                    preferred_0 = equity_change_df.iloc[i, preferred_equity_columns[1]]
                    preferred_0 = 0 if pd.isna(preferred_0) else preferred_0
                perpetual = equity_change_df.iloc[i, perpetual_equity_columns[index]]
                perpetual = 0 if pd.isna(perpetual) else perpetual
                if len(project_equity_columns) == 1 and len(perpetual_equity_columns) == 2:
                    perpetual_0 = equity_change_df.iloc[i, perpetual_equity_columns[1]]
                    perpetual_0 = 0 if pd.isna(perpetual_0) else perpetual_0
                other = equity_change_df.iloc[i, other_equity_columns[index]]
                other = 0 if pd.isna(other) else other
                if len(project_equity_columns) == 1 and len(other_equity_columns) == 2:
                    other_0 = equity_change_df.iloc[i, other_equity_columns[1]]
                    other_0 = 0 if pd.isna(other_0) else other_0
                reserve = equity_change_df.iloc[i, reserve_equity_columns[index]]
                reserve = 0 if pd.isna(reserve) else reserve
                if len(project_equity_columns) == 1 and len(reserve_equity_columns) == 2:
                    reserve_0 = equity_change_df.iloc[i, reserve_equity_columns[1]]
                    reserve_0 = 0 if pd.isna(reserve_0) else reserve_0
                stock = equity_change_df.iloc[i, stock_equity_columns[index]]
                stock = 0 if pd.isna(stock) else stock
                if len(project_equity_columns) == 1 and len(stock_equity_columns) == 2:
                    stock_0 = equity_change_df.iloc[i, stock_equity_columns[1]]
                    stock_0 = 0 if pd.isna(stock_0) else stock_0
                income = equity_change_df.iloc[i, income_equity_columns[index]]
                income = 0 if pd.isna(income) else income
                if len(project_equity_columns) == 1 and len(income_equity_columns) == 2:
                    income_0 = equity_change_df.iloc[i, income_equity_columns[1]]
                    income_0 = 0 if pd.isna(income_0) else income_0
                special = equity_change_df.iloc[i, special_equity_columns[index]]
                special = 0 if pd.isna(special) else special
                if len(project_equity_columns) == 1 and len(special_equity_columns) == 2:
                    special_0 = equity_change_df.iloc[i, special_equity_columns[1]]
                    special_0 = 0 if pd.isna(special_0) else special_0
                surplus = equity_change_df.iloc[i, surplus_equity_columns[index]]
                surplus = 0 if pd.isna(surplus) else surplus
                if len(project_equity_columns) == 1 and len(surplus_equity_columns) == 2:
                    surplus_0 = equity_change_df.iloc[i, surplus_equity_columns[1]]
                    surplus_0 = 0 if pd.isna(surplus_0) else surplus_0
                risk = equity_change_df.iloc[i, risk_equity_columns[index]]
                risk = 0 if pd.isna(risk) else risk
                if len(project_equity_columns) == 1 and len(risk_equity_columns) == 2:
                    risk_0 = equity_change_df.iloc[i, risk_equity_columns[1]]
                    risk_0 = 0 if pd.isna(risk_0) else risk_0
                profit = equity_change_df.iloc[i, profit_equity_columns[index]]
                profit = 0 if pd.isna(profit) else profit
                if len(project_equity_columns) == 1 and len(profit_equity_columns) == 2:
                    profit_0 = equity_change_df.iloc[i, profit_equity_columns[1]]
                    profit_0 = 0 if pd.isna(profit_0) else profit_0
                subtotal = equity_change_df.iloc[i, subtotal_equity_columns[index]]
                subtotal = 0 if pd.isna(subtotal) else subtotal
                if len(project_equity_columns) == 1 and len(subtotal_equity_columns) == 2:
                    subtotal_0 = equity_change_df.iloc[i, subtotal_equity_columns[1]]
                    subtotal_0 = 0 if pd.isna(subtotal_0) else subtotal_0
                minority = equity_change_df.iloc[i, minority_equity_columns[index]]
                minority = 0 if pd.isna(minority) else minority
                if len(project_equity_columns) == 1 and len(minority_equity_columns) == 2:
                    minority_0 = equity_change_df.iloc[i, minority_equity_columns[1]]
                    minority_0 = 0 if pd.isna(minority_0) else minority_0
                total = equity_change_df.iloc[i, total_equity_columns[index]]
                total = 0 if pd.isna(total) else total
                if len(project_equity_columns) == 1 and len(total_equity_columns) == 2:
                    total_0 = equity_change_df.iloc[i, total_equity_columns[1]]
                    total_0 = 0 if pd.isna(total_0) else total_0
                if capital != 0 or capital_0 != 0 or preferred != 0 or preferred_0 != 0 \
                        or perpetual != 0 or perpetual_0 != 0 or other != 0 or other_0 != 0 \
                        or reserve != 0 or reserve_0 != 0 or stock != 0 or stock_0 != 0\
                        or income != 0 or income_0 != 0 or special != 0 or special_0 != 0\
                        or surplus != 0 or surplus_0 != 0 or risk != 0 or risk_0 != 0\
                        or profit != 0 or profit_0 != 0 or subtotal != 0 or subtotal_0 != 0\
                        or minority != 0 or minority_0 != 0 or total != 0 or total_0 != 0:
                    if pd.notna(capital) or pd.notna(capital_0) or pd.notna(preferred) or pd.notna(preferred_0)\
                        or pd.notna(perpetual) or pd.notna(perpetual_0) or pd.notna(other) or pd.notna(other_0)\
                        or pd.notna(reserve) or pd.notna(reserve_0) or pd.notna(stock) or pd.notna(stock_0)\
                        or pd.notna(income) or pd.notna(income_0) or pd.notna(special) or pd.notna(special_0)\
                        or pd.notna(surplus) or pd.notna(surplus_0) or pd.notna(risk) or pd.notna(risk_0)\
                        or pd.notna(profit) or pd.notna(profit_0) or pd.notna(subtotal) or pd.notna(subtotal_0)\
                        or pd.notna(minority) or pd.notna(minority_0) or pd.notna(total) or pd.notna(total_0):
                        num += 1
                        display_equity_dict[num] = [key, capital, preferred, perpetual, other,\
                                                    reserve, stock, income, special,\
                                                    surplus, risk, profit, subtotal,\
                                                    minority, total,\
                                                    capital_0, preferred_0, perpetual_0, other_0,\
                                                    reserve_0, stock_0, income_0, special_0,\
                                                    surplus_0, risk_0, profit_0, subtotal_0,\
                                                    minority_0, total_0]

    # 转换为DataFrame
    equity_change_data = pd.DataFrame(equity_change_dict)

    # 确定保存路径
    save_path = os.path.join(data_folder, 'equity_change_statement.csv')

    # 保存为CSV文件
    equity_change_data.to_csv(save_path, index=False, encoding='utf-8-sig')

    # 返回数据字典和保存路径
    return display_equity_dict, save_path