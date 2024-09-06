# extract_balance_sheet.py

import os
import re
import pandas as pd

def extract_balance_sheet_xlsx(filepath, data_folder):
    balance_sheet_list = ['货币资金', '结算备付金*', '拆出资金*', '交易性金融资产', '△以公允价值计量且其变动计入当期损益的金融资产',\
                            '衍生金融资产', '应收票据', '应收账款', '应收款项融资','预付款项', '应收保费*', '应收分保账款*',\
                            '应收分保合同准备金*', '其他应收款', '买入返售金融资产*', '存货', '合同资产', '持有待售资产', \
                            '一年内到期的非流动资产', '其他流动资产',\
                            '流动资产合计',\
                            '发放贷款和垫款*', '债权投资', '△可供出售金融资产', '其他债权投资', '△持有至到期投资', '长期应收款',\
                            '长期股权投资', '其他权益工具投资', '其他非流动金融资产', '投资性房地产', '固定资产', '在建工程',\
                            '生产性生物资产', '油气资产', '使用权资产', '无形资产', '开发支出', '商誉', '长期待摊费用',\
                            '递延所得税资产', '其他非流动资产',\
                            '非流动资产合计',\
                            '资产总计',\
                            '短期借款', '向中央银行借款*', '拆入资金*', '交易性金融负债', '△以公允价值计量且其变动计入当期损益的金融负债',\
                            '衍生金融负债', '应付票据', '应付账款', '预收款项', '合同负债', '卖出回购金融资产款*', '吸收存款及同业存放*',\
                            '代理买卖证券款*', '代理承销证券款*', '应付职工薪酬', '应交税费', '其他应付款', '应付手续费及佣金*',\
                            '应付分保账款*', '持有待售负债', '一年内到期的非流动负债', '其他流动负债',\
                            '流动负债合计',\
                            '保险合同准备金*', '长期借款', '应付债券','应付债券：优先股', '应付债券：永续债', '租赁负债', '长期应付款',\
                            '预计负债', '递延收益', '递延所得税负债', '其他非流动负债',\
                            '非流动负债合计',\
                            '负债合计',\
                            '实收资本（或股本）', '#减：已归还投资', '#实收资本（或股本）净额', '其他权益工具', '其他权益工具：优先股',\
                            '其他权益工具：永续债', '资本公积', '减：库存股', '其他综合收益', '专项储备', '盈余公积', '一般风险准备*',\
                            '未分配利润', '*归属于母公司所有者权益（或股东权益）合计', '*少数股东权益',\
                            '所有者权益（或股东权益）合计',\
                            '负债和所有者权益（或股东权益）总计']

    balance_sheet_dict = {key: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0] for key in balance_sheet_list}

    balance_sheet_regex_dict = {'货币资金': re.compile(r'\S*货币资金\S*'),\
                                '结算备付金*': re.compile(r'\S*结算备付金\S*'),\
                                '拆出资金*': re.compile(r'\S*拆出资金\S*'),\
                                '交易性金融资产': re.compile(r'\S*交易性金融资产\S*'),\
                                '△以公允价值计量且其变动计入当期损益的金融资产': re.compile(r'\S*以公允价值计量且其变动计入当期损益的金融资产\S*'),\
                                '衍生金融资产': re.compile(r'\S*衍生金融资产\S*'),\
                                '应收票据': re.compile(r'\S*应收票据\S*'),\
                                '应收账款': re.compile(r'\S*应收账款\S*|\S*应收帐款\S*'),\
                                '应收款项融资': re.compile(r'\S*应收款项融资\S*'),\
                                '预付款项': re.compile(r'\S*预付款项\S*|\S*预付账款\S*|\S*预付帐款\S*'),\
                                '应收保费*': re.compile(r'\S*应收保费\S*'),\
                                '应收分保账款*': re.compile(r'\S*应收分保账款\S*'),\
                                '应收分保合同准备金*': re.compile(r'\S*应收分保合同准备金\S*'),\
                                '其他应收款': re.compile(r'\S*其他应收款\S*'),\
                                '买入返售金融资产*': re.compile(r'\S*买入返售金融资产\S*'),\
                                '存货': re.compile(r'\S*存货\S*'),\
                                '合同资产': re.compile(r'\S*合同资产\S*'),\
                                '持有待售资产': re.compile(r'\S*持有待售资产\S*'),\
                                '一年内到期的非流动资产': re.compile(r'\S*一年内到期的非流动资产\S*'),\
                                '其他流动资产': re.compile(r'\S*其他流动资产\S*'),\
                                '流动资产合计': re.compile(r'^流动资产合计\S*'),\
                                '发放贷款和垫款*': re.compile(r'\S*发放贷款和垫款\S*'),\
                                '债权投资': re.compile(r'\S*债权投资\S*'),\
                                '△可供出售金融资产': re.compile(r'\S*可供出售金融资产\S*'),\
                                '其他债权投资': re.compile(r'\S*其他债权投资\S*'),\
                                '△持有至到期投资': re.compile(r'\S*持有至到期投资\S*'),\
                                '长期应收款': re.compile(r'\S*长期应收款\S*'),\
                                '长期股权投资': re.compile(r'\S*长期股权投资\S*'),\
                                '其他权益工具投资': re.compile(r'\S*其他权益工具投资\S*'),\
                                '其他非流动金融资产': re.compile(r'\S*其他非流动金融资产\S*'),\
                                '投资性房地产': re.compile(r'\S*投资性房地产\S*'),\
                                '固定资产': re.compile(r'\S*固定资产$|\S*固定资产净额$'),\
                                '在建工程': re.compile(r'\S*在建工程\S*'),\
                                '生产性生物资产': re.compile(r'\S*生产性生物资产\S*'),\
                                '油气资产': re.compile(r'\S*油气资产\S*'),\
                                '使用权资产': re.compile(r'\S*使用权资产\S*'),\
                                '无形资产': re.compile(r'\S*无形资产\S*'),\
                                '开发支出': re.compile(r'\S*开发支出\S*|\S*研发支出\S*'),\
                                '商誉': re.compile(r'\S*商誉\S*'),\
                                '长期待摊费用': re.compile(r'\S*长期待摊费用\S*'),\
                                '递延所得税资产': re.compile(r'\S*递延所得税资产\S*|\S*递延税款借项\S*'),\
                                '其他非流动资产': re.compile(r'\S*其他非流动资产\S*|\S*其它非流动资产\S*'),\
                                '非流动资产合计': re.compile(r'^非流动资产合计\S*'),\
                                '资产总计': re.compile(r'\S*资产总计\S*'),\
                                '短期借款': re.compile(r'\S*短期借款\S*'),\
                                '向中央银行借款*': re.compile(r'\S*向中央银行借款\S*'),\
                                '拆入资金*': re.compile(r'\S*拆入资金\S*'),\
                                '交易性金融负债': re.compile(r'\S*交易性金融负债\S*'),\
                                '△以公允价值计量且其变动计入当期损益的金融负债': re.compile(r'\S*以公允价值计量且其变动计入当期损益的金融负债\S*'),\
                                '衍生金融负债': re.compile(r'\S*衍生金融负债\S*'),\
                                '应付票据': re.compile(r'\S*应付票据\S*'),\
                                '应付账款': re.compile(r'\S*应付账款\S*|\S*应付帐款\S*'),\
                                '预收款项': re.compile(r'\S*预收款项\S*|\S*预收账款\S*|\S*预收帐款\S*'),\
                                '合同负债': re.compile(r'\S*合同负债\S*'),\
                                '卖出回购金融资产款*': re.compile(r'\S*卖出回购金融资产款\S*'),\
                                '吸收存款及同业存放*': re.compile(r'\S*吸收存款及同业存放\S*'),\
                                '代理买卖证券款*': re.compile(r'\S*代理买卖证券款\S*'),\
                                '代理承销证券款*': re.compile(r'\S*代理承销证券款\S*'),\
                                '应付职工薪酬': re.compile(r'\S*应付职工薪酬\S*'),\
                                '应交税费': re.compile(r'\S*应交税费\S*'),\
                                '其他应付款': re.compile(r'\S*其他应付款\S*'),\
                                '应付手续费及佣金*': re.compile(r'\S*应付手续费及佣金\S*'),\
                                '应付分保账款*': re.compile(r'\S*应付分保账款\S*'),\
                                '持有待售负债': re.compile(r'\S*持有待售负债\S*'),\
                                '一年内到期的非流动负债': re.compile(r'\S*一年内到期的非流动负债\S*'),\
                                '其他流动负债': re.compile(r'\S*其他流动负债\S*'),\
                                '流动负债合计': re.compile(r'^流动负债合计\S*'),\
                                '保险合同准备金*': re.compile(r'\S*保险合同准备金\S*'),\
                                '长期借款': re.compile(r'\S*长期借款\S*'),\
                                '应付债券': re.compile(r'\S*应付债券\S*'),\
                                '租赁负债': re.compile(r'\S*租赁负债\S*'),\
                                '长期应付款': re.compile(r'\S*长期应付款\S*'),\
                                '预计负债': re.compile(r'\S*预计负债\S*'),\
                                '递延收益': re.compile(r'\S*递延收益\S*'),\
                                '递延所得税负债': re.compile(r'\S*递延所得税负债\S*|\S*递延税款贷项\S*'),\
                                '其他非流动负债': re.compile(r'\S*其他非流动负债\S*'),\
                                '非流动负债合计': re.compile(r'^非流动负债合计\S*'),\
                                '负债合计': re.compile(r'\S*负债合计\S*'),\
                                '实收资本（或股本）': re.compile(r'\S*实收资本\S{1,2}股本\S{1}$|\S*实收资本$|\S*股本$'),\
                                '#减：已归还投资': re.compile(r'\S*已归还投资\S*'),\
                                '#实收资本（或股本）净额': re.compile(r'\S*实收资本\S{1,2}股本\S{1}净额$|\S*实收资本净额$|\S*股本净额$'),\
                                '其他权益工具': re.compile(r'\S*其他权益工具\S*'),\
                                '资本公积': re.compile(r'\S*资本公积\S*'),\
                                '减：库存股': re.compile(r'\S*减\S{1}库存股\S*'),\
                                '其他综合收益': re.compile(r'\S*其他综合收益\S*'),\
                                '专项储备': re.compile(r'\S*专项储备\S*'),\
                                '盈余公积': re.compile(r'\S*盈余公积\S*'),\
                                '一般风险准备*': re.compile(r'\S*一般风险准备\S*'),\
                                '未分配利润': re.compile(r'\S*未分配利润\S*'),\
                                '*归属于母公司所有者权益（或股东权益）合计': re.compile(r'\S*归属于母公司所有者权益\S{1,2}股东权益\S{1}合计\S*|\S*归属于母公司所有者权益合计\S*|\S*归属于母公司股东权益合计\S*'),\
                                '*少数股东权益': re.compile(r'\S*少数股东权益\S*'),\
                                '所有者权益（或股东权益）合计': re.compile(r'\S*所有者权益\S{1,2}股东权益\S{1}合计\S*|\S*所有者权益合计\S*|\S*股东权益合计\S*'),\
                                '负债和所有者权益（或股东权益）总计': re.compile(r'\S*负债和所有者权益\S{1,2}股东权益\S{1}总计\S*|\S*负债和所有者权益总计\S*|\S*负债和股东权益总计\S*')}

    display_balance_dict = {}

    num = 0
    
    file_extension = os.path.splitext(filepath)[1].lower()
    if file_extension == '.xlsx':
        balance_sheet_df = pd.read_excel(filepath, engine='openpyxl')
    elif file_extension == '.xls':
        balance_sheet_df = pd.read_excel(filepath, engine='xlrd')

    # 检查文件有多少列及表头所在的行次
    num_columns = balance_sheet_df.shape[1]

    # 使用正则表达式来匹配前后和中间可能存在空格的“项目”变体
    pattern_project_balance = re.compile(
        r'\s*项\s*目\s*|'                                                               # 匹配“项目”及其变体
        r'\s*资\s*产\s*|'                                                               # 匹配“资产”及其变体
        r'\s*负\s*债\s*和\s*所\s*有\s*者\s*权\s*益\s*|'                                   # 匹配“负债和所有者权益”及其变体
        r'\s*负\s*债\s*和\s*所\s*有\s*者\s*权\s*益\s*（\s*或\s*股\s*东\s*权\s*益\s*）\s*|'   # 匹配“负债和所有者权益（或股东权益）”及其变体
        r'\s*负\s*债\s*和\s*所\s*有\s*者\s*权\s*益\s*\(\s*或\s*股\s*东\s*权\s*益\s*\)\s*'   # 匹配带有半角括号的情况
    )
    # 查找含有“项目”变体的行
    start_row_index = balance_sheet_df[balance_sheet_df.astype(str).apply(lambda row: row.str.contains(pattern_project_balance).any(), axis=1)].index[0]

    # 正则表达式匹配“期末余额，期末数”关键字
    pattern_end_balance = re.compile(r'\s*期\s*末\s*余\s*额\s*|\s*期\s*末\s*数\s*')

    # 正则表达式匹配“上年年末余额，上年年末数，年初余额，年初数”关键字
    pattern_previous_balance = re.compile(
        r'\s*上\s*年\s*年\s*末\s*余\s*额\s*|'   # 匹配“上年年末余额”
        r'\s*上\s*年\s*年\s*末\s*数\s*|'        # 匹配“上年年末数”
        r'\s*年\s*初\s*余\s*额\s*|'            # 匹配“年初余额”
        r'\s*年\s*初\s*数\s*'                  # 匹配“年初数”
    )

    # 列表来保存找到的列索引
    project_balance_columns = []
    end_balance_columns = []
    previous_balance_columns = []

    # 遍历列，寻找匹配的列索引
    for n in range(num_columns):
        header = str(balance_sheet_df.iloc[start_row_index, n])
        if re.search(pattern_project_balance, header):
            project_balance_columns.append(n)
        elif re.search(pattern_end_balance, header):
            end_balance_columns.append(n)
        elif re.search(pattern_previous_balance, header):
            previous_balance_columns.append(n)

    for n in project_balance_columns:
        index = project_balance_columns.index(n)
        for i in range(start_row_index + 1, len(balance_sheet_df)):
            keywords = balance_sheet_df.iloc[i, n]
            if isinstance(keywords, str):
                keywords = keywords.replace(' ', '')
            if isinstance(keywords, float):
                continue
            for key in balance_sheet_list:
                if key in ['应付债券：优先股', '其他权益工具：优先股']:
                    if re.search(re.compile(r'\S*其中\S{1}优先股\S*'), keywords):
                        if balance_sheet_df.iloc[i-1, n].replace(' ', '') == '应付债券':
                            key = '应付债券：优先股'
                        elif balance_sheet_df.iloc[i-1, n].replace(' ', '') != '其他权益工具':
                            key = '其他权益工具：优先股'
                        previous = balance_sheet_df.iloc[i, previous_balance_columns[index]]
                        previous = 0 if pd.isna(previous) else previous
                        end = balance_sheet_df.iloc[i, end_balance_columns[index]]
                        end = 0 if pd.isna(end) else end
                        balance_sheet_dict[key][0] = previous
                        balance_sheet_dict[key][5] = end
                        if previous != 0 or end != 0:
                            if pd.notna(previous) or pd.notna(end):
                                num += 1
                                display_balance_dict[num] = [key, end, previous]
                                break
                elif key in ['应付债券：永续债', '其他权益工具：永续债']:
                    if re.search(re.compile(r'\S*永续债\S*'), keywords):
                        if balance_sheet_df.iloc[i-2, n].replace(' ', '') == '应付债券':
                            key = '应付债券：永续债'
                        elif balance_sheet_df.iloc[i-2, n].replace(' ', '') != '其他权益工具':
                            key = '其他权益工具：永续债'
                        previous = balance_sheet_df.iloc[i, previous_balance_columns[index]]
                        previous = 0 if pd.isna(previous) else previous
                        end = balance_sheet_df.iloc[i, end_balance_columns[index]]
                        end = 0 if pd.isna(end) else end
                        balance_sheet_dict[key][0] = previous
                        balance_sheet_dict[key][5] = end
                        if previous != 0 or end != 0:
                            if pd.notna(previous) or pd.notna(end):
                                num += 1
                                display_balance_dict[num] = [key, end, previous]
                                break
                elif re.search(balance_sheet_regex_dict[key], keywords):
                    previous = balance_sheet_df.iloc[i, previous_balance_columns[index]]
                    previous = 0 if pd.isna(previous) else previous
                    end = balance_sheet_df.iloc[i, end_balance_columns[index]]
                    end = 0 if pd.isna(end) else end
                    balance_sheet_dict[key][0] = previous
                    balance_sheet_dict[key][5] = end
                    if previous != 0 or end != 0:
                        if pd.notna(previous) or pd.notna(end):
                            num += 1
                            display_balance_dict[num] = [key, end, previous]
                            break
            else:
                if keywords in ['流动资产：', '非流动资产：', '流动负债：', '非流动负债：', '所有者权益（或股东权益）：']:
                    key = keywords
                else:
                    key = '@未识别报表项目：' + keywords
                previous = balance_sheet_df.iloc[i, previous_balance_columns[index]]
                previous = 0 if pd.isna(previous) else previous
                end = balance_sheet_df.iloc[i, end_balance_columns[index]]
                end = 0 if pd.isna(end) else end
                if previous != 0 or end != 0:
                    if pd.notna(previous) or pd.notna(end):
                        num += 1
                        display_balance_dict[num] = [key, end, previous]

    # 转换为DataFrame
    balance_sheet_data = pd.DataFrame(balance_sheet_dict)

    row0 = balance_sheet_data.iloc[0]
    row1 = balance_sheet_data.iloc[1]
    row2 = balance_sheet_data.iloc[2]
    new_row3 = row0 + row1 + row2
    balance_sheet_data.iloc[3] = new_row3
    row5 = balance_sheet_data.iloc[5]
    row6 = balance_sheet_data.iloc[6]
    row7 = balance_sheet_data.iloc[7]
    new_row8 = row5 + row6 + row7
    balance_sheet_data.iloc[8] = new_row8
    balance_sheet_data = balance_sheet_data.round(2)

    # 确定保存路径
    save_path = os.path.join(data_folder, 'balance_sheet.csv')

    # 保存为CSV文件
    balance_sheet_data.to_csv(save_path, index=False, encoding='utf-8-sig')

    # 返回数据字典和保存路径
    return display_balance_dict, save_path