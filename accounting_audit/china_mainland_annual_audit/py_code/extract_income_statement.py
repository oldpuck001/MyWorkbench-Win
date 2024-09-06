# extract_income_statement.py

import os
import re
import pandas as pd

def extract_income_statement_xlsx(filepath, data_folder):
    income_statement_list = ['营业收入', '营业收入*', '利息收入*', '已赚保费*', '手续费及佣金收入*',\
                            '营业成本', '利息支出*', '手续费及佣金支出*', '退保金*', '赔付支出净额*', '提取保险责任准备金净额*',\
                            '保单红利支出*', '分保费用*',\
                            '税金及附加', '销售费用', '管理费用', '研发费用',\
                            '财务费用', '财务费用：利息费用', '财务费用：利息收入',\
                            '其他收益',\
                            '投资收益',\
                            '投资收益：对联营企业和合营企业的投资收益', '投资收益：以摊余成本计量的金融资产终止确认收益',\
                            '汇兑收益*',\
                            '净敞口套期收益', '公允价值变动收益', '信用减值损失', '资产减值损失', '资产处置收益',\
                            '营业利润',\
                            '营业外收入', '营业外支出',\
                            '利润总额',\
                            '所得税费用',\
                            '净利润',\
                            '持续经营净利润', '终止经营净利润',\
                            '*归属于母公司股东的净利润', '*少数股东损益',\
                            '其他综合收益的税后净额',\
                            '*归属于母公司所有者的其他综合收益的税后净额',\
                            '不能重分类进损益的其他综合收益',\
                            '重新计量设定受益计划变动额',\
                            '权益法下不能转损益的其他综合收益',\
                            '其他权益工具投资公允价值变动',\
                            '企业自身信用风险公允价值变动',\
                            '其他综合收益的税后净额：其它',\
                            '将重分类进损益的其他综合收益',\
                            '权益法下可转损益的其他综合收益',\
                            '其他债权投资公允价值变动',\
                            '金融资产重分类计入其他综合收益的金额',\
                            '其他债权投资信用减值准备',\
                            '现金流量套期储备',\
                            '外币财务报表折算差额',\
                            '将重分类进损益的其他综合收益：其它',\
                            '*归属于少数股东的其他综合收益的税后净额',\
                            '综合收益总额',\
                            '*归属于母公司所有者的综合收益总额', '*归属于少数股东的综合收益总额',\
                            '基本每股收益',\
                            '稀释每股收益']

    income_statement_dict = {key: [0, 0, 0, 0, 0, 0] for key in income_statement_list}

    income_statement_regex_dict = {'营业收入': re.compile(r'\S*一\S{1}营业收入\S*|\S*一\S{1}营业总收入\S*'),\
                                   '营业收入*': re.compile(r'\S*其中\S{1}营业收入\S*'),\
                                   '已赚保费*': re.compile(r'\S*已赚保费\S*'),\
                                   '手续费及佣金收入*': re.compile(r'\S*手续费及佣金收入\S*'),\
                                   '营业成本': re.compile(r'\S*营业成本\S*'),\
                                   '利息支出*': re.compile(r'^利息支出\S*'),\
                                   '手续费及佣金支出*': re.compile(r'\S*手续费及佣金支出\S*'),\
                                   '退保金*': re.compile(r'\S*退保金\S*'),\
                                   '赔付支出净额*': re.compile(r'\S*赔付支出净额\S*'),\
                                   '提取保险责任准备金净额*': re.compile(r'\S*提取保险责任准备金净额\S*'),\
                                   '保单红利支出*': re.compile(r'\S*保单红利支出\S*'),\
                                   '分保费用*': re.compile(r'\S*分保费用\S*'),\
                                   '税金及附加': re.compile(r'\S*税金及附加\S*'),\
                                   '销售费用': re.compile(r'\S*销售费用\S*|\S*营业费用\S*'),\
                                   '管理费用': re.compile(r'\S*管理费用\S*'),\
                                   '研发费用': re.compile(r'\S*研发费用\S*'),\
                                   '财务费用': re.compile(r'\S*财务费用\S*'),\
                                   '财务费用：利息费用': re.compile(r'\S*其中\S{1}利息费用\S*|\S*其中\S{1}利息支出\S*'),\
                                   '其他收益': re.compile(r'\S*加\S{1}其他收益\S*|\S*加\S{1}其它收益\S*'),\
                                   '投资收益': re.compile(r'\S*投资收益\S*'),\
                                   '投资收益：对联营企业和合营企业的投资收益': re.compile(r'\S*其中\S{1}对联营企业和合营企业的投资收益\S*'),\
                                   '投资收益：以摊余成本计量的金融资产终止确认收益': re.compile(r'\S*以摊余成本计量的金融资产终止确认收益\S*'),\
                                   '汇兑收益*': re.compile(r'\S*汇兑收益\S*'),\
                                   '净敞口套期收益': re.compile(r'\S*净敞口套期收益\S*'),\
                                   '公允价值变动收益': re.compile(r'\S*公允价值变动收益\S*'),\
                                   '信用减值损失': re.compile(r'\S*信用减值损失\S*'),\
                                   '资产减值损失': re.compile(r'\S*资产减值损失\S*'),\
                                   '资产处置收益': re.compile(r'\S*资产处置收益\S*'),\
                                   '营业利润': re.compile(r'\S*营业利润\S*'),\
                                   '营业外收入': re.compile(r'\S*营业外收入\S*'),\
                                   '营业外支出': re.compile(r'\S*营业外支出\S*'),\
                                   '利润总额': re.compile(r'\S*利润总额\S*'),\
                                   '所得税费用': re.compile(r'\S*所得税费用\S*'),\
                                   '净利润': re.compile(r'\S*净利润\S*'),\
                                   '持续经营净利润': re.compile(r'\S*持续经营净利润\S*'),\
                                   '终止经营净利润': re.compile(r'\S*终止经营净利润\S*'),\
                                   '*归属于母公司股东的净利润': re.compile(r'\S*归属于母公司股东的净利润\S*'),\
                                   '*少数股东损益': re.compile(r'\S*少数股东损益\S*'),\
                                   '其他综合收益的税后净额': re.compile(r'\S*其他综合收益的税后净额\S*'),\
                                   '*归属于母公司所有者的其他综合收益的税后净额': re.compile(r'\S*归属于母公司所有者的其他综合收益的税后净额\S*'),\
                                   '不能重分类进损益的其他综合收益': re.compile(r'\S*不能重分类进损益的其他综合收益\S*'),\
                                   '重新计量设定受益计划变动额': re.compile(r'\S*重新计量设定受益计划变动额\S*'),\
                                   '权益法下不能转损益的其他综合收益': re.compile(r'\S*权益法下不能转损益的其他综合收益\S*'),\
                                   '其他权益工具投资公允价值变动': re.compile(r'\S*其他权益工具投资公允价值变动\S*'),\
                                   '企业自身信用风险公允价值变动': re.compile(r'\S*企业自身信用风险公允价值变动\S*'),\
                                   '将重分类进损益的其他综合收益': re.compile(r'\S*将重分类进损益的其他综合收益\S*'),\
                                   '权益法下可转损益的其他综合收益': re.compile(r'\S*权益法下可转损益的其他综合收益\S*'),\
                                   '其他债权投资公允价值变动': re.compile(r'\S*其他债权投资公允价值变动\S*'),\
                                   '金融资产重分类计入其他综合收益的金额': re.compile(r'\S*金融资产重分类计入其他综合收益的金额\S*'),\
                                   '其他债权投资信用减值准备': re.compile(r'\S*其他债权投资信用减值准备\S*'),\
                                   '现金流量套期储备': re.compile(r'\S*现金流量套期储备\S*'),\
                                   '外币财务报表折算差额': re.compile(r'\S*外币财务报表折算差额\S*'),\
                                   '*归属于少数股东的其他综合收益的税后净额': re.compile(r'\S*归属于少数股东的其他综合收益的税后净额\S*'),\
                                   '综合收益总额': re.compile(r'\S*综合收益总额\S*'),\
                                   '*归属于母公司所有者的综合收益总额': re.compile(r'\S*归属于母公司所有者的综合收益总额\S*'),\
                                   '*归属于少数股东的综合收益总额': re.compile(r'\S*归属于少数股东的综合收益总额\S*'),\
                                   '基本每股收益': re.compile(r'\S*基本每股收益\S*'),\
                                   '稀释每股收益': re.compile(r'\S*稀释每股收益\S*')}

    display_income_dict = {}

    num = 0
    
    file_extension = os.path.splitext(filepath)[1].lower()
    if file_extension == '.xlsx':
        income_statement_df = pd.read_excel(filepath, engine='openpyxl')
    elif file_extension == '.xls':
        income_statement_df = pd.read_excel(filepath, engine='xlrd')

    # 检查文件有多少列及表头所在的行次
    num_columns = income_statement_df.shape[1]

    # 使用正则表达式来匹配前后和中间可能存在空格的“项目”
    pattern_project_income = re.compile(r'\s*项\s*目\s*')
    # 查找含有“项目”变体的行
    start_row_index = income_statement_df[income_statement_df.astype(str).apply(lambda row: row.str.contains(pattern_project_income).any(), axis=1)].index[0]

    # 正则表达式匹配“本期金额，本期数”关键字
    pattern_this_income = re.compile(r'\s*本\s*期\s*金\s*额\s*|\s*本\s*期\s*数\s*')

    # 正则表达式匹配“上期余额，上期数”关键字
    pattern_previous_income = re.compile(
        r'\s*上\s*期\s*金\s*额\s*|'
        r'\s*上\s*期\s*数\s*'
    )

    # 列表来保存找到的列索引
    project_income_columns = []
    this_income_columns = []
    previous_income_columns = []

    # 遍历列，寻找匹配的列索引
    for n in range(num_columns):
        header = str(income_statement_df.iloc[start_row_index, n])
        if re.search(pattern_project_income, header):
            project_income_columns.append(n)
        elif re.search(pattern_this_income, header):
            this_income_columns.append(n)
        elif re.search(pattern_previous_income, header):
            previous_income_columns.append(n)

    for n in project_income_columns:
        index = project_income_columns.index(n)
        for i in range(start_row_index + 1, len(income_statement_df)):
            keywords = income_statement_df.iloc[i, n]
            if isinstance(keywords, str):
                keywords = keywords.replace(' ', '')
            if isinstance(keywords, float):
                continue
            for key in income_statement_list:
                if key in ['其他综合收益的税后净额：其它', '将重分类进损益的其他综合收益：其它']:
                    continue
                elif key in ['财务费用：利息收入', '利息收入*']:
                    if re.search(re.compile(r'\S*利息收入\S*'), keywords):
                        if income_statement_df.iloc[i-2, n].replace(' ', '') == '财务费用':
                            key = '财务费用：利息收入'
                        elif income_statement_df.iloc[i-2, n].replace(' ', '') != '财务费用':
                            key = '利息收入*'
                        previous = income_statement_df.iloc[i, previous_income_columns[index]]
                        previous = 0 if pd.isna(previous) else previous
                        this = income_statement_df.iloc[i, this_income_columns[index]]
                        this = 0 if pd.isna(this) else this
                        income_statement_dict[key][0] = previous
                        income_statement_dict[key][1] = this
                        if previous != 0 or this != 0:
                            if pd.notna(previous) or pd.notna(this):
                                num += 1
                                display_income_dict[num] = [key, this, previous]
                                break
                elif re.search(income_statement_regex_dict[key], keywords):
                    previous = income_statement_df.iloc[i, previous_income_columns[index]]
                    previous = 0 if pd.isna(previous) else previous
                    this = income_statement_df.iloc[i, this_income_columns[index]]
                    this = 0 if pd.isna(this) else this
                    income_statement_dict[key][0] = previous
                    income_statement_dict[key][1] = this
                    if previous != 0 or this != 0:
                        if pd.notna(previous) or pd.notna(this):
                            num += 1
                            display_income_dict[num] = [key, this, previous]
                            break
            else:
                key = '@未识别报表项目：' + keywords
                previous = income_statement_df.iloc[i, previous_income_columns[index]]
                previous = 0 if pd.isna(previous) else previous
                this = income_statement_df.iloc[i, this_income_columns[index]]
                this = 0 if pd.isna(this) else this
                if previous != 0 or this != 0:
                    if pd.notna(previous) or pd.notna(this):
                        num += 1
                        display_income_dict[num] = [key, this, previous]

    # 转换为DataFrame
    income_statement_data = pd.DataFrame(income_statement_dict)

    row1 = income_statement_data.iloc[1]
    row2 = income_statement_data.iloc[2]
    row3 = income_statement_data.iloc[3]
    new_row4 = row1 + row2 + row3
    income_statement_data.iloc[4] = new_row4
    income_statement_data = income_statement_data.round(2)

    # 确定保存路径
    save_path = os.path.join(data_folder, 'income_statement.csv')
        
    # 保存为CSV文件
    income_statement_data.to_csv(save_path, index=False, encoding='utf-8-sig')

    # 返回数据字典和保存路径
    return display_income_dict, save_path