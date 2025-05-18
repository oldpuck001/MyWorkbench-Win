# import_income_statement.py

import os
import re
import pandas as pd

def select_income_statement(request):

    file_path = request.get("data", {}).get("file_path", "")

    sheet_file = pd.ExcelFile(file_path)                                   # 使用pandas讀取Excel文件
    sheetnames = sheet_file.sheet_names                                    # 獲取所有工作表名稱

    return ['select_income_statement', sheetnames]


def import_income_statement(request):

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
                             '将重分类进损益的其他综合收益',\
                             '权益法下可转损益的其他综合收益',\
                             '其他债权投资公允价值变动',\
                             '金融资产重分类计入其他综合收益的金额',\
                             '其他债权投资信用减值准备',\
                             '现金流量套期储备',\
                             '外币财务报表折算差额',\
                             '*归属于少数股东的其他综合收益的税后净额',\
                             '综合收益总额',\
                             '*归属于母公司所有者的综合收益总额', '*归属于少数股东的综合收益总额',\
                             '基本每股收益',\
                             '稀释每股收益']

    income_statement_dict = {key: [0, 0] for key in income_statement_list}

    income_statement_regex_dict = {'营业收入': re.compile(r'(?:^|\s+)(?:一\S{1}营业收入|一\S{1}营业总收入)(?:\s+|$)'),\
                                   '营业收入*': re.compile(r'(?:^|\s+)(\S?其中\S{1}营业收入\S?)(?:\s+|$)'),\
                                   '利息收入*': re.compile(r'(?:^|\s+)(\S?利息收入\S?)(?:\s+|$)'),\
                                   '已赚保费*': re.compile(r'(?:^|\s+)(\S?已赚保费\S?)(?:\s+|$)'),\
                                   '手续费及佣金收入*': re.compile(r'(?:^|\s+)(\S?手续费及佣金收入\S?)(?:\s+|$)'),\
                                   '营业成本': re.compile(r'(?:^|\s+)(?:营业成本|其中\S{1}营业成本)(?:\s+|$)'),\
                                   '利息支出*': re.compile(r'(?:^|\s+)(\S?利息支出\S?)(?:\s+|$)'),\
                                   '手续费及佣金支出*': re.compile(r'(?:^|\s+)(\S?手续费及佣金支出\S?)(?:\s+|$)'),\
                                   '退保金*': re.compile(r'(?:^|\s+)(\S?退保金\S?)(?:\s+|$)'),\
                                   '赔付支出净额*': re.compile(r'(?:^|\s+)(\S?赔付支出净额\S?)(?:\s+|$)'),\
                                   '提取保险责任准备金净额*': re.compile(r'(?:^|\s+)(\S?提取保险责任准备金净额\S?)(?:\s+|$)'),\
                                   '保单红利支出*': re.compile(r'(?:^|\s+)(\S?保单红利支出\S?)(?:\s+|$)'),\
                                   '分保费用*': re.compile(r'(?:^|\s+)(\S?分保费用\S?)(?:\s+|$)'),\
                                   '税金及附加': re.compile(r'(?:^|\s+)(?:税金及附加|营业税金及附加|主营业务税金及附加)(?:\s+|$)'),\
                                   '销售费用': re.compile(r'(?:^|\s+)(?:销售费用|营业费用)(?:\s+|$)'),\
                                   '管理费用': re.compile(r'(?:^|\s+)管理费用(?:\s+|$)'),\
                                   '研发费用': re.compile(r'(?:^|\s+)研发费用(?:\s+|$)'),\
                                   '财务费用': re.compile(r'(?:^|\s+)财务费用(?:\s+|$)'),\
                                   '财务费用：利息费用': re.compile(r'(?:^|\s+)(?:利息费用|其中\S{1}利息费用)(?:\s+|$)'),\
                                   '财务费用：利息收入': re.compile(r'(?:^|\s+)(?:利息收入|其中\S{1}利息收入)(?:\s+|$)'),\
                                   '其他收益': re.compile(r'(?:^|\s+)(?:其他收益|其中\S{1}其他收益)(?:\s+|$)'),\
                                   '投资收益':re.compile(r'(?:^|\s+)(?:投资收益|其中\S{1}投资收益)(?:\s+|$)'),\
                                   '投资收益：对联营企业和合营企业的投资收益': re.compile(r'(?:^|\s+)(?:对联营企业和合营企业的投资收益|其中\S{1}对联营企业和合营企业的投资收益)(?:\s+|$)'),\
                                   '投资收益：以摊余成本计量的金融资产终止确认收益': re.compile(r'(?:^|\s+)(?:以摊余成本计量的金融资产终止确认收益|其中\S{1}以摊余成本计量的金融资产终止确认收益)(?:\s+|$)'),\
                                   '汇兑收益*': re.compile(r'(?:^|\s+)(?:(\S?汇兑收益\S?)|(\S?其中\S{1}汇兑收益\S?))(?:\s+|$)'),\
                                   '净敞口套期收益': re.compile(r'(?:^|\s+)(?:净敞口套期收益|其中\S{1}净敞口套期收益)(?:\s+|$)'),\
                                   '公允价值变动收益': re.compile(r'(?:^|\s+)(?:公允价值变动收益|其中\S{1}公允价值变动收益)(?:\s+|$)'),\
                                   '信用减值损失': re.compile(r'(?:^|\s+)(?:信用减值损失|其中\S{1}信用减值损失)(?:\s+|$)'),\
                                   '资产减值损失': re.compile(r'(?:^|\s+)(?:资产减值损失|其中\S{1}资产减值损失)(?:\s+|$)'),\
                                   '资产处置收益': re.compile(r'(?:^|\s+)(?:资产处置收益|其中\S{1}资产处置收益)(?:\s+|$)'),\
                                   '营业利润': re.compile(r'(?:^|\s+)(?:二\S{1}营业利润|三\S{1}营业利润)(?:\s+|$)'),\
                                   '营业外收入': re.compile(r'(?:^|\s+)(?:加\S{1}营业外收入|营业外收入)(?:\s+|$)'),\
                                   '营业外支出': re.compile(r'(?:^|\s+)(?:减\S{1}营业外支出|营业外支出)(?:\s+|$)'),\
                                   '利润总额': re.compile(r'(?:^|\s+)(?:三\S{1}利润总额|四\S{1}利润总额)(?:\s+|$)'),\
                                   '所得税费用': re.compile(r'(?:^|\s+)所得税费用(?:\s+|$)'),\
                                   '净利润': re.compile(r'(?:^|\s+)(?:四\S{1}净利润|五\S{1}净利润)(?:\s+|$)'),\
                                   '持续经营净利润': re.compile(r'(?:^|\s+)持续经营净利润(?:\s+|$)'),\
                                   '终止经营净利润': re.compile(r'(?:^|\s+)终止经营净利润(?:\s+|$)'),\
                                   '*归属于母公司股东的净利润': re.compile(r'(?:^|\s+)(\S?归属于母公司股东的净利润\S?)(?:\s+|$)'),\
                                   '*少数股东损益': re.compile(r'(?:^|\s+)(\S?少数股东损益\S?)(?:\s+|$)'),\
                                   '其他综合收益的税后净额': re.compile(r'(?:^|\s+)(\S{1,3}其他综合收益的税后净额\S?)(?:\s+|$)'),\
                                   '*归属于母公司所有者的其他综合收益的税后净额': re.compile(r'(?:^|\s+)(\S?归属于母公司所有者的其他综合收益的税后净额\S?)(?:\s+|$)'),\
                                   '不能重分类进损益的其他综合收益': re.compile(r'(?:^|\s+)(\S?不能重分类进损益的其他综合收益\S?)(?:\s+|$)'),\
                                   '重新计量设定受益计划变动额': re.compile(r'(?:^|\s+)(\S?重新计量设定受益计划变动额\S?)(?:\s+|$)'),\
                                   '权益法下不能转损益的其他综合收益': re.compile(r'(?:^|\s+)(\S?权益法下不能转损益的其他综合收益\S?)(?:\s+|$)'),\
                                   '其他权益工具投资公允价值变动': re.compile(r'(?:^|\s+)(\S?其他权益工具投资公允价值变动\S?)(?:\s+|$)'),\
                                   '企业自身信用风险公允价值变动': re.compile(r'(?:^|\s+)(\S?企业自身信用风险公允价值变动\S?)(?:\s+|$)'),\
                                   '将重分类进损益的其他综合收益': re.compile(r'(?:^|\s+)(\S?将重分类进损益的其他综合收益\S?)(?:\s+|$)'),\
                                   '权益法下可转损益的其他综合收益': re.compile(r'(?:^|\s+)(\S?权益法下可转损益的其他综合收益\S?)(?:\s+|$)'),\
                                   '其他债权投资公允价值变动': re.compile(r'(?:^|\s+)(\S?其他债权投资公允价值变动\S?)(?:\s+|$)'),\
                                   '金融资产重分类计入其他综合收益的金额': re.compile(r'(?:^|\s+)(\S?金融资产重分类计入其他综合收益的金额\S?)(?:\s+|$)'),\
                                   '其他债权投资信用减值准备': re.compile(r'(?:^|\s+)(\S?其他债权投资信用减值准备\S?)(?:\s+|$)'),\
                                   '现金流量套期储备': re.compile(r'(?:^|\s+)(\S?现金流量套期储备\S?)(?:\s+|$)'),\
                                   '外币财务报表折算差额': re.compile(r'(?:^|\s+)(\S?外币财务报表折算差额\S?)(?:\s+|$)'),\
                                   '*归属于少数股东的其他综合收益的税后净额': re.compile(r'(?:^|\s+)(\S?归属于少数股东的其他综合收益的税后净额\S?)(?:\s+|$)'),\
                                   '综合收益总额': re.compile(r'(?:^|\s+)(\S{1,3}综合收益总额\S?)(?:\s+|$)'),\
                                   '*归属于母公司所有者的综合收益总额': re.compile(r'(?:^|\s+)(\S?归属于母公司所有者的综合收益总额\S?)(?:\s+|$)'),\
                                   '*归属于少数股东的综合收益总额': re.compile(r'(?:^|\s+)(\S?归属于少数股东的综合收益总额\S?)(?:\s+|$)'),\
                                   '基本每股收益': re.compile(r'(?:^|\s+)(\S?基本每股收益\S?)(?:\s+|$)'),\
                                   '稀释每股收益':re.compile(r'(?:^|\s+)(\S?稀释每股收益\S?)(?:\s+|$)')}

    num = 0
    display_income_dict = {}

    folder_path = request.get('data', {}).get('project_folder', '')
    file_path = request.get('data', {}).get('file_path', '')
    sheet_name = request.get('data', {}).get('sheet_name', '')
    items_name = int(request.get('data', {}).get('items_name', ''))
    items_this = int(request.get('data', {}).get('items_this', ''))
    items_previous = int(request.get('data', {}).get('items_previous', ''))

    file_extension = os.path.splitext(file_path)[1].lower()

    if file_extension == '.xlsx':
        import_income_statement_df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='openpyxl')
    elif file_extension == '.xls':
        import_income_statement_df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='xlrd')
    else:
        return

    data_income_statement_df = import_income_statement_df.iloc[:, [items_name, items_previous, items_this]].copy()
    data_income_statement_df.columns = ['item', 'previous_year', 'current_year']

    # 数据清洗：将数值列的字符串数据类型转换为浮点数类型
    data_income_statement_df['previous_year'] = data_income_statement_df['previous_year'].apply(convert_to_numeric)
    data_income_statement_df['current_year'] = data_income_statement_df['current_year'].apply(convert_to_numeric)

    for i in range(1, len(data_income_statement_df)):
        keywords = data_income_statement_df.iloc[i, 0]
        if pd.isna(keywords):                           # 处理 NaN/None
            keywords = ''
        elif isinstance(keywords, (float, int)):        # 处理 float/int
            keywords = str(keywords)

        for key in income_statement_list:
            if re.search(income_statement_regex_dict[key], keywords):
                previous = data_income_statement_df.iloc[i, 1]
                previous = 0 if pd.isna(previous) else previous
                end = data_income_statement_df.iloc[i, 2]
                end = 0 if pd.isna(end) else end
                income_statement_dict[key][0] = previous
                income_statement_dict[key][1] = end
                if previous != 0 or end != 0:
                    if pd.notna(previous) or pd.notna(end):
                        num += 1
                        display_income_dict[num] = [key, end, previous]
                        break

        else:
            key = '@未识别报表项目：' + keywords
            previous = data_income_statement_df.iloc[i, 1]
            previous = 0 if pd.isna(previous) else previous
            end = data_income_statement_df.iloc[i, 2]
            end = 0 if pd.isna(end) else end
            if previous != 0 or end != 0:
                if pd.notna(previous) or pd.notna(end):
                    num += 1
                    display_income_dict[num] = [key, end, previous]

    # 转换为DataFrame
    income_statement_data = pd.DataFrame(income_statement_dict).T

    # 确定保存路径
    export_path = os.path.join(folder_path, '项目数据', '利润表.xlsx')

    # 保存为xlsx文件
    income_statement_data.to_excel(export_path, sheet_name='Sheet1', engine='openpyxl')

    # 返回数据字典和保存路径
    return ['import_income_statement', display_income_dict]


def convert_to_numeric(value):
    
    if pd.isna(value):
        return 0.0

    try:
        return pd.to_numeric(str(value).strip().replace(',', '').replace(' ', ''))
    except:
        return 0.0