# import_balance_sheet.py

import os
import re
import pandas as pd

def select_balance_sheet(request):

    file_path = request.get("data", {}).get("file_path", "")

    sheet_file = pd.ExcelFile(file_path)                                   # 使用pandas讀取Excel文件
    sheetnames = sheet_file.sheet_names                                    # 獲取所有工作表名稱

    return ['select_balance_sheet', sheetnames]


def import_balance_sheet(request):

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

    balance_sheet_dict = {key: [0, 0] for key in balance_sheet_list}

    balance_sheet_regex_dict = {'货币资金': re.compile(r'(?:^|\s+)货币资金(?:\s+|$)'),\
                                '结算备付金*': re.compile(r'(?:^|\s+)(\S?结算备付金\S?)(?:\s+|$)'),\
                                '拆出资金*': re.compile(r'(?:^|\s+)(\S?拆出资金\S?)(?:\s+|$)'),\
                                '交易性金融资产': re.compile(r'(?:^|\s+)交易性金融资产(?:\s+|$)'),\
                                '△以公允价值计量且其变动计入当期损益的金融资产': re.compile(r'(?:^|\s+)(\S?以公允价值计量且其变动计入当期损益的金融资产\S?)(?:\s+|$)'),\
                                '衍生金融资产': re.compile(r'(?:^|\s+)衍生金融资产(?:\s+|$)'),\
                                '应收票据': re.compile(r'(?:^|\s+)应收票据(?:\s+|$)'),\
                                '应收账款': re.compile(r'(?:^|\s+)(?:应收账款|应收帐款)(?:\s+|$)'),\
                                '应收款项融资': re.compile(r'(?:^|\s+)应收款项融资(?:\s+|$)'),\
                                '预付款项': re.compile(r'(?:^|\s+)(?:预付款项|预付账款|预付帐款)(?:\s+|$)'),\
                                '应收保费*': re.compile(r'(?:^|\s+)(\S?应收保费\S?)(?:\s+|$)'),\
                                '应收分保账款*': re.compile(r'(?:^|\s+)(\S?应收分保账款\S?)(?:\s+|$)'),\
                                '应收分保合同准备金*': re.compile(r'(?:^|\s+)(\S?应收分保合同准备金\S?)(?:\s+|$)'),\
                                '其他应收款': re.compile(r'(?:^|\s+)(?:其他应收款|其它应收款)(?:\s+|$)'),\
                                '买入返售金融资产*': re.compile(r'(?:^|\s+)(\S?买入返售金融资产\S?)(?:\s+|$)'),\
                                '存货': re.compile(r'(?:^|\s+)存货(?:\s+|$)'),\
                                '合同资产': re.compile(r'(?:^|\s+)合同资产(?:\s+|$)'),\
                                '持有待售资产': re.compile(r'(?:^|\s+)持有待售资产(?:\s+|$)'),\
                                '一年内到期的非流动资产': re.compile(r'(?:^|\s+)一年内到期的非流动资产(?:\s+|$)'),\
                                '其他流动资产': re.compile(r'(?:^|\s+)(?:其他流动资产|其它流动资产)(?:\s+|$)'),\
                                '流动资产合计': re.compile(r'(?:^|\s+)(流动资产合计\S?)(?:\s+|$)'),\
                                '发放贷款和垫款*': re.compile(r'(?:^|\s+)(\S?发放贷款和垫款\S?)(?:\s+|$)'),\
                                '债权投资': re.compile(r'(?:^|\s+)债权投资(?:\s+|$)'),\
                                '△可供出售金融资产': re.compile(r'(?:^|\s+)(\S?可供出售金融资产\S?)(?:\s+|$)'),\
                                '其他债权投资': re.compile(r'(?:^|\s+)(?:其他债权投资|其它债权投资)(?:\s+|$)'),\
                                '△持有至到期投资': re.compile(r'(?:^|\s+)(\S?持有至到期投资\S?)(?:\s+|$)'),\
                                '长期应收款': re.compile(r'(?:^|\s+)长期应收款(?:\s+|$)'),\
                                '长期股权投资': re.compile(r'(?:^|\s+)长期股权投资(?:\s+|$)'),\
                                '其他权益工具投资': re.compile(r'(?:^|\s+)(?:其他权益工具投资|其它权益工具投资)(?:\s+|$)'),\
                                '其他非流动金融资产': re.compile(r'(?:^|\s+)(?:其他非流动金融资产|其它非流动金融资产)(?:\s+|$)'),\
                                '投资性房地产': re.compile(r'(?:^|\s+)投资性房地产(?:\s+|$)'),\
                                '固定资产': re.compile(r'(?:^|\s+)(?:固定资产|固定资产净额|固定资产净值)(?:\s+|$)'),\
                                '在建工程': re.compile(r'(?:^|\s+)在建工程(?:\s+|$)'),\
                                '生产性生物资产': re.compile(r'(?:^|\s+)生产性生物资产(?:\s+|$)'),\
                                '油气资产': re.compile(r'(?:^|\s+)油气资产(?:\s+|$)'),\
                                '使用权资产': re.compile(r'(?:^|\s+)使用权资产(?:\s+|$)'),\
                                '无形资产': re.compile(r'(?:^|\s+)无形资产(?:\s+|$)'),\
                                '开发支出': re.compile(r'(?:^|\s+)开发支出(?:\s+|$)'),\
                                '商誉': re.compile(r'(?:^|\s+)商誉(?:\s+|$)'),\
                                '长期待摊费用': re.compile(r'(?:^|\s+)长期待摊费用(?:\s+|$)'),\
                                '递延所得税资产': re.compile(r'(?:^|\s+)(?:递延所得税资产|递延所得税借项)(?:\s+|$)'),\
                                '其他非流动资产': re.compile(r'(?:^|\s+)(?:其他非流动资产|其它非流动资产)(?:\s+|$)'),\
                                '非流动资产合计': re.compile(r'(?:^|\s+)(非流动资产合计\S?)(?:\s+|$)'),\
                                '资产总计': re.compile(r'(?:^|\s+)(资产总计\S?)(?:\s+|$)'),\
                                '短期借款': re.compile(r'(?:^|\s+)短期借款(?:\s+|$)'),\
                                '向中央银行借款*': re.compile(r'(?:^|\s+)(\S?向中央银行借款\S?)(?:\s+|$)'),\
                                '拆入资金*': re.compile(r'(?:^|\s+)(\S?拆入资金\S?)(?:\s+|$)'),\
                                '交易性金融负债': re.compile(r'(?:^|\s+)交易性金融负债(?:\s+|$)'),\
                                '△以公允价值计量且其变动计入当期损益的金融负债': re.compile(r'(?:^|\s+)(\S?以公允价值计量且其变动计入当期损益的金融负债\S?)(?:\s+|$)'),\
                                '衍生金融负债': re.compile(r'(?:^|\s+)衍生金融负债(?:\s+|$)'),\
                                '应付票据': re.compile(r'(?:^|\s+)应付票据(?:\s+|$)'),\
                                '应付账款': re.compile(r'(?:^|\s+)(?:应付账款|应付帐款)(?:\s+|$)'),\
                                '预收款项': re.compile(r'(?:^|\s+)(?:预收款项|预收账款|预收帐款)(?:\s+|$)'),\
                                '合同负债': re.compile(r'(?:^|\s+)合同负债(?:\s+|$)'),\
                                '卖出回购金融资产款*': re.compile(r'(?:^|\s+)(\S?卖出回购金融资产款\S?)(?:\s+|$)'),\
                                '吸收存款及同业存放*': re.compile(r'(?:^|\s+)(\S?吸收存款及同业存放\S?)(?:\s+|$)'),\
                                '代理买卖证券款*': re.compile(r'(?:^|\s+)(\S?代理买卖证券款\S?)(?:\s+|$)'),\
                                '代理承销证券款*': re.compile(r'(?:^|\s+)(\S?代理承销证券款\S?)(?:\s+|$)'),\
                                '应付职工薪酬': re.compile(r'(?:^|\s+)应付职工薪酬(?:\s+|$)'),\
                                '应交税费': re.compile(r'(?:^|\s+)应交税费(?:\s+|$)'),\
                                '其他应付款': re.compile(r'(?:^|\s+)(?:其他应付款|其它应付款)(?:\s+|$)'),\
                                '应付手续费及佣金*': re.compile(r'(?:^|\s+)(\S?应付手续费及佣金\S?)(?:\s+|$)'),\
                                '应付分保账款*': re.compile(r'(?:^|\s+)(\S?应付分保账款\S?)(?:\s+|$)'),\
                                '持有待售负债': re.compile(r'(?:^|\s+)持有待售负债(?:\s+|$)'),\
                                '一年内到期的非流动负债': re.compile(r'(?:^|\s+)一年内到期的非流动负债(?:\s+|$)'),\
                                '其他流动负债': re.compile(r'(?:^|\s+)(?:其他流动负债|其它流动负债)(?:\s+|$)'),\
                                '流动负债合计': re.compile(r'(?:^|\s+)(流动负债合计\S?)(?:\s+|$)'),\
                                '保险合同准备金*': re.compile(r'(?:^|\s+)(\S?保险合同准备金\S?)(?:\s+|$)'),\
                                '长期借款': re.compile(r'(?:^|\s+)长期借款(?:\s+|$)'),\
                                '应付债券': re.compile(r'(?:^|\s+)应付债券(?:\s+|$)'),\
                                '租赁负债': re.compile(r'(?:^|\s+)租赁负债(?:\s+|$)'),\
                                '长期应付款': re.compile(r'(?:^|\s+)长期应付款(?:\s+|$)'),\
                                '预计负债': re.compile(r'(?:^|\s+)预计负债(?:\s+|$)'),\
                                '递延收益': re.compile(r'(?:^|\s+)递延收益(?:\s+|$)'),\
                                '递延所得税负债': re.compile(r'(?:^|\s+)(?:递延所得税负债|递延所得税贷项)(?:\s+|$)'),\
                                '其他非流动负债': re.compile(r'(?:^|\s+)(?:其他非流动负债|其它非流动负债)(?:\s+|$)'),\
                                '非流动负债合计': re.compile(r'(?:^|\s+)(非流动负债合计\S?)(?:\s+|$)'),\
                                '负债合计': re.compile(r'(?:^|\s+)(负债合计\S?)(?:\s+|$)'),\
                                '实收资本（或股本）': re.compile(r'(?:^|\s+)(?:实收资本\S{1,2}股本\S{1}$|实收资本$|股本$)(?:\s+|$)'),\
                                '#减：已归还投资': re.compile(r'(?:^|\s+)(\S?减\S?已归还投资\S?)(?:\s+|$)'),\
                                '#实收资本（或股本）净额': re.compile(r'(?:^|\s+)(\S?实收资本\S{1,2}股本\S{1}净额$|\S*实收资本净额$|\S*股本净额$\S?)(?:\s+|$)'),\
                                '其他权益工具': re.compile(r'(?:^|\s+)(?:其他权益工具|其它权益工具)(?:\s+|$)'),\
                                '资本公积': re.compile(r'(?:^|\s+)资本公积(?:\s+|$)'),\
                                '减：库存股': re.compile(r'(?:^|\s+)(减\S?库存股)(?:\s+|$)'),\
                                '其他综合收益': re.compile(r'(?:^|\s+)(?:其他综合收益|其它综合收益)(?:\s+|$)'),\
                                '专项储备': re.compile(r'(?:^|\s+)专项储备(?:\s+|$)'),\
                                '盈余公积': re.compile(r'(?:^|\s+)盈余公积(?:\s+|$)'),\
                                '一般风险准备*': re.compile(r'(?:^|\s+)(\S?一般风险准备\S?)(?:\s+|$)'),\
                                '未分配利润': re.compile(r'(?:^|\s+)未分配利润(?:\s+|$)'),\
                                '*归属于母公司所有者权益（或股东权益）合计': re.compile(r'(?:^|\s+)(\S?归属于母公司所有者权益\S{1,2}股东权益\S{1}合计\S{0,2}|\S?归属于母公司所有者权益合计\S{0,2}|\S?归属于母公司股东权益合计\S{0,2})(?:\s+|$)'),\
                                '*少数股东权益': re.compile(r'(?:^|\s+)(\S?\S?)少数股东权益(?:\s+|$)'),\
                                '所有者权益（或股东权益）合计': re.compile(r'(?:^|\s+)(?:所有者权益\S{1,2}股东权益\S{1}合计\S?|所有者权益合计\S?|股东权益合计\S?)(?:\s+|$)'),\
                                '负债和所有者权益（或股东权益）总计': re.compile(r'(?:^|\s+)(?:负债和所有者权益\S{1,2}股东权益\S{1}总计\S?|负债和所有者权益总计\S?|负债和股东权益总计\S?)(?:\s+|$)')}

    num = 0
    display_balance_dict = {}

    folder_path = request.get('data', {}).get('project_folder', '')
    file_path = request.get('data', {}).get('file_path', '')
    sheet_name = request.get('data', {}).get('sheet_name', '')
    assets_name = int(request.get('data', {}).get('assets_name', ''))
    assets_this = int(request.get('data', {}).get('assets_this', ''))
    assets_previous = int(request.get('data', {}).get('assets_previous', ''))
    liabilities_name = int(request.get('data', {}).get('liabilities_name', ''))
    liabilities_this = int(request.get('data', {}).get('liabilities_this', ''))
    liabilities_previous = int(request.get('data', {}).get('liabilities_previous', ''))

    file_extension = os.path.splitext(file_path)[1].lower()

    if file_extension == '.xlsx':
        import_balance_sheet_df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='openpyxl')
    elif file_extension == '.xls':
        import_balance_sheet_df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='xlrd')
    else:
        return

    if assets_name == liabilities_name:

        data_balance_sheet_df = import_balance_sheet_df.iloc[:, [assets_name, assets_previous, assets_this]].copy()
        data_balance_sheet_df.columns = ['item', 'previous_year', 'current_year']

    else:

        data_balance_sheet_df_assets = import_balance_sheet_df.iloc[:, [assets_name, assets_previous, assets_this]].copy()
        data_balance_sheet_df_assets.columns = ['item', 'previous_year', 'current_year']

        data_balance_sheet_df_liabilities = import_balance_sheet_df.iloc[:, [liabilities_name, liabilities_previous, liabilities_this]].copy()
        data_balance_sheet_df_liabilities.columns = ['item', 'previous_year', 'current_year']

        data_balance_sheet_df = pd.concat([data_balance_sheet_df_assets, data_balance_sheet_df_liabilities], axis=0)

    # 数据清洗：将数值列的字符串数据类型转换为浮点数类型
    data_balance_sheet_df['previous_year'] = data_balance_sheet_df['previous_year'].apply(convert_to_numeric)
    data_balance_sheet_df['current_year'] = data_balance_sheet_df['current_year'].apply(convert_to_numeric)

    for i in range(1, len(data_balance_sheet_df)):
        keywords = data_balance_sheet_df.iloc[i, 0]
        if pd.isna(keywords):                           # 处理 NaN/None
            keywords = ''
        elif isinstance(keywords, (float, int)):        # 处理 float/int
            keywords = str(keywords)

        for key in balance_sheet_list:

            if key in ['应付债券：优先股', '其他权益工具：优先股']:
                if re.search(re.compile(r'\S*其中\S{1}优先股\S*'), keywords):
                    if data_balance_sheet_df.iloc[i-1, 0].replace(' ', '') == '应付债券':
                        key = '应付债券：优先股'
                    elif data_balance_sheet_df.iloc[i-1, 0].replace(' ', '') != '其他权益工具':
                        key = '其他权益工具：优先股'
                    previous = data_balance_sheet_df.iloc[i, 1]
                    previous = 0 if pd.isna(previous) else previous
                    end = data_balance_sheet_df.iloc[i, 2]
                    end = 0 if pd.isna(end) else end
                    balance_sheet_dict[key][0] = previous
                    balance_sheet_dict[key][1] = end
                    if previous != 0 or end != 0:
                        if pd.notna(previous) or pd.notna(end):
                            num += 1
                            display_balance_dict[num] = [key, end, previous]
                            break

            elif key in ['应付债券：永续债', '其他权益工具：永续债']:
                if re.search(re.compile(r'\S*永续债\S*'), keywords):
                    if data_balance_sheet_df.iloc[i-2, 0].replace(' ', '') == '应付债券':
                        key = '应付债券：永续债'
                    elif data_balance_sheet_df.iloc[i-2, 0].replace(' ', '') != '其他权益工具':
                        key = '其他权益工具：永续债'
                    previous = data_balance_sheet_df.iloc[i, 1]
                    previous = 0 if pd.isna(previous) else previous
                    end = data_balance_sheet_df.iloc[i, 2]
                    end = 0 if pd.isna(end) else end
                    balance_sheet_dict[key][0] = previous
                    balance_sheet_dict[key][1] = end
                    if previous != 0 or end != 0:
                        if pd.notna(previous) or pd.notna(end):
                            num += 1
                            display_balance_dict[num] = [key, end, previous]
                            break

            elif re.search(balance_sheet_regex_dict[key], keywords):
                previous = data_balance_sheet_df.iloc[i, 1]
                previous = 0 if pd.isna(previous) else previous
                end = data_balance_sheet_df.iloc[i, 2]
                end = 0 if pd.isna(end) else end
                balance_sheet_dict[key][0] = previous
                balance_sheet_dict[key][1] = end
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
            previous = data_balance_sheet_df.iloc[i, 1]
            previous = 0 if pd.isna(previous) else previous
            end = data_balance_sheet_df.iloc[i, 2]
            end = 0 if pd.isna(end) else end
            if previous != 0 or end != 0:
                if pd.notna(previous) or pd.notna(end):
                    num += 1
                    display_balance_dict[num] = [key, end, previous]

    # 转换为DataFrame
    balance_sheet_data = pd.DataFrame(balance_sheet_dict).T

    # 确定保存路径
    export_path = os.path.join(folder_path, '项目数据', '资产负债表.xlsx')

    # 保存为xlsx文件
    balance_sheet_data.to_excel(export_path, sheet_name='Sheet1', engine='openpyxl')

    # 返回数据字典和保存路径
    return ['import_balance_sheet', display_balance_dict]


def convert_to_numeric(value):
    
    if pd.isna(value):
        return 0.0

    try:
        return pd.to_numeric(str(value).strip().replace(',', '').replace(' ', ''))
    except:
        return 0.0