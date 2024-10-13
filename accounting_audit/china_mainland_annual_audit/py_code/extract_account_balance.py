# extract_account_balance.py

import os
import pandas as pd

def extract_account_balance_xlsx(filepath, data_folder):

    account_csv_list = ['库存现金', '银行存款', '存放中央银行款项', '其他货币资金',\
                        '结算备付金',\
                        '存放同业',\
                        '交易性金融资产',\
                        '衍生工具'
                        '应收票据',\
                        '应收账款', '坏账准备',\
                        '预付账款',\
                        '应收分保账款',\
                        '应收分保合同准备金',\
                        '其他应收款', '应收股利', '应收利息',\
                        '买入返售金融资产',\
                        '材料采购', '在途物资', '原材料', '库存商品', '周转材料', '委托加工物资', '发出商品', '生产成本', '受托代销商品',\
                        '受托代销商品款', '存货跌价准备',\
                        '商品进销差价', '材料成本差异',\
                        '合同资产', '合同负债',\
                        '存出保证金', '应收代位追偿款',\
                        '贷款', '贷款损失准备',\
                        '债权投资', '债券投资减值准备',\
                        '长期应收款', '未实现融资收益',\
                        '长期股权投资', '长期股权投资减值准备',\
                        '其他权益工具投资',\
                        '投资性房地产', '投资性房地产累计折旧', '投资性房地产减值准备',\
                        '固定资产', '累计折旧', '固定资产减值准备', '固定资产清理',\
                        '在建工程', '在建工程减值准备', '工程物资', '工程物资减值准备',\
                        '生产性生物资产', '生产性生物资产累计折旧', '生产性生物资产减值准备',\
                        '油气资产', '累计折耗', '油气资产减值准备',\
                        '使用权资产', '使用权资产累计折旧', '使用权资产减值准备',\
                        '无形资产', '累计摊销', '无形资产减值准备',\
                        '研发支出-资本化支出',\
                        '商誉', '商誉减值准备',\
                        '长期待摊费用',\
                        '递延所得税资产',\
                        '短期借款',\
                        '向中央银行借款',\
                        '拆入资金',\
                        '交易性金融负债',\
                        '应付票据',\
                        '应付账款',\
                        '预收账款',\
                        '卖出回购金融资产款',\
                        '吸收存款', '同业存放',\
                        '代理买卖证券款', '代理承销证券款',\
                        '应付职工薪酬',\
                        '应交税费',\
                        '其他应付款', '应付利息', '应付股利',\
                        '应付分保账款',\
                        '持有待售负债',\
                        '未到期责任准备金', '保险责任准备金',\
                        '长期借款',\
                        '应付债券',\
                        '租赁负债',\
                        '长期应付款', '未确认融资费用', '专项应付款',\
                        '预计负债'\
                        '递延收益',\
                        '递延所得税负债',\
                        '实收资本',\
                        '其他权益工具',\
                        '资本公积',\
                        '库存股',\
                        '其他综合收益',\
                        '专项储备',\
                        '盈余公积',\
                        '一般风险准备',\
                        '本年利润', '利润分配',\
                        '主营业务收入', '其他业务收入',\
                        '利息收入',\
                        '保费收入',\
                        '手续费及佣金收入',\
                        '主营业务成本', '其他业务成本',\
                        '利息支出',\
                        '手续费及佣金支出',\
                        '退保金',\
                        '赔付支出', '摊回赔付支出',\
                        '提取保险责任准备金', '摊回保险责任准备金',\
                        '保单红利支出',\
                        '分保费用', '摊回分保费用',\
                        '税金及附加',\
                        '销售费用',\
                        '管理费用',\
                        '研发支出-费用化支出',\
                        '财务费用',\
                        '其他收益',\
                        '投资收益',\
                        '汇兑损益',\
                        '净敞口套期损益'
                        '公允价值变动损益',\
                        '资产减值损失',\
                        '资产减值损失',\
                        '资产处置损益',\
                        '营业外收入',\
                        '营业外支出',\
                        '所得税费用']

    empty_df = pd.DataFrame()
    for csv_name in account_csv_list:
        output_csv_path = os.path.join(data_folder, f'{csv_name}.csv')
        empty_df.to_csv(output_csv_path, index=False, encoding='utf-8-sig')

    account_balance_dict = {}

    file_extension = os.path.splitext(filepath)[1].lower()
    if file_extension == '.xlsx':
        account_balance_df = pd.read_excel(filepath, engine='openpyxl')
    elif file_extension == '.xls':
        account_balance_df = pd.read_excel(filepath, engine='xlrd')

    # Convert the '科目编码' column to string type
    account_balance_df['科目编码'] = account_balance_df['科目编码'].astype(str)
    # Apply the conversion function to the relevant columns
    for col in account_balance_df.columns[2:]:
        account_balance_df[col] = account_balance_df[col].apply(convert_to_numeric)

    # Find the minimum length of the subject code
    min_length = account_balance_df['科目编码'].apply(determine_subject_length).min()

    # Traverse the DataFrame to find primary subjects
    for index, row in account_balance_df.iterrows():
        subject_name = row['科目名称']
        if subject_name in account_csv_list:
            subject_code = str(row['科目编码'])
            subject_length = determine_subject_length(subject_code)
            if subject_length == min_length:
                account_balance_dict[subject_name] = []
                account_balance_dict[subject_name].append(row['期初借方'])
                account_balance_dict[subject_name].append(row['期初贷方'])
                account_balance_dict[subject_name].append(row['本期借方'])
                account_balance_dict[subject_name].append(row['本期贷方'])
                account_balance_dict[subject_name].append(row['期末借方'])
                account_balance_dict[subject_name].append(row['期末贷方'])
                # Find all matching codes with the same prefix
                matching_codes = account_balance_df[account_balance_df['科目编码'].str.startswith(subject_code)]
                # Find the maximum length of these matching codes
                max_length = matching_codes['科目编码'].apply(determine_subject_length).max()
                # Filter to get the end-level subjects with the maximum length
                final_level_subjects = matching_codes[matching_codes['科目编码'].apply(determine_subject_length) == max_length]
                # Output the end-level subjects to a CSV file named by the primary subject code
                output_csv_path = os.path.join(data_folder, f'{subject_name}.csv')
                final_level_subjects.to_csv(output_csv_path, index=False, encoding='utf-8-sig')

    # 确定保存路径
    save_path = os.path.join(data_folder, 'account_balance_statement.xlsx')

    # 保存为xlsx文件
    account_balance_df.to_excel(save_path, index=False, sheet_name='Sheet1', engine='openpyxl')

    return {'message': '导入成功。'}, save_path

def convert_to_numeric(value):
    try:
        return pd.to_numeric(value.replace(',', ''))
    except AttributeError:
        return 0 if pd.isna(value) else value

def determine_subject_length(code):
    return len(str(code))