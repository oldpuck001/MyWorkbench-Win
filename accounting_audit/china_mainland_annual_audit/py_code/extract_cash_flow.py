# extract_cash_flow.py

import os
import re
import pandas as pd

def extract_cash_flow_xlsx(filepath, data_folder):
    cash_flow_statement_list = ['销售商品、提供劳务收到的现金',\
                                '客户存款和同业存放款项净增加额*', '向中央银行借款净增加额*', '向其他金融机构拆入资金净增加额*',\
                                '收到原保险合同保费取得的现金*', '收到再保业务现金净额*', '保户储金及投资款净增加额*',\
                                '收取利息、手续费及佣金的现金*', '拆入资金净增加额*', '回购业务资金净增加额*',\
                                '代理买卖证券收到的现金净额*',\
                                '收到的税费返还', '收到其他与经营活动有关的现金',\
                                '经营活动现金流入小计',\
                                '购买商品、接受劳务支付的现金',\
                                '客户贷款及垫款净增加额*', '存放中央银行和同业款项净增加额*', '支付原保险合同赔付款项的现金*',\
                                '拆出资金净增加额*', '支付利息、手续费及佣金的现金*', '支付保单红利的现金*',\
                                '支付给职工及为职工支付的现金', '支付的各项税费', '支付其他与经营活动有关的现金',\
                                '经营活动现金流出小计',\
                                '经营活动产生的现金流量净额',\
                                '收回投资收到的现金', '取得投资收益收到的现金', '处置固定资产、无形资产和其他长期资产收回的现金净额',\
                                '处置子公司及其他营业单位收到的现金净额', '收到其他与投资活动有关的现金',\
                                '投资活动现金流入小计',\
                                '购建固定资产、无形资产和其他长期资产支付的现金', '投资支付的现金',\
                                '质押贷款净增加额*',\
                                '取得子公司及其他营业单位支付的现金净额', '支付其他与投资活动有关的现金',\
                                '投资活动现金流出小计',\
                                '投资活动产生的现金流量净额',\
                                '吸收投资收到的现金', '吸收投资收到的现金：子公司吸收少数股东投资收到的现金',\
                                '取得借款收到的现金', '收到其他与筹资活动有关的现金',\
                                '筹资活动现金流入小计',\
                                '偿还债务支付的现金',\
                                '分配股利、利润或偿付利息支付的现金', '分配股利、利润或偿付利息支付的现金：子公司支付给少数股东的股利、利润',\
                                '支付其他与筹资活动有关的现金',\
                                '筹资活动现金流出小计',\
                                '筹资活动产生的现金流量净额',\
                                '四、汇率变动对现金及现金等价物的影响',\
                                '五、现金及现金等价物净增加额',\
                                '期初现金及现金等价物余额',\
                                '六、期末现金及现金等价物余额']

    cash_flow_statement_dict = {key: [0, 0, 0, 0, 0, 0] for key in cash_flow_statement_list}

    cash_flow_regex_dict = {'销售商品、提供劳务收到的现金': re.compile(r'\S*销售商品\S{1}提供劳务收到的现金\S*'),\
                            '客户存款和同业存放款项净增加额*': re.compile(r'\S*客户存款和同业存放款项净增加额\S*'),\
                            '向中央银行借款净增加额*': re.compile(r'\S*向中央银行借款净增加额\S*'),\
                            '向其他金融机构拆入资金净增加额*': re.compile(r'\S*向其他金融机构拆入资金净增加额\S*'),\
                            '收到原保险合同保费取得的现金*': re.compile(r'\S*收到原保险合同保费取得的现金\S*'),\
                            '收到再保业务现金净额*': re.compile(r'\S*收到再保业务现金净额\S*'),\
                            '保户储金及投资款净增加额*': re.compile(r'\S*保户储金及投资款净增加额\S*'),\
                            '收取利息、手续费及佣金的现金*': re.compile(r'\S*收取利息\S{1}手续费及佣金的现金\S*'),\
                            '拆入资金净增加额*': re.compile(r'\S*拆入资金净增加额\S*'),\
                            '回购业务资金净增加额*': re.compile(r'\S*回购业务资金净增加额\S*'),\
                            '代理买卖证券收到的现金净额*': re.compile(r'\S*代理买卖证券收到的现金净额\S*'),\
                            '收到的税费返还': re.compile(r'\S*收到的税费返还\S*'),\
                            '收到其他与经营活动有关的现金': re.compile(r'\S*收到其他与经营活动有关的现金\S*'),\
                            '经营活动现金流入小计': re.compile(r'\S*经营活动现金流入小计\S*'),\
                            '购买商品、接受劳务支付的现金': re.compile(r'\S*购买商品\S{1}接受劳务支付的现金\S*'),\
                            '客户贷款及垫款净增加额*': re.compile(r'\S*客户贷款及垫款净增加额\S*'),\
                            '存放中央银行和同业款项净增加额*': re.compile(r'\S*存放中央银行和同业款项净增加额\S*'),\
                            '支付原保险合同赔付款项的现金*': re.compile(r'\S*支付原保险合同赔付款项的现金\S*'),\
                            '拆出资金净增加额*': re.compile(r'\S*拆出资金净增加额\S*'),\
                            '支付利息、手续费及佣金的现金*': re.compile(r'\S*支付利息\S{1}手续费及佣金的现金\S*'),\
                            '支付保单红利的现金*': re.compile(r'\S*支付保单红利的现金\S*'),\
                            '支付给职工及为职工支付的现金': re.compile(r'\S*支付给职工及为职工支付的现金\S*|\S*支付给职工以及为职工支付的现金\S*'),\
                            '支付的各项税费': re.compile(r'\S*支付的各项税费\S*'),\
                            '支付其他与经营活动有关的现金': re.compile(r'\S*支付其他与经营活动有关的现金\S*'),\
                            '经营活动现金流出小计': re.compile(r'\S*经营活动现金流出小计\S*'),\
                            '经营活动产生的现金流量净额': re.compile(r'\S*经营活动产生的现金流量净额\S*'),\
                            '收回投资收到的现金': re.compile(r'\S*收回投资收到的现金\S*'),\
                            '取得投资收益收到的现金': re.compile(r'\S*取得投资收益收到的现金\S*'),\
                            '处置固定资产、无形资产和其他长期资产收回的现金净额': re.compile(r'\S*处置固定资产\S{1}无形资产和其他长期资产收回的现金净额\S*'),\
                            '处置子公司及其他营业单位收到的现金净额': re.compile(r'\S*处置子公司及其他营业单位收到的现金净额\S*'),\
                            '收到其他与投资活动有关的现金': re.compile(r'\S*收到其他与投资活动有关的现金\S*'),\
                            '投资活动现金流入小计': re.compile(r'\S*投资活动现金流入小计\S*'),\
                            '购建固定资产、无形资产和其他长期资产支付的现金': re.compile(r'\S*购建固定资产\S{1}无形资产和其他长期资产支付的现金\S*'),\
                            '投资支付的现金': re.compile(r'\S*投资支付的现金\S*'),\
                            '质押贷款净增加额*': re.compile(r'\S*质押贷款净增加额\S*'),\
                            '取得子公司及其他营业单位支付的现金净额': re.compile(r'\S*取得子公司及其他营业单位支付的现金净额\S*'),\
                            '支付其他与投资活动有关的现金': re.compile(r'\S*支付其他与投资活动有关的现金\S*'),\
                            '投资活动现金流出小计': re.compile(r'\S*投资活动现金流出小计\S*'),\
                            '投资活动产生的现金流量净额': re.compile(r'\S*投资活动产生的现金流量净额\S*'),\
                            '吸收投资收到的现金': re.compile(r'\S*吸收投资收到的现金\S*'),\
                            '吸收投资收到的现金：子公司吸收少数股东投资收到的现金': re.compile(r'\S*子公司吸收少数股东投资收到的现金\S*'),\
                            '取得借款收到的现金': re.compile(r'\S*取得借款收到的现金\S*'),\
                            '收到其他与筹资活动有关的现金': re.compile(r'\S*收到其他与筹资活动有关的现金\S*'),\
                            '筹资活动现金流入小计': re.compile(r'\S*筹资活动现金流入小计\S*'),\
                            '偿还债务支付的现金': re.compile(r'\S*偿还债务支付的现金\S*'),\
                            '分配股利、利润或偿付利息支付的现金': re.compile(r'\S*分配股利\S{1}利润或偿付利息支付的现金\S*'),\
                            '分配股利、利润或偿付利息支付的现金：子公司支付给少数股东的股利、利润': re.compile(r'\S*子公司支付给少数股东的股利\S{1}利润\S*'),\
                            '支付其他与筹资活动有关的现金': re.compile(r'\S*支付其他与筹资活动有关的现金\S*'),\
                            '筹资活动现金流出小计': re.compile(r'\S*筹资活动现金流出小计\S*'),\
                            '筹资活动产生的现金流量净额': re.compile(r'\S*筹资活动产生的现金流量净额\S*'),\
                            '四、汇率变动对现金及现金等价物的影响': re.compile(r'\S*汇率变动对现金及现金等价物的影响\S*'),\
                            '五、现金及现金等价物净增加额': re.compile(r'\S*现金及现金等价物净增加额\S*'),\
                            '期初现金及现金等价物余额': re.compile(r'\S*期初现金及现金等价物余额\S*'),\
                            '六、期末现金及现金等价物余额': re.compile(r'\S*期末现金及现金等价物余额\S*')
                            }

    display_cash_dict = {}

    num = 0
    
    file_extension = os.path.splitext(filepath)[1].lower()
    if file_extension == '.xlsx':
        cash_flow_df = pd.read_excel(filepath, engine='openpyxl')
    elif file_extension == '.xls':
        cash_flow_df = pd.read_excel(filepath, engine='xlrd')

    # 检查文件有多少列及表头所在的行次
    num_columns = cash_flow_df.shape[1]

    # 使用正则表达式来匹配前后和中间可能存在空格的“项目”
    pattern_project_cash = re.compile(r'\s*项\s*目\s*')
    # 查找含有“项目”变体的行
    start_row_index = cash_flow_df[cash_flow_df.astype(str).apply(lambda row: row.str.contains(pattern_project_cash).any(), axis=1)].index[0]

    # 正则表达式匹配“本期金额，本期数”关键字
    pattern_this_cash = re.compile(r'\s*本\s*期\s*金\s*额\s*|\s*本\s*期\s*数\s*')

    # 正则表达式匹配“上期余额，上期数”关键字
    pattern_previous_cash = re.compile(
        r'\s*上\s*期\s*金\s*额\s*|'
        r'\s*上\s*期\s*数\s*'
    )

    # 列表来保存找到的列索引
    project_cash_columns = []
    this_cash_columns = []
    previous_cash_columns = []

    # 遍历列，寻找匹配的列索引
    for n in range(num_columns):
        header = str(cash_flow_df.iloc[start_row_index, n])
        if re.search(pattern_project_cash, header):
            project_cash_columns.append(n)
        elif re.search(pattern_this_cash, header):
            this_cash_columns.append(n)
        elif re.search(pattern_previous_cash, header):
            previous_cash_columns.append(n)

    for n in project_cash_columns:
        index = project_cash_columns.index(n)
        for i in range(start_row_index + 1, len(cash_flow_df)):
            keywords = cash_flow_df.iloc[i, n]
            if isinstance(keywords, str):
                keywords = keywords.replace(' ', '')
            if isinstance(keywords, float):
                continue
            for key in cash_flow_statement_list:
                if re.search(cash_flow_regex_dict[key], keywords):
                    previous = cash_flow_df.iloc[i, previous_cash_columns[index]]
                    previous = 0 if pd.isna(previous) else previous
                    this = cash_flow_df.iloc[i, this_cash_columns[index]]
                    this = 0 if pd.isna(this) else this
                    cash_flow_statement_dict[key][0] = previous
                    cash_flow_statement_dict[key][1] = this
                    if previous != 0 or this != 0:
                        if pd.notna(previous) or pd.notna(this):
                            num += 1
                            display_cash_dict[num] = [key, this, previous]
                            break
            else:
                if keywords in ['一、经营活动产生的现金流量：', '二、投资活动产生的现金流量：', '三、筹资活动产生的现金流量：']:
                    key = keywords
                else:
                    key = '@未识别报表项目：' + keywords
                previous = cash_flow_df.iloc[i, previous_cash_columns[index]]
                previous = 0 if pd.isna(previous) else previous
                this = cash_flow_df.iloc[i, this_cash_columns[index]]
                this = 0 if pd.isna(this) else this
                if previous != 0 or this != 0:
                    if pd.notna(previous) or pd.notna(this):
                        num += 1
                        display_cash_dict[num] = [key, this, previous]

    # 转换为DataFrame
    cash_flow_data = pd.DataFrame(cash_flow_statement_dict)

    # 确定保存路径
    save_path = os.path.join(data_folder, 'cash_flow_statement.csv')
        
    # 保存为CSV文件
    cash_flow_data.to_csv(save_path, index=False, encoding='utf-8-sig')

    # 返回数据字典和保存路径
    return display_cash_dict, save_path