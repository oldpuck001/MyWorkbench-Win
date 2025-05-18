# import_cash_flow_statement.py

import os
import re
import pandas as pd

def select_cash_flow_statement(request):

    file_path = request.get("data", {}).get("file_path", "")

    sheet_file = pd.ExcelFile(file_path)                                   # 使用pandas讀取Excel文件
    sheetnames = sheet_file.sheet_names                                    # 獲取所有工作表名稱

    return ['select_cash_flow_statement', sheetnames]


def import_cash_flow_statement(request):

    cash_flow_statement_list = ['销售商品、提供劳务收到的现金',\
                                '客户存款和同业存放款项净增加额*', '向中央银行借款净增加额*', '向其他金融机构拆入资金净增加额*',\
                                '收到原保险合同保费取得的现金*', '收到再保业务现金净额*', '保户储金及投资款净增加额*',\
                                '收取利息、手续费及佣金的现金*', '拆入资金净增加额*', '回购业务资金净增加额*',\
                                '代理买卖证券收到的现金净额*',\
                                '收到的税费返还',\
                                '收到其他与经营活动有关的现金',\
                                '经营活动现金流入小计',\
                                '购买商品、接受劳务支付的现金',\
                                '客户贷款及垫款净增加额*', '存放中央银行和同业款项净增加额*', '支付原保险合同赔付款项的现金*',\
                                '拆出资金净增加额*', '支付利息、手续费及佣金的现金*', '支付保单红利的现金*',\
                                '支付给职工及为职工支付的现金',\
                                '支付的各项税费',\
                                '支付其他与经营活动有关的现金',\
                                '经营活动现金流出小计',\
                                '经营活动产生的现金流量净额',\
                                '收回投资收到的现金',\
                                '取得投资收益收到的现金',\
                                '处置固定资产、无形资产和其他长期资产收回的现金净额',\
                                '处置子公司及其他营业单位收到的现金净额',\
                                '收到其他与投资活动有关的现金',\
                                '投资活动现金流入小计',\
                                '购建固定资产、无形资产和其他长期资产支付的现金',\
                                '投资支付的现金',\
                                '质押贷款净增加额*',\
                                '取得子公司及其他营业单位支付的现金净额',\
                                '支付其他与投资活动有关的现金',\
                                '投资活动现金流出小计',\
                                '投资活动产生的现金流量净额',\
                                '吸收投资收到的现金', '吸收投资收到的现金：子公司吸收少数股东投资收到的现金',\
                                '取得借款收到的现金',\
                                '收到其他与筹资活动有关的现金',\
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

    cash_flow_statement_dict = {key: [0, 0] for key in cash_flow_statement_list}

    cash_flow_statement_regex_dict = {'销售商品、提供劳务收到的现金': re.compile(r'(?:^|\s+)销售商品、提供劳务收到的现金(?:\s+|$)'),\
                                      '客户存款和同业存放款项净增加额*': re.compile(r'(?:^|\s+)客户存款和同业存放款项净增加额(?:\s+|$)'),\
                                      '向中央银行借款净增加额*': re.compile(r'(?:^|\s+)向中央银行借款净增加额(?:\s+|$)'),\
                                      '向其他金融机构拆入资金净增加额*': re.compile(r'(?:^|\s+)向其他金融机构拆入资金净增加额(?:\s+|$)'),\
                                      '收到原保险合同保费取得的现金*': re.compile(r'(?:^|\s+)收到原保险合同保费取得的现金(?:\s+|$)'),\
                                      '收到再保业务现金净额*': re.compile(r'(?:^|\s+)收到再保业务现金净额(?:\s+|$)'),\
                                      '保户储金及投资款净增加额*': re.compile(r'(?:^|\s+)保户储金及投资款净增加额(?:\s+|$)'),\
                                      '收取利息、手续费及佣金的现金*': re.compile(r'(?:^|\s+)收取利息、手续费及佣金的现金(?:\s+|$)'),\
                                      '拆入资金净增加额*': re.compile(r'(?:^|\s+)拆入资金净增加额(?:\s+|$)'),\
                                      '回购业务资金净增加额*': re.compile(r'(?:^|\s+)回购业务资金净增加额(?:\s+|$)'),\
                                      '代理买卖证券收到的现金净额*': re.compile(r'(?:^|\s+)代理买卖证券收到的现金净额(?:\s+|$)'),\
                                      '收到的税费返还': re.compile(r'(?:^|\s+)收到的税费返还(?:\s+|$)'),\
                                      '收到其他与经营活动有关的现金': re.compile(r'(?:^|\s+)收到其他与经营活动有关的现金(?:\s+|$)'),\
                                      '经营活动现金流入小计': re.compile(r'(?:^|\s+)经营活动现金流入小计(?:\s+|$)'),\
                                      '购买商品、接受劳务支付的现金': re.compile(r'(?:^|\s+)购买商品、接受劳务支付的现金(?:\s+|$)'),\
                                      '客户贷款及垫款净增加额*': re.compile(r'(?:^|\s+)客户贷款及垫款净增加额(?:\s+|$)'),\
                                      '存放中央银行和同业款项净增加额*': re.compile(r'(?:^|\s+)存放中央银行和同业款项净增加额(?:\s+|$)'),\
                                      '支付原保险合同赔付款项的现金*': re.compile(r'(?:^|\s+)支付原保险合同赔付款项的现金(?:\s+|$)'),\
                                      '拆出资金净增加额*': re.compile(r'(?:^|\s+)拆出资金净增加额(?:\s+|$)'),\
                                      '支付利息、手续费及佣金的现金*': re.compile(r'(?:^|\s+)支付利息、手续费及佣金的现金(?:\s+|$)'),\
                                      '支付保单红利的现金*': re.compile(r'(?:^|\s+)支付保单红利的现金(?:\s+|$)'),\
                                      '支付给职工及为职工支付的现金': re.compile(r'(?:^|\s+)支付给职工及为职工支付的现金(?:\s+|$)'),\
                                      '支付的各项税费': re.compile(r'(?:^|\s+)支付的各项税费(?:\s+|$)'),\
                                      '支付其他与经营活动有关的现金': re.compile(r'(?:^|\s+)支付其他与经营活动有关的现金(?:\s+|$)'),\
                                      '经营活动现金流出小计': re.compile(r'(?:^|\s+)经营活动现金流出小计(?:\s+|$)'),\
                                      '经营活动产生的现金流量净额': re.compile(r'(?:^|\s+)经营活动产生的现金流量净额(?:\s+|$)'),\
                                      '收回投资收到的现金': re.compile(r'(?:^|\s+)收回投资收到的现金(?:\s+|$)'),\
                                      '取得投资收益收到的现金': re.compile(r'(?:^|\s+)取得投资收益收到的现金(?:\s+|$)'),\
                                      '处置固定资产、无形资产和其他长期资产收回的现金净额': re.compile(r'(?:^|\s+)处置固定资产、无形资产和其他长期资产收回的现金净额(?:\s+|$)'),\
                                      '处置子公司及其他营业单位收到的现金净额': re.compile(r'(?:^|\s+)处置子公司及其他营业单位收到的现金净额(?:\s+|$)'),\
                                      '收到其他与投资活动有关的现金': re.compile(r'(?:^|\s+)收到其他与投资活动有关的现金(?:\s+|$)'),\
                                      '投资活动现金流入小计': re.compile(r'(?:^|\s+)投资活动现金流入小计(?:\s+|$)'),\
                                      '购建固定资产、无形资产和其他长期资产支付的现金': re.compile(r'(?:^|\s+)购建固定资产、无形资产和其他长期资产支付的现金(?:\s+|$)'),\
                                      '投资支付的现金': re.compile(r'(?:^|\s+)投资支付的现金(?:\s+|$)'),\
                                      '质押贷款净增加额*': re.compile(r'(?:^|\s+)质押贷款净增加额(?:\s+|$)'),\
                                      '取得子公司及其他营业单位支付的现金净额': re.compile(r'(?:^|\s+)取得子公司及其他营业单位支付的现金净额(?:\s+|$)'),\
                                      '支付其他与投资活动有关的现金': re.compile(r'(?:^|\s+)支付其他与投资活动有关的现金(?:\s+|$)'),\
                                      '投资活动现金流出小计': re.compile(r'(?:^|\s+)投资活动现金流出小计(?:\s+|$)'),\
                                      '投资活动产生的现金流量净额': re.compile(r'(?:^|\s+)投资活动产生的现金流量净额(?:\s+|$)'),\
                                      '吸收投资收到的现金': re.compile(r'(?:^|\s+)吸收投资收到的现金(?:\s+|$)'),\
                                      '吸收投资收到的现金：子公司吸收少数股东投资收到的现金': re.compile(r'(?:^|\s+)(\S?其中\S{1}子公司吸收少数股东投资收到的现金\S?)(?:\s+|$)'),\
                                      '取得借款收到的现金': re.compile(r'(?:^|\s+)取得借款收到的现金(?:\s+|$)'),\
                                      '收到其他与筹资活动有关的现金': re.compile(r'(?:^|\s+)收到其他与筹资活动有关的现金(?:\s+|$)'),\
                                      '筹资活动现金流入小计': re.compile(r'(?:^|\s+)筹资活动现金流入小计(?:\s+|$)'),\
                                      '偿还债务支付的现金': re.compile(r'(?:^|\s+)偿还债务支付的现金(?:\s+|$)'),\
                                      '分配股利、利润或偿付利息支付的现金': re.compile(r'(?:^|\s+)分配股利、利润或偿付利息支付的现金(?:\s+|$)'),\
                                      '分配股利、利润或偿付利息支付的现金：子公司支付给少数股东的股利、利润': re.compile(r'(?:^|\s+)(\S?其中\S{1}子公司支付给少数股东的股利、利润\S?)(?:\s+|$)'),\
                                      '支付其他与筹资活动有关的现金': re.compile(r'(?:^|\s+)支付其他与筹资活动有关的现金(?:\s+|$)'),\
                                      '筹资活动现金流出小计': re.compile(r'(?:^|\s+)筹资活动现金流出小计(?:\s+|$)'),\
                                      '筹资活动产生的现金流量净额': re.compile(r'(?:^|\s+)筹资活动产生的现金流量净额(?:\s+|$)'),\
                                      '四、汇率变动对现金及现金等价物的影响': re.compile(r'(?:^|\s+)四、汇率变动对现金及现金等价物的影响(?:\s+|$)'),\
                                      '五、现金及现金等价物净增加额': re.compile(r'(?:^|\s+)五、现金及现金等价物净增加额(?:\s+|$)'),\
                                      '期初现金及现金等价物余额': re.compile(r'(?:^|\s+)期初现金及现金等价物余额(?:\s+|$)'),\
                                      '六、期末现金及现金等价物余额': re.compile(r'(?:^|\s+)六、期末现金及现金等价物余额(?:\s+|$)')}

    num = 0
    display_cash_flow_dict = {}

    folder_path = request.get('data', {}).get('project_folder', '')
    file_path = request.get('data', {}).get('file_path', '')
    sheet_name = request.get('data', {}).get('sheet_name', '')
    items_name = int(request.get('data', {}).get('items_name', ''))
    items_this = int(request.get('data', {}).get('items_this', ''))
    items_previous = int(request.get('data', {}).get('items_previous', ''))

    file_extension = os.path.splitext(file_path)[1].lower()

    if file_extension == '.xlsx':
        import_cash_flow_statement_df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='openpyxl')
    elif file_extension == '.xls':
        import_cash_flow_statement_df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='xlrd')
    else:
        return

    data_cash_flow_statement_df = import_cash_flow_statement_df.iloc[:, [items_name, items_previous, items_this]].copy()
    data_cash_flow_statement_df.columns = ['item', 'previous_year', 'current_year']

    # 数据清洗：将数值列的字符串数据类型转换为浮点数类型
    data_cash_flow_statement_df['previous_year'] = data_cash_flow_statement_df['previous_year'].apply(convert_to_numeric)
    data_cash_flow_statement_df['current_year'] = data_cash_flow_statement_df['current_year'].apply(convert_to_numeric)

    for i in range(1, len(data_cash_flow_statement_df)):
        keywords = data_cash_flow_statement_df.iloc[i, 0]
        if pd.isna(keywords):                           # 处理 NaN/None
            keywords = ''
        elif isinstance(keywords, (float, int)):        # 处理 float/int
            keywords = str(keywords)

        for key in cash_flow_statement_list:
            if re.search(cash_flow_statement_regex_dict[key], keywords):
                previous = data_cash_flow_statement_df.iloc[i, 1]
                previous = 0 if pd.isna(previous) else previous
                end = data_cash_flow_statement_df.iloc[i, 2]
                end = 0 if pd.isna(end) else end
                cash_flow_statement_dict[key][0] = previous
                cash_flow_statement_dict[key][1] = end
                if previous != 0 or end != 0:
                    if pd.notna(previous) or pd.notna(end):
                        num += 1
                        display_cash_flow_dict[num] = [key, end, previous]
                        break

        else:
            key = '@未识别报表项目：' + keywords
            previous = data_cash_flow_statement_df.iloc[i, 1]
            previous = 0 if pd.isna(previous) else previous
            end = data_cash_flow_statement_df.iloc[i, 2]
            end = 0 if pd.isna(end) else end
            if previous != 0 or end != 0:
                if pd.notna(previous) or pd.notna(end):
                    num += 1
                    display_cash_flow_dict[num] = [key, end, previous]

    # 转换为DataFrame
    cash_flow_statement_data = pd.DataFrame(cash_flow_statement_dict).T

    # 确定保存路径
    export_path = os.path.join(folder_path, '项目数据', '现金流量表.xlsx')

    # 保存为xlsx文件
    cash_flow_statement_data.to_excel(export_path, sheet_name='Sheet1', engine='openpyxl')

    # 返回数据字典和保存路径
    return ['import_cash_flow_statement', display_cash_flow_dict]


def convert_to_numeric(value):
    
    if pd.isna(value):
        return 0.0

    try:
        return pd.to_numeric(str(value).strip().replace(',', '').replace(' ', ''))
    except:
        return 0.0