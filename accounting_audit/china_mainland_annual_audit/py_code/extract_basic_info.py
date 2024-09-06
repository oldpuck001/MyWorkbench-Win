# extract_basic_info.py

import os
import pandas as pd

def extract_basic_info_xlsx(filepath, data_folder):
    # 读取基本信息表格
    basic_info_df = pd.read_excel(filepath, sheet_name='基本信息')
    
    # 提取所需信息
    enterprise_name = basic_info_df.iloc[1, 1]               # 企业名称
    date_of_establishment = basic_info_df.iloc[4, 3]         # 成立日期
    approval_date = basic_info_df.iloc[7, 1]                 # 核准日期
    unified_social_credit_code = basic_info_df.iloc[4, 1]    # 统一社会信用代码
    registered_capital = basic_info_df.iloc[2, 3]            # 注册资本
    legal_representative = basic_info_df.iloc[2, 1]          # 法定代表人
    registered_address = basic_info_df.iloc[13, 1]           # 注册地址
    business_scope = basic_info_df.iloc[15, 1]               # 经营范围
    
    # 创建一个字典来保存这些信息
    info_dict = {
        '企业名称': [enterprise_name],
        '成立日期': [date_of_establishment],
        '核准日期': [approval_date],
        '统一社会信用代码': [unified_social_credit_code],
        '注册资本': [registered_capital],
        '法定代表人': [legal_representative],
        '注册地址': [registered_address],
        '经营范围': [business_scope]
    }
    
    # 转换为DataFrame
    info_df = pd.DataFrame(info_dict)
    
    # 确定保存路径
    save_path = os.path.join(data_folder, 'basic_info.csv')
    
    # 保存为CSV文件
    info_df.to_csv(save_path, index=False, encoding='utf-8-sig')
    
    # 返回数据字典和保存路径
    return info_dict, save_path