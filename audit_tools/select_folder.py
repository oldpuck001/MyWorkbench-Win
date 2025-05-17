# select_folder.py

import os
import json

def select_folder_path(request):

    folder_path = request.get('data', {}).get('project_folder', '')

    settings_path = os.path.join(folder_path, '项目数据', 'settings.json')

    # 检查路径是否存在
    if os.path.exists(settings_path):

        # 獲取設定值
        result_text = {
            'result_message': f'选择项目文件夹{folder_path}成功！\n'
        }

    else:
        
        # 创建目录及文件
        os.makedirs(os.path.join(folder_path, '项目数据'), exist_ok=True)
        os.makedirs(os.path.join(folder_path, '审计底稿'), exist_ok=True)
        os.makedirs(os.path.join(folder_path, '审计底稿', '1.初步业务活动'), exist_ok=True)
        os.makedirs(os.path.join(folder_path, '审计底稿', '2.风险评估'), exist_ok=True)
        os.makedirs(os.path.join(folder_path, '审计底稿', '3.控制测试'), exist_ok=True)
        os.makedirs(os.path.join(folder_path, '审计底稿', '4.实质性程序'), exist_ok=True)
        os.makedirs(os.path.join(folder_path, '审计底稿', '5.其他项目'), exist_ok=True)
        os.makedirs(os.path.join(folder_path, '审计底稿', '6.完成审计工作'), exist_ok=True)
        os.makedirs(os.path.join(folder_path, '审计底稿', '7.永久性档案'), exist_ok=True)
        os.makedirs(os.path.join(folder_path, '审计底稿', '8.底稿附件'), exist_ok=True)
        os.makedirs(os.path.join(folder_path, '审计底稿', '8.底稿附件', '1.资产类资料'), exist_ok=True)
        os.makedirs(os.path.join(folder_path, '审计底稿', '8.底稿附件', '2.负债类资料'), exist_ok=True)
        os.makedirs(os.path.join(folder_path, '审计底稿', '8.底稿附件', '3.权益类资料'), exist_ok=True)
        os.makedirs(os.path.join(folder_path, '审计底稿', '8.底稿附件', '4.损益类资料'), exist_ok=True)
        os.makedirs(os.path.join(folder_path, '审计底稿', '9.记账凭证检查拍照'), exist_ok=True)
        os.makedirs(os.path.join(folder_path, '审计报告'), exist_ok=True)
        os.makedirs(os.path.join(folder_path, '原始资料'), exist_ok=True)

        settings_dict = {'被审计会计期间': '',
                         '被审计会计报表截止日': '',
                         '会计师事务所名称': '',
                         '保护单元格工作表密码': '',
                         '企业名称': '',
                         '成立日期': '',
                         '核准日期': '',
                         '统一社会信用代码': '',
                         '注册资本': '',
                         '法定代表人': '',
                         '注册地址': '',
                         '经营范围': ''
                          }

        # 使用 UTF-8 编码写入 JSON 文件
        with open(settings_path, 'w', encoding='utf-8') as f:
            json.dump(settings_dict, f, indent=4)

        result_text = {
            'result_message': f'创建项目文件夹{folder_path}的子文件夹成功！\n创建项目文件夹{folder_path}的项目信息配置文件成功！\n初始化项目文件夹{folder_path}成功！\n'
        }

    return ['select_folder_path', result_text]