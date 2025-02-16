# sql_sqlite.py

import os
import shutil
import sqlite3
import pandas as pd

def sql_sqlite_folder(request):

    folder_path = request.get("data", {}).get("folder_path", "")
    file_path = os.path.join(folder_path, 'sqlite_database.db')

    # 检测文件路径是否存在
    if os.path.exists(file_path):
        result_text = '连接数据库成功！\n'
    else:
        result_text = '创建数据库成功！\n'

    conn = sqlite3.connect(file_path)
    conn.close()

    return ['sql_sqlite_folder', [result_text]]


def sql_sqlite_sql(request):

    folder_path = request.get("data", {}).get("folder_path", "")
    file_path = os.path.join(folder_path, 'sqlite_database.db')
    sql_command = request.get("data", {}).get("sql_command", "")

    # 连接数据库
    conn = sqlite3.connect(file_path)

    # 从连接获取游标
    curs = conn.cursor()

    # 执行数据库指令
    curs.execute(sql_command)

    # 提交所做的修改，將其保存到文件中
    conn.commit()

    # 关闭游标
    curs.close()

    # 关闭数据库连接
    conn.close()

    result_text = f'指令：\n{sql_command}\n执行完毕！\n'

    return ['sql_sqlite_sql', [result_text]]

def sql_sqlite_backup(request):

    folder_path = request.get("data", {}).get("folder_path", "")
    file_path = os.path.join(folder_path, 'sqlite_database.db')

    target_path = os.path.join(folder_path, 'sqlite_database_1.db')

    counter = 2
    while os.path.exists(target_path):
        new_filename = f'sqlite_database_{counter}.xlsx'
        target_path = os.path.join(folder_path, new_filename)
        counter += 1

    shutil.copy(file_path, target_path)

    result_text = f'还原点建立成功，文件路径：{target_path}\n'

    return ['sql_sqlite_backup', [result_text]]


def sql_sqlite_select(request):

    folder_path = request.get("data", {}).get("folder_path", "")
    file_path = os.path.join(folder_path, 'sqlite_database.db')
    sql_command = request.get("data", {}).get("sql_command", "")
    save_path = request.get("data", {}).get("save_path", "")

    # 连接数据库
    conn = sqlite3.connect(file_path)

    # 使用 pandas 读取 SQL 查询结果
    try:
        # 将查询结果直接读取为 DataFrame
        df = pd.read_sql_query(sql_command, conn)

        # 将 DataFrame 导出为 Excel 文件
        df.to_excel(save_path, index=False, engine='openpyxl')

        result_text = f'查询指令：\n{sql_command}\n执行完毕！\n导出文件路径：{save_path}\n'

    except Exception as e:
        
        result_text = f'查询指令：\n{sql_command}\n执行失败！\n错误信息：{str(e)}\n'

    finally:

        conn.close()

    return ['sql_sqlite_select', [result_text]]