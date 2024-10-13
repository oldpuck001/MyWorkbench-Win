# flask_app.py

import os
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import safe_join
import json
import pandas as pd
from py_code.accounting_audit.china_mainland_annual_audit import extract_basic_info
from py_code.accounting_audit.china_mainland_annual_audit import extract_balance_sheet
from py_code.accounting_audit.china_mainland_annual_audit import extract_income_statement
from py_code.accounting_audit.china_mainland_annual_audit import extract_cash_flow
from py_code.accounting_audit.china_mainland_annual_audit import extract_equity_change
from py_code.accounting_audit.china_mainland_annual_audit import extract_account_balance
from py_code.accounting_audit.china_mainland_annual_audit import extract_chronological_account
from py_code.accounting_audit.china_mainland_annual_audit import generate_audit_report
from py_code.accounting_audit.china_mainland_annual_audit import generate_financial_report
from py_code.accounting_audit.china_mainland_annual_audit import preview_audit_report
from py_code.accounting_audit.china_mainland_annual_audit import preview_balance_sheet
from py_code.accounting_audit.china_mainland_annual_audit import preview_income_statement

app = Flask(__name__, static_folder='static', template_folder='templates')

project_folder = None

CONFIG_FILE = 'config.json'

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as f:
            return json.load(f)
    return {"defaultPath": "/Users/lei/Downloads"}

def save_config(config):
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f)

def handle_file_upload(file_type):
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    if file:
        global project_folder
        data_folder = os.path.join(project_folder, 'audit_support_software_folder', 'data_floder')
        if not os.path.exists(data_folder):
            os.makedirs(data_folder)
        filepath = upload_save(file, file_type)
        if file_type == 'qcc':
            extracted_data = extract_basic_info.extract_basic_info_xlsx(filepath, data_folder)
        elif file_type == 'balanceSheet':
            extracted_data = extract_balance_sheet.extract_balance_sheet_xlsx(filepath, data_folder)
        elif file_type == 'profitSheet':
            extracted_data = extract_income_statement.extract_income_statement_xlsx(filepath, data_folder)
        elif file_type == 'cashFlow':
            extracted_data = extract_cash_flow.extract_cash_flow_xlsx(filepath, data_folder)
        elif file_type == 'equityChange':
            extracted_data = extract_equity_change.extract_equity_change_xlsx(filepath, data_folder)
        elif file_type == 'accountBalance':
            extracted_data = extract_account_balance.extract_account_balance_xlsx(filepath, data_folder)
        elif file_type == 'chronologicalAccount':
            extracted_data = extract_chronological_account.extract_chronological_account_xlsx(filepath, data_folder)
        else:
            return 'Unknown file type'
        return jsonify({'extracted_data': extracted_data})

def upload_save(file, file_type):
            original_filename = file.filename
            filename, ext = os.path.splitext(original_filename)
            if file_type == 'qcc':
                processed_filename = f'basic_info{ext}'
            elif file_type == 'balanceSheet':
                processed_filename = f'balance_sheet{ext}'
            elif file_type == 'profitSheet':
                processed_filename = f'income_statement{ext}'
            elif file_type == 'cashFlow':
                processed_filename = f'cash_flow_statement{ext}'
            elif file_type == 'equityChange':
                processed_filename = f'equity_change_statement{ext}'
            elif file_type == 'accountBalance':
                processed_filename = f'account_balance{ext}'
            elif file_type == 'chronologicalAccount':
                processed_filename = f'chronological_account{ext}'
            global project_folder
            uploads_folder = os.path.join(project_folder, 'audit_support_software_folder', 'uploads')
            if not os.path.exists(uploads_folder):
                os.makedirs(uploads_folder)
            filepath = os.path.join(uploads_folder, processed_filename)
            file.save(filepath)
            return filepath

@app.route('/')
def index():
    return render_template('/index.html')

@app.route('/<dir_name_1>/<path:filename>')
def first_level_directory(dir_name_1, filename):
    # 使用safe_join来确保路径是安全的
    template_path = safe_join(dir_name_1, filename)
    return render_template(template_path)

@app.route('/<dir_name_1>/<dir_name_2>/<path:filename>')
def second_level_directory(dir_name_1, dir_name_2, filename):
    # 使用safe_join来确保路径是安全的
    template_path = safe_join(dir_name_1, dir_name_2, filename)
    return render_template(template_path)

@app.route('/<dir_name_1>/<dir_name_2>/<dir_name_3>/<path:filename>')
def third_level_directory(dir_name_1, dir_name_2, dir_name_3, filename):
    template_path = safe_join(dir_name_1, dir_name_2, dir_name_3, filename)
    return render_template(template_path)

@app.route('/get-default-path', methods=['GET'])
def get_default_path():
    config = load_config()
    return jsonify(config)

@app.route('/your-backend-endpoint', methods=['POST'])
def receive_folder_path():
    data = request.json
    folder_path = data.get('folderPath')
    default_path = data.get('defaultPath')
    if not folder_path or not default_path:
        return jsonify({"error": "未找到项目文件夹路径或项目文件夹所在路径。"}), 400
    global project_folder
    # 将路径转换为绝对路径
    project_folder = os.path.abspath(os.path.join(default_path, folder_path))
    audit_support_software_folder = os.path.join(project_folder, 'audit_support_software_folder')
    if not os.path.exists(audit_support_software_folder):
        os.makedirs(audit_support_software_folder)
    # 更新配置文件中的默认路径
    config = load_config()
    config['defaultPath'] = default_path
    save_config(config)
    return jsonify({"message": "Folder path received successfully", "projectFolder": project_folder, "defaultPath": default_path}), 200

@app.route('/upload-file/qcc', methods=['POST'])
def upload_file_qcc():
    return handle_file_upload('qcc')

@app.route('/upload-file/balanceSheet', methods=['POST'])
def upload_file_balance_sheet():
    return handle_file_upload('balanceSheet')

@app.route('/upload-file/profitSheet', methods=['POST'])
def upload_file_profit_sheet():
    return handle_file_upload('profitSheet')

@app.route('/upload-file/cashFlow', methods=['POST'])
def upload_file_cash_flow():
    return handle_file_upload('cashFlow')

@app.route('/upload-file/equityChange', methods=['POST'])
def upload_file_equity_change():
    return handle_file_upload('equityChange')

@app.route('/upload-file/accountBalance', methods=['POST'])
def upload_file_account_balance():
    return handle_file_upload('accountBalance')

@app.route('/upload-file/chronologicalAccount', methods=['POST'])
def upload_file_chronological_account():
    return handle_file_upload('chronologicalAccount')

@app.route('/save-period-deadline', methods=['POST'])
def save_period_deadline():
    data = request.get_json()
    period_audit = data.get('period_audit')
    deadline_audit = data.get('deadline_audit')
    if not period_audit or not deadline_audit:
        return jsonify({'error': 'No period_audit or deadline_audit provided'}), 400
    global project_folder
    data_folder = os.path.join(project_folder, 'audit_support_software_folder', 'data_floder', 'basic_info.csv')
    df = pd.read_csv(data_folder)
    new_data = {'审计期间': period_audit, '审计截止日': deadline_audit}
    for key, value in new_data.items():
        df.at[df.index[-1], key] = value
    df.to_csv(data_folder, index=False, encoding='utf-8-sig')
    return jsonify({'message': '审计期间与截止日保存成功'})

@app.route('/save-api-key', methods=['POST'])
def save_api_key():
    data = request.get_json()
    api_key = data.get('api_key')
    if not api_key:
        return jsonify({'error': 'No API key provided'}), 400
    config = load_config()
    config['apiKey'] = api_key
    save_config(config)
    return jsonify({'message': 'API Key保存成功'})

@app.route('/generate_audit_report', methods=['POST'])
def generate_audit_report():
    global project_folder
    software_folder = os.path.join(project_folder, 'audit_support_software_folder')
    # 从前端获取审计意见和报告出具日期
    audit_opinion = request.form.get('auditOpinion')
    report_number = request.form.get('reportNumber')
    report_date = request.form.get('reportDate')
    if not audit_opinion or not report_date:
        return "Missing required fields", 400
    try:
        # 假设 generate 方法生成报告并返回生成文件的路径
        output_file_path = generate_audit_report.generate(software_folder, audit_opinion, report_number, report_date)
    except Exception as e:
        return f"An error occurred: {str(e)}", 500
    # 验证文件是否生成成功
    if not os.path.exists(output_file_path):
        return "File not found after generation", 500
    # 生成成功后返回状态 200 和成功消息
    return jsonify({"message": "Report generated successfully"}), 200

@app.route('/generate_report_attachment', methods=['POST'])
def generate_report_attachment():
    global project_folder
    software_folder = os.path.join(project_folder, 'audit_support_software_folder')
    try:
        output_file_path = generate_financial_report.generate(software_folder)
    except Exception as e:
        return f"An error occurred: {str(e)}", 500
    # 验证文件是否生成成功
    if not os.path.exists(output_file_path):
        return "File not found after generation", 500
    # 生成成功后返回状态200和成功消息
    return jsonify({"message": "Attachment generated successfully"}), 200

@app.route('/get_audit_report')
def get_audit_report():
    global project_folder
    report_path = os.path.join(project_folder, 'audit_support_software_folder', 'report', '审计报告正文.docx')
    content = preview_audit_report.get_report_docx(report_path)
    return jsonify({'audit_content': content})

@app.route('/get_balance_sheet')
def get_balance_sheet():
    global project_folder
    balance_path = os.path.join(project_folder, 'audit_support_software_folder', 'data_floder', 'balance_sheet.csv')
    content = preview_balance_sheet.get_balance_sheet(balance_path)
    return jsonify({'balance_content': content})

@app.route('/get_income_statement')
def get_income_statement():
    global project_folder
    income_path = os.path.join(project_folder, 'audit_support_software_folder', 'data_floder', 'income_statement.csv')
    content = preview_income_statement.get_income_statement(income_path)
    return jsonify({'income_content': content})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5003, debug=True)