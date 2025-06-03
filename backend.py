# backend.py

import sys
import json
from file_tools import modifythefilename
from file_tools import character
from file_tools import image
from file_tools import export
from file_tools import sort
from file_tools import collect_file
from file_tools import copy_folder

from xlsx_tools import splice
from xlsx_tools import subtotals
from xlsx_tools import regex

from audit_tools import select_folder
from audit_tools import set_up
from audit_tools import import_account_balance_sheet
from audit_tools import import_chronological_account
from audit_tools import import_balance_sheet
from audit_tools import import_income_statement
from audit_tools import import_cash_flow_statement


from data_analysis_tools import data_cleaning
from data_analysis_tools import sql_sqlite
from data_analysis_tools import fill
from data_analysis_tools import bank_statement_sort
from data_analysis_tools import generate_chronological_account

from other_tools import find_subset
from other_tools import text_comparison
from other_tools import docx_comparison
from other_tools import xlsx_comparison

# windows下解决编码问题的语句
import io
sys.stdin = io.TextIOWrapper(sys.stdin.buffer, encoding='utf-8')

def main():

    # 從標準輸入讀取數據
    input_data = sys.stdin.read()

    # 將數據轉換為 Python 對象
    request = json.loads(input_data)

    # 處理數據（這裡我們假設收到的是計算請求）
    if request["command"] == "filename_modify":
        result = modifythefilename.modify(request)
    elif request["command"] == "filename_character":
        result = character.character(request)
    elif request["command"] == "filename_image":
        result = image.image(request)
    elif request["command"] == "filename_export":
        result = export.export(request)
    elif request["command"] == "filename_sort":
        result = sort.sort(request)
    elif request["command"] == "collect_file":
        result = collect_file.collect_file(request)
    elif request["command"] == "copy_folder":
        result = copy_folder.copy_folder(request)


    elif request["command"] == "splice_sheet_input":
        result = splice.input_sheet(request)
    elif request["command"] == "splice_sheet_output":
        result = splice.output_sheet(request)
    elif request["command"] == "subtotals_import":
        result = subtotals.subtotals_import(request)
    elif request["command"] == "subtotals_index":
        result = subtotals.subtotals_index(request)
    elif request["command"] == "subtotals_generate":
        result = subtotals.subtotals_generate(request)
    elif request['command'] == "regex_import":
        result = regex.select_file(request)
    elif request['command'] == "regex_index":
        result = regex.road_sheet(request)
    elif request['command'] == "regex_generate":
        result = regex.regex_generate(request)


    elif request["command"] == "select_folder_path":
        result = select_folder.select_folder_path(request)

    elif request["command"] == "import_config":
        result = set_up.import_config(request)
    elif request['command'] == 'select_basic_file':
        result =set_up.select_basic_file(request)
    elif request["command"] == "import_basic":
        result = set_up.import_basic(request)
    elif request["command"] == "save_settings":
        result = set_up.save_settings(request)

    elif request["command"] == "import_account_balance_sheet":
        result = import_account_balance_sheet.import_account_balance_sheet(request)
    elif request["command"] == "index_account_balance_sheet":
        result = import_account_balance_sheet.index_account_balance_sheet(request)
    elif request["command"] == "export_account_balance_sheet":
        result = import_account_balance_sheet.export_account_balance_sheet(request)

    elif request["command"] == "import_chronological_account":
        result = import_chronological_account.import_chronological_account(request)
    elif request["command"] == "index_chronological_account":
        result = import_chronological_account.index_chronological_account(request)
    elif request["command"] == "export_chronological_account":
        result = import_chronological_account.export_chronological_account(request)

    elif request["command"] == "select_balance_sheet":
        result = import_balance_sheet.select_balance_sheet(request)
    elif request["command"] == "import_balance_sheet":
        result = import_balance_sheet.import_balance_sheet(request)

    elif request["command"] == "select_income_statement":
        result = import_income_statement.select_income_statement(request)
    elif request["command"] == "import_income_statement":
        result = import_income_statement.import_income_statement(request)

    elif request["command"] == "select_cash_flow_statement":
        result = import_cash_flow_statement.select_cash_flow_statement(request)
    elif request["command"] == "import_cash_flow_statement":
        result = import_cash_flow_statement.import_cash_flow_statement(request)






    elif request["command"] == "data_cleaning_import":
        result = data_cleaning.data_cleaning_import(request)
    elif request["command"] == "data_cleaning_index":
        result = data_cleaning.data_cleaning_index(request)
    elif request["command"] == "data_cleaning_export":
        result = data_cleaning.data_cleaning_export(request)
    elif request["command"] == "sql_sqlite_folder":
        result = sql_sqlite.sql_sqlite_folder(request)
    elif request["command"] == "sql_sqlite_sql":
        result = sql_sqlite.sql_sqlite_sql(request)
    elif request["command"] == "sql_sqlite_backup":
        result = sql_sqlite.sql_sqlite_backup(request)
    elif request["command"] == "sql_sqlite_select":
        result = sql_sqlite.sql_sqlite_select(request)
    elif request['command'] == "fill_import":
        result = fill.select_file(request)
    elif request['command'] == "fill_index":
        result = fill.road_sheet(request)
    elif request['command'] == "fill_generate":
        result = fill.fill_generate(request)
    elif request["command"] == "bank_statement_sort_import":
        result = bank_statement_sort.bank_statement_sort_import(request)
    elif request["command"] == "bank_statement_sort_index":
        result = bank_statement_sort.bank_statement_sort_index(request)
    elif request["command"] == "bank_statement_sort_debit_or_credit":
        result = bank_statement_sort.bank_statement_sort_debit_or_credit(request)
    elif request["command"] == "bank_statement_sort_export":
        result = bank_statement_sort.bank_statement_sort_export(request)
    elif request["command"] == "generate_chronological_account_import":
        result = generate_chronological_account.generate_chronological_account_import(request)
    elif request["command"] == "generate_chronological_account_index":
        result = generate_chronological_account.generate_chronological_account_index(request)
    elif request["command"] == "generate_chronological_account_debit_or_credit":
        result = generate_chronological_account.generate_chronological_account_debit_or_credit(request)
    elif request["command"] == "generate_chronological_account_export":
        result = generate_chronological_account.generate_chronological_account_export(request)

    elif request["command"] == "find_subset_sheetnames_import":
        result = find_subset.find_subset_sheetnames_import(request)
    elif request["command"] == "find_subset_columns_index":
        result = find_subset.find_subset_columns_index(request)
    elif request["command"] == "find_subset_import":
        result = find_subset.find_subset_import(request)
    elif request["command"] == "find_subset_export":
        result = find_subset.find_subset_export(request)

    elif request["command"] == "text_comparison":
        result = text_comparison.text_comparison(request)
    elif request["command"] == "docx_comparison":
        result = docx_comparison.compare_word_documents(request)
    elif request["command"] == "xlsx_comparision_sheetnames":
        result = xlsx_comparison.xlsx_comparision_sheetnames(request)
    elif request["command"] == "xlsx_comparison":
        result = xlsx_comparison.compare_excels(request)

    else:
        result = "Unknown command"

    # 返回結果
    response = {"result": result}
    print(json.dumps(response))
    sys.stdout.flush()

if __name__ == "__main__":
    main()