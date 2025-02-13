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
from xlsx_xls_tools import splice
from xlsx_xls_tools import subtotals
from xlsx_xls_tools import fill
from xlsx_xls_tools import regex
from data_analysis_tools import single_sort_export_py
from data_analysis_tools import bank_statement_sort_py

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
    elif request['command'] == "fill_import":
        result = fill.select_file(request)
    elif request['command'] == "fill_index":
        result = fill.road_sheet(request)
    elif request['command'] == "fill_generate":
        result = fill.fill_generate(request)
    elif request['command'] == "regex_import":
        result = regex.select_file(request)
    elif request['command'] == "regex_index":
        result = regex.road_sheet(request)
    elif request['command'] == "regex_generate":
        result = regex.regex_generate(request)
    elif request["command"] == "single_sort_export_import":
        result = single_sort_export_py.single_sort_export_import(request)
    elif request["command"] == "single_sort_export_index":
        result = single_sort_export_py.single_sort_export_index(request)
    elif request["command"] == "single_sort_export_export":
        result = single_sort_export_py.single_sort_export_export(request)

    elif request["command"] == "bank_statement_sort_import":
        result = bank_statement_sort_py.bank_statement_sort_import(request)
    elif request["command"] == "bank_statement_sort_index":
        result = bank_statement_sort_py.bank_statement_sort_index(request)
    elif request["command"] == "bank_statement_sort_debit_or_credit":
        result = bank_statement_sort_py.bank_statement_sort_debit_or_credit(request)
    elif request["command"] == "bank_statement_sort_export":
        result = bank_statement_sort_py.bank_statement_sort_export(request)

    else:
        result = "Unknown command"

    # 返回結果
    response = {"result": result}
    print(json.dumps(response))
    sys.stdout.flush()

if __name__ == "__main__":
    main()