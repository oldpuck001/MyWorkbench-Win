# backend.py

import sys
import json
from filename_tools import modifythefilename_py
from filename_tools import character_py
from filename_tools import image_py
from filename_tools import sort_py
from filename_tools import export_py
from xlsx_xls_tools import splice_py
from xlsx_xls_tools import subtotals_py
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
        result = modifythefilename_py.modify(request)
    elif request["command"] == "filename_character":
        result = character_py.character(request)
    elif request["command"] == "filename_image":
        result = image_py.image(request)
    elif request["command"] == "filename_sort":
        result = sort_py.sort(request)
    elif request["command"] == "filename_export":
        result = export_py.export(request)

    elif request["command"] == "splice_sheet_input":
        result = splice_py.input_sheet(request)
    elif request["command"] == "splice_sheet_output":
        result = splice_py.output_sheet(request)

    elif request["command"] == "subtotals_import":
        result = subtotals_py.subtotals_import(request)
    elif request["command"] == "subtotals_index":
        result = subtotals_py.subtotals_index(request)
    elif request["command"] == "subtotals_generate":
        result = subtotals_py.subtotals_generate(request)

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