# backend.py
import sys
import json
from xlsx_xls_tools import splice_py
from xlsx_xls_tools import subtotals_py
from study_language_tools import vocabulary_to_audio_py

def main():

    # 從標準輸入讀取數據
    input_data = sys.stdin.read()

    # 將數據轉換為 Python 對象
    request = json.loads(input_data)

    # 處理數據（這裡我們假設收到的是計算請求）
    if request["command"] == "importButton":
        result = vocabulary_to_audio_py.import_vocabulary(request)
    elif request["command"] == "generateButton":
        result = vocabulary_to_audio_py.generate_audio(request)

    elif request["command"] == "subtotals_import":
        result = subtotals_py.subtotals_import(request)
    elif request["command"] == "subtotals_index":
        result = subtotals_py.subtotals_index(request)
    elif request["command"] == "subtotals_generate":
        result = subtotals_py.subtotals_generate(request)


    elif request["command"] == "splice_sheet_input":
        result = splice_py.input_sheet(request)
    elif request["command"] == "splice_sheet_output":
        result = splice_py.output_sheet(request)


    else:
        result = "Unknown command"

    # 返回結果
    response = {"result": result}
    print(json.dumps(response))
    sys.stdout.flush()

if __name__ == "__main__":
    main()