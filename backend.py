# backend.py
import sys
import json
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
    else:
        result = "Unknown command"

    # 返回結果
    response = {"result": result}
    print(json.dumps(response))
    sys.stdout.flush()

if __name__ == "__main__":
    main()