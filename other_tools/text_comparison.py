# text_comparison.py

import difflib

def text_comparison(request):

    text_1 = request.get("data", {}).get("text_1", "")
    text_2 = request.get("data", {}).get("text_2", "")

    d = difflib.Differ()
    diff = d.compare(text_1.splitlines(), text_2.splitlines())

    return ['text_comparison', ['\n'.join(diff)]]