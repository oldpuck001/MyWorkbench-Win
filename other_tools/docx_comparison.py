# docx_comparison.py

from docx import Document

def get_paragraphs(file_path):

    doc = Document(file_path)
    return [para.text.strip() for para in doc.paragraphs if para.text.strip()]

def compare_word_documents(request):

    docx_path_1 = request.get("data", {}).get("docx_path_1", "")
    docx_path_2 = request.get("data", {}).get("docx_path_2", "")

    paras1 = get_paragraphs(docx_path_1)
    paras2 = get_paragraphs(docx_path_2)

    max_len = max(len(paras1), len(paras2))
    differences = []

    for i in range(max_len):
        text1 = paras1[i] if i < len(paras1) else ''
        text2 = paras2[i] if i < len(paras2) else ''
        if text1 != text2:
            differences.append({
                'Paragraph': i + 1,
                'File1': text1,
                'File2': text2
            })

    if differences:

        text = f'发现 {len(differences)} 处段落不同：\n\n'

        for diff in differences:

            text += f'第 {diff['Paragraph']} 段：\n'
            text += f'  File1: {diff['File1']}\n'
            text += f'  File2: {diff['File2']}\n'
            text += '\n'

        result_text = {'result_message': text}
        return ['docx_comparison', result_text]

    else:

        result_text = {'result_message': '两个Word文档内容一致。'}
        return ['docx_comparison', result_text]