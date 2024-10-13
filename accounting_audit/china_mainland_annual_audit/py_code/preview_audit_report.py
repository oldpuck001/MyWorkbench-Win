# audit_report_preview.py

from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

def get_report_docx(report_path):
    # 读取 Word 文档内容
    doc = Document(report_path)    
    # 将文档内容拼接成 HTML 格式的字符串，同时保留对齐方式
    content = ""
    for para in doc.paragraphs:
        align = para.alignment
        if align == WD_PARAGRAPH_ALIGNMENT.LEFT:  # 左对齐
            align_style = "text-align: left;"
        elif align == WD_PARAGRAPH_ALIGNMENT.CENTER:  # 居中对齐
            align_style = "text-align: center;"
        elif align == WD_PARAGRAPH_ALIGNMENT.RIGHT:  # 右对齐
            align_style = "text-align: right;"
        elif align == WD_PARAGRAPH_ALIGNMENT.JUSTIFY:  # 两端对齐
            align_style = "text-align: justify;"
        else:  # 默认左对齐
            align_style = "text-align: left;"
        # 获取段落的字体大小
        font_size = None
        for run in para.runs:
            if run.font.size:
                font_size = run.font.size.pt  # 将尺寸转为pt单位
                break  # 假设段落中所有文字使用相同的字体大小
        font_size_style = f"font-size: {font_size}pt;" if font_size else ""

        # 设置段落缩进（如果有）
        text_indent = para.paragraph_format.first_line_indent.pt if para.paragraph_format.first_line_indent else 0

        # 获取段首的空格
        leading_spaces = para.text[:len(para.text) - len(para.text.lstrip())]
        # 将空格转换为 HTML 中的 `&nbsp;`，每个空格对应一个 `&nbsp;`
        leading_spaces_html = '&nbsp;' * len(leading_spaces)

        # 组合样式
        spacing_style = f"text-indent: {text_indent}pt;"

        # 将段落的样式应用到HTML，并保留段首空格
        content += f"<p style='{align_style} {font_size_style} {spacing_style}'>{leading_spaces_html}{para.text.lstrip()}</p>\n"

    return content