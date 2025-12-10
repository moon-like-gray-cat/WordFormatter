import win32com.client as win32
from docx import Document

def expand_numbering(input_path, output_path):
    word = win32.Dispatch("Word.Application")
    word.Visible = False

    doc = word.Documents.Open(input_path)
    doc.ConvertNumbersToText()

    doc.SaveAs(output_path)
    doc.Close()
    word.Quit()

    return output_path


# ---------------------- 主流程 ----------------------
input_path = r"D:\code\WordFormatter\tests\test1.docx"
output_path = r"D:\code\WordFormatter\tests\expanded.docx"

expanded = expand_numbering(input_path, output_path)

print("生成文件：", expanded)

doc = Document(expanded)

for i, p in enumerate(doc.paragraphs):
    print(f"段落 {i}: {p.text}")
