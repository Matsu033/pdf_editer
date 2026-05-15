import pdfplumber
from tkinter import Tk, filedialog

# Tkinter のウィンドウを非表示にする
root = Tk()
root.withdraw()

# PDF を選択
pdf_path = filedialog.askopenfilename(
    title="PDFファイルを選択してください",
    filetypes=[("PDF Files", "*.pdf")]
)

if not pdf_path:
    print("PDF が選択されませんでした")
    exit()

# 出力する TXT ファイル名を決める
txt_path = pdf_path.replace(".pdf", ".txt")

# PDF → TXT 変換
with pdfplumber.open(pdf_path) as pdf:
    with open(txt_path, "w", encoding="utf-8") as out:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                out.write(text + "\n")

print("変換完了:", txt_path)
