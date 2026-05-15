import pdfplumber
from tkinter import Tk, filedialog
from pptx import Presentation

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

# 出力する PPTX ファイル名
ppt_path = pdf_path.replace(".pdf", ".pptx").replace(".PDF", ".pptx")

# PowerPoint プレゼンを作成
prs = Presentation()

# PDF → PPT
with pdfplumber.open(pdf_path) as pdf:
    for page in pdf.pages:
        text = page.extract_text() or ""

        # スライド追加（タイトル＋本文レイアウト）
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)

        # タイトル（空のまま）
        title = slide.shapes.title
        title.text = ""

        # 本文
        body = slide.placeholders[1]
        body.text = text

# 保存
prs.save(ppt_path)

print("PPT 変換完了:", ppt_path)
