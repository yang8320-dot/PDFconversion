from pypdf import PdfWriter

def process_merge_pdfs(input_files, output_path, status_callback):
    """
    合併多個 PDF 檔案
    :param status_callback: 傳入一個函式，用來即時更新主介面的文字狀態
    """
    status_callback("📑 正在合併 PDF 檔案...")
    merger = PdfWriter()
    for pdf in input_files:
        merger.append(pdf)
    merger.write(output_path)
    merger.close()
